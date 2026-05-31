# -*- coding: utf-8 -*-
"""
国家赔偿（错误执行）适配器
适用于刑事赔偿中的错误执行赔偿申请

模板：国家赔偿申请书-错误执行赔偿.docx
模板结构（3表格）：
- T0 (4行): 赔偿请求人 - 自然人行(Row2)/法人行(Row3)
- T1 (8行): 委托代理人(Row0) + 赔偿义务机关(Row1) + 自赔决定(Row2) 
            + 执行标的(Row3) + 赔偿请求标题(Row4-Row5) + 财产权赔偿(Row6) 
            + 停产停业损失(Row7)
- T2 (8行): 人身自由赔偿(Row0) + 其他赔偿(Row1) + 事实与理由标题(Row2-Row3) 
            + 法律依据(Row4) + 其他说明(Row5) + 同类案件(Row6) + 证据清单(Row7)
"""

import os
import sys
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional

# 添加项目根目录到路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, os.path.join(BASE_DIR, 'scripts'))

# 注册命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class StateCompensationExecutionAdapter:
    """
    国家赔偿（错误执行）适配器
    
    适用于错误执行赔偿申请，包括：
    - 财产权赔偿（返还原物/恢复原状/赔偿损失）
    - 停产停业损失赔偿
    - 侵犯人身自由赔偿金
    """
    
    TEMPLATE_FILENAME = "国家赔偿申请书-错误执行赔偿.docx"
    
    def __init__(self, data: Dict[str, Any]):
        """
        初始化适配器
        """
        self.data = data
        self._unpacked_dir = None
        self._root = None
        self._xml_decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    
    def _get_template_path(self) -> str:
        """获取模板文件路径"""
        return os.path.join(BASE_DIR, "template_cache", self.TEMPLATE_FILENAME)
    
    def _get_output_dir(self) -> str:
        """获取输出目录"""
        return os.path.join(BASE_DIR, "output")
    
    def _unpack(self):
        """解压模板"""
        self._unpacked_dir = self._get_template_path().replace('.docx', '_work')
        if os.path.exists(self._unpacked_dir):
            shutil.rmtree(self._unpacked_dir)
        with zipfile.ZipFile(self._get_template_path(), 'r') as zf:
            zf.extractall(self._unpacked_dir)
    
    def _parse_xml(self):
        """解析 document.xml"""
        doc_path = os.path.join(self._unpacked_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            content = f.read()
        self._root = ET.fromstring(content)
    
    def _save_and_pack(self, output_path: str) -> str:
        """保存并打包"""
        doc_path = os.path.join(self._unpacked_dir, 'word', 'document.xml')
        content = ET.tostring(self._root, encoding='unicode')
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(self._xml_decl + content)
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(self._unpacked_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, self._unpacked_dir)
                    zf.write(file_path, arc_name)
        
        return output_path
    
    def _cleanup(self):
        """清理临时目录"""
        if self._unpacked_dir and os.path.exists(self._unpacked_dir):
            shutil.rmtree(self._unpacked_dir)
    
    def _get_para_text(self, p) -> str:
        """获取段落文本"""
        return ''.join(t.text or '' for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
    
    def _fill_field_in_cell(self, cell, label: str, value: str):
        """
        在单元格内填充字段
        通过在标签文本后追加值来填充
        """
        if not value:
            return
        
        W_NS = NAMESPACES['w']
        for p in cell.iter(f'{{{W_NS}}}p'):
            full_text = self._get_para_text(p)
            if label not in full_text:
                continue
            
            # 找到包含标签的run
            for r in p.iter(f'{{{W_NS}}}r'):
                for t in r.iter(f'{{{W_NS}}}t'):
                    if t.text and label in t.text:
                        idx = t.text.index(label) + len(label)
                        # 检查标签后的内容
                        after = t.text[idx:]
                        if not after.strip():
                            # 标签后是空白，直接追加
                            t.text = t.text + value
                            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            return
                        elif re.match(r'^\s+$', after):
                            # 标签后是空白符，替换为值
                            t.text = t.text[:idx] + value
                            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            return
            
            # 如果标签在不同的run中，查找标签后的空run
            runs = list(p.iter(f'{{{W_NS}}}r'))
            label_run_idx = None
            for ri, r in enumerate(runs):
                for t in r.iter(f'{{{W_NS}}}t'):
                    if t.text and label in t.text:
                        label_run_idx = ri
                        break
                if label_run_idx is not None:
                    break
            
            if label_run_idx is not None:
                # 在标签run后找空run
                for next_r in runs[label_run_idx + 1:]:
                    for t in next_r.iter(f'{{{W_NS}}}t'):
                        if not t.text or not t.text.strip():
                            t.text = value
                            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            return
                
                # 没有空run，在标签文本后追加
                for t in runs[label_run_idx].iter(f'{{{W_NS}}}t'):
                    if t.text and label in t.text:
                        idx = t.text.index(label) + len(label)
                        t.text = t.text[:idx] + value
                        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        return
    
    def _fill_checkbox_in_cell(self, cell, before_text: str, checked: bool):
        """
        在单元格内勾选框
        
        before_text 格式：
        - "性别：男" - 标签+选项
        - "有□" - 独立勾选框（当勾选时，整个文本变为"有☑"）
        - "是" - 是/否选项（配合独立的□勾选框）
        """
        if not checked:
            return
        
        W_NS = NAMESPACES['w']
        
        # 收集单元格中所有段落
        paragraphs = list(cell.iter(f'{{{W_NS}}}p'))
        
        for p in paragraphs:
            full_text = self._get_para_text(p)
            
            if before_text not in full_text:
                continue
            
            # 收集段落中所有 <w:t> 元素
            t_elems = []
            for r in p.iter(f'{{{W_NS}}}r'):
                for t in r.iter(f'{{{W_NS}}}t'):
                    if t.text:
                        t_elems.append(t)
            
            # 查找 before_text 在 t_elems 中的位置
            text_pos = 0
            before_text_start = -1
            before_text_end = -1
            
            for i, t in enumerate(t_elems):
                t_start = text_pos
                t_end = text_pos + len(t.text)
                
                if before_text_start < 0:
                    # 查找 before_text 开始位置
                    idx = t.text.find(before_text)
                    if idx >= 0:
                        before_text_start = t_start + idx
                        before_text_end = before_text_start + len(before_text)
                        before_t_idx = i
                        before_t_offset = idx
                
                text_pos = t_end
            
            if before_text_start < 0:
                continue
            
            # 情况0: before_text 本身就是一个完整的勾选框文本（如 "特别授权□"）
            # 直接替换整个文本中的 □
            if '□' in before_text:
                # 在段落中查找并替换
                full_text_new = full_text.replace(before_text, before_text.replace('□', '☑'))
                # 找到第一个匹配的 t 并修改
                for t in t_elems:
                    if before_text in t.text:
                        t.text = t.text.replace(before_text, before_text.replace('□', '☑'))
                        return
                continue
            
            # 情况1: before_text 本身就是一个独立勾选框（如 "有□"）
            # 检查 before_text 是否就是整个 t.text
            if len(t_elems) > before_t_idx:
                t = t_elems[before_t_idx]
                if t.text == before_text and before_t_offset == 0:
                    # before_text 正好是这个 t 的全部内容
                    if '□' in t.text:
                        t.text = t.text.replace('□', '☑')
                        return
                    elif '☐' in t.text:
                        t.text = t.text.replace('☐', '☑')
                        return
            
            # 情况2: before_text + □ 在同一个 t 中
            for t in t_elems:
                if before_text in t.text:
                    after = t.text[t.text.index(before_text) + len(before_text):]
                    if '□' in after:
                        t.text = t.text[:t.text.index(before_text) + len(before_text)] + after.replace('□', '☑', 1)
                        return
            
            # 情况3: before_text + □ 跨多个 t
            # 找到 before_text 结束位置后的第一个 □
            current_pos = before_text_end
            
            for t in t_elems:
                t_start = current_pos
                t_end = t_start + len(t.text)
                
                if '□' in t.text:
                    # 计算 □ 在 t 中的位置
                    idx = t.text.index('□')
                    # 检查这个 □ 是否在 before_text 之后
                    relative_pos = t_start + idx - before_text_end
                    if relative_pos >= 0:
                        t.text = t.text.replace('□', '☑', 1)
                        return
                
                current_pos = t_end
    
    def _fill_text_in_cell(self, cell, value: str, target_para_idx: int = 0):
        """
        在单元格的指定段落填充文本（清空后填充）
        """
        if not value:
            return
        
        W_NS = NAMESPACES['w']
        paragraphs = list(cell.iter(f'{{{W_NS}}}p'))
        if target_para_idx >= len(paragraphs):
            return
        
        p = paragraphs[target_para_idx]
        
        # 清空段落中的所有run文本
        for r in p.iter(f'{{{W_NS}}}r'):
            for t in r.iter(f'{{{W_NS}}}t'):
                t.text = ''
        
        # 在第一个run中填充文本
        runs = list(p.iter(f'{{{W_NS}}}r'))
        if runs:
            t = list(runs[0].iter(f'{{{W_NS}}}t'))
            if t:
                t[0].text = value
                t[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        else:
            # 创建新的run和text
            run = ET.SubElement(p, f'{{{W_NS}}}r')
            t = ET.SubElement(run, f'{{{W_NS}}}t')
            t.text = value
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    
    def generate(self, output_filename: str = None) -> str:
        """
        生成填充后的文书
        """
        # 解压并解析模板
        self._unpack()
        self._parse_xml()
        
        W_NS = NAMESPACES['w']
        tables = list(self._root.iter(f'{{{W_NS}}}tbl'))
        
        # ========== T0: 赔偿请求人 ==========
        self._fill_requester(tables[0])
        
        # ========== T1: 委托代理人、赔偿义务机关等 ==========
        self._fill_parties_and_requests(tables[1])
        
        # ========== T2: 人身自由赔偿、事实与理由等 ==========
        self._fill_compensation_and_reasons(tables[2])
        
        # 保存
        os.makedirs(self._get_output_dir(), exist_ok=True)
        if output_filename is None:
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            requester = self.data.get('requester_name', '') or self.data.get('requester_company_name', '未知')
            output_filename = f"国家赔偿申请书-错误执行-{requester[:10]}_{timestamp}.docx"
        
        output_path = os.path.join(self._get_output_dir(), output_filename)
        self._save_and_pack(output_path)
        self._cleanup()
        
        return output_path
    
    def _fill_requester(self, table):
        """填充赔偿请求人区域"""
        rows = list(table.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))
        
        requester_type = self.data.get('requester_type', 'natural')
        
        if requester_type == 'natural':
            # T0R2: 自然人行
            self._fill_natural_person(rows[2])
        else:
            # T0R3: 法人行
            self._fill_legal_person(rows[3])
    
    def _fill_natural_person(self, row):
        """填充自然人信息"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        # 姓名
        name = self.data.get('requester_name', '')
        self._fill_field_in_cell(cell, '姓名：', name)
        
        # 性别
        gender = self.data.get('requester_gender', '')
        if gender:
            self._fill_checkbox_in_cell(cell, f'性别：{gender}', True)
            other_gender = '女' if gender == '男' else '男'
            self._fill_checkbox_in_cell(cell, f'性别：{other_gender}', False)
        
        # 出生日期
        birthdate = self.data.get('requester_birthdate', '')
        self._fill_field_in_cell(cell, '出生日期：', birthdate)
        
        # 民族
        ethnicity = self.data.get('requester_ethnicity', '')
        self._fill_field_in_cell(cell, '民族：', ethnicity)
        
        # 工作单位
        work_unit = self.data.get('requester_work_unit', '')
        self._fill_field_in_cell(cell, '工作单位：', work_unit)
        
        # 职务
        position = self.data.get('requester_position', '')
        self._fill_field_in_cell(cell, '职务：', position)
        
        # 联系电话
        phone = self.data.get('requester_phone', '')
        self._fill_field_in_cell(cell, '联系电话：', phone)
        
        # 住所地
        residence = self.data.get('requester_residence', '')
        self._fill_field_in_cell(cell, '住所地（户籍所在地', '')
        self._fill_field_in_cell(cell, '住所地（户籍所在地）', residence)
        
        # 经常居住地
        living_place = self.data.get('requester_living_place', '')
        self._fill_field_in_cell(cell, '经常居住地：', living_place)
    
    def _fill_legal_person(self, row):
        """填充法人信息"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        # 名称
        name = self.data.get('requester_company_name', '')
        self._fill_field_in_cell(cell, '名称：', name)
        
        # 住所地
        address = self.data.get('requester_company_address', '')
        self._fill_field_in_cell(cell, '住所地（主要办事机构所在地', '')
        self._fill_field_in_cell(cell, '住所地（主要办事机构所在地）', address)
        
        # 注册地
        registered_place = self.data.get('requester_company_registered_place', '')
        self._fill_field_in_cell(cell, '注册地 / 登记地：', registered_place)
        
        # 法定代表人
        legal_rep = self.data.get('requester_legal_representative', '')
        self._fill_field_in_cell(cell, '法定代表人 / 负责人：', legal_rep)
        
        # 职务
        position = self.data.get('requester_representative_position', '')
        self._fill_field_in_cell(cell, '职务：', position)
        
        # 联系电话
        phone = self.data.get('requester_phone', '')
        self._fill_field_in_cell(cell, '联系电话：', phone)
        
        # 统一社会信用代码
        credit_code = self.data.get('requester_unified_social_credit_code', '')
        self._fill_field_in_cell(cell, '统一社会信用代码：', credit_code)
        
        # 公司类型勾选
        company_type = self.data.get('requester_company_type', '')
        type_map = {
            '有限责任公司': '有限责任公司',
            '股份有限公司': '股份有限公司',
            '上市公司': '上市公司',
        }
        for db_type, label in type_map.items():
            checked = (company_type == db_type)
            self._fill_checkbox_in_cell(cell, f'{label}□', checked)
    
    def _fill_parties_and_requests(self, table):
        """填充T1: 委托代理人、赔偿义务机关、各项赔偿请求"""
        rows = list(table.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))
        
        # T1R0: 委托代理人
        self._fill_agent(rows[0])
        
        # T1R1: 赔偿义务机关
        self._fill_compensation_authority(rows[1])
        
        # T1R2: 自赔决定
        self._fill_self_decision(rows[2])
        
        # T1R3: 执行标的
        self._fill_execution_subject(rows[3])
        
        # T1R6: 财产权赔偿
        self._fill_property_compensation(rows[6])
        
        # T1R7: 停产停业损失
        self._fill_business_loss(rows[7])
    
    def _fill_agent(self, row):
        """填充委托代理人"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_agent = self.data.get('has_agent', False)
        self._fill_checkbox_in_cell(cell, '有□', has_agent)
        self._fill_checkbox_in_cell(cell, '无□', not has_agent)
        
        if has_agent:
            # 代理人类型
            agent_type = self.data.get('agent_type', [])
            if isinstance(agent_type, str):
                agent_type = [agent_type]
            for at in agent_type:
                self._fill_checkbox_in_cell(cell, f'{at}□', True)
            
            # 姓名
            name = self.data.get('agent_name', '')
            self._fill_field_in_cell(cell, '姓名：', name)
            
            # 单位
            unit = self.data.get('agent_unit', '')
            self._fill_field_in_cell(cell, '单位：', unit)
            
            # 职务
            position = self.data.get('agent_position', '')
            self._fill_field_in_cell(cell, '职务：', position)
            
            # 联系电话
            phone = self.data.get('agent_phone', '')
            self._fill_field_in_cell(cell, '联系电话：', phone)
            
            # 代理权限
            authority = self.data.get('agent_authority', '')
            if authority == '特别授权':
                self._fill_checkbox_in_cell(cell, '特别授权□', True)
            elif authority == '一般授权':
                self._fill_checkbox_in_cell(cell, '一般授权□', True)
            elif authority == '无':
                self._fill_checkbox_in_cell(cell, '无□', True)
    
    def _fill_compensation_authority(self, row):
        """填充赔偿义务机关"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        name = self.data.get('compensation_authority_name', '')
        self._fill_field_in_cell(cell, '名称：', name)
        
        address = self.data.get('compensation_authority_address', '')
        self._fill_field_in_cell(cell, '住所地：', address)
        
        legal_rep = self.data.get('compensation_authority_legal_representative', '')
        self._fill_field_in_cell(cell, '法定代表人 / 负责人：', legal_rep)
        
        position = self.data.get('compensation_authority_position', '')
        self._fill_field_in_cell(cell, '职务：', position)
    
    def _fill_self_decision(self, row):
        """填充自赔决定"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_decision = self.data.get('has_self_decision', False)
        self._fill_checkbox_in_cell(cell, '是□', has_decision)
        self._fill_checkbox_in_cell(cell, '否□', not has_decision)
        
        if has_decision:
            number = self.data.get('decision_number', '')
            self._fill_field_in_cell(cell, '决定书文号：', number)
            
            result = self.data.get('decision_result', '')
            self._fill_field_in_cell(cell, '决定书结果：', result)
    
    def _fill_execution_subject(self, row):
        """填充执行标的"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        subject = self.data.get('execution_subject_matter', '')
        self._fill_text_in_cell(cell, subject)
    
    def _fill_property_compensation(self, row):
        """填充财产权赔偿"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_property = self.data.get('has_property_compensation', False)
        self._fill_checkbox_in_cell(cell, '是否主张以下赔偿：', has_property)
        self._fill_checkbox_in_cell(cell, '是□', has_property)
        
        if has_property:
            comp_type = self.data.get('property_compensation_type', '')
            if comp_type == '返还原物':
                self._fill_checkbox_in_cell(cell, '返还原物□', True)
            elif comp_type == '恢复原状':
                self._fill_checkbox_in_cell(cell, '恢复原状□', True)
            elif '赔偿损失' in str(comp_type):
                self._fill_checkbox_in_cell(cell, '赔偿损失□', True)
                amount = self.data.get('property_compensation_amount', '')
                self._fill_field_in_cell(cell, '金额：', amount)
    
    def _fill_business_loss(self, row):
        """填充停产停业损失"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_loss = self.data.get('has_business_loss', False)
        self._fill_checkbox_in_cell(cell, '是否主张以下赔偿：', has_loss)
        self._fill_checkbox_in_cell(cell, '是□', has_loss)
        
        if has_loss:
            loss_type = self.data.get('business_loss_type', '')
            
            if loss_type == '留守职工工资':
                self._fill_checkbox_in_cell(cell, '1. 必要留守职工工资□', True)
                amount = self.data.get('employee_wages_amount', '')
                self._fill_field_in_cell(cell, '金额', amount)
                calc = self.data.get('employee_wages_calculation', '')
                self._fill_field_in_cell(cell, '计算依据：', calc)
            
            elif loss_type == '税款社会保险费':
                self._fill_checkbox_in_cell(cell, '2. 必须缴纳的税款、社会保险费□', True)
                amount = self.data.get('tax_social_fee_amount', '')
                self._fill_field_in_cell(cell, '金额', amount)
                calc = self.data.get('tax_social_fee_calculation', '')
                self._fill_field_in_cell(cell, '计算依据：', calc)
            
            elif loss_type == '水电费保管费仓储费承包费':
                self._fill_checkbox_in_cell(cell, '3. 应当缴纳的水电费、保管费、仓储费、承包费□', True)
                amount = self.data.get('utilities_rent_amount', '')
                self._fill_field_in_cell(cell, '金额', amount)
                calc = self.data.get('utilities_rent_calculation', '')
                self._fill_field_in_cell(cell, '计算依据：', calc)
    
    def _fill_compensation_and_reasons(self, table):
        """填充T2: 人身自由赔偿、事实与理由等"""
        rows = list(table.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))
        
        # T2R0: 人身自由赔偿
        self._fill_personal_freedom(rows[0])
        
        # T2R1: 其他赔偿请求
        self._fill_other_compensation(rows[1])
        
        # T2R4: 法律依据
        self._fill_legal_basis(rows[4])
        
        # T2R5: 其他说明
        self._fill_other_explanations(rows[5])
        
        # T2R6: 同类案件
        self._fill_similar_cases(rows[6])
        
        # T2R7: 证据清单
        self._fill_evidence(rows[7])
    
    def _fill_personal_freedom(self, row):
        """填充人身自由赔偿"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_freedom = self.data.get('has_personal_freedom_compensation', False)
        self._fill_checkbox_in_cell(cell, '是否主张以下赔偿：', has_freedom)
        self._fill_checkbox_in_cell(cell, '是□', has_freedom)
        
        if has_freedom:
            days = self.data.get('illegal_detention_days', '')
            self._fill_field_in_cell(cell, '错误拘留时间共', days)
            
            period = self.data.get('detention_period', '')
            self._fill_field_in_cell(cell, '起止日期为：', period)
            
            amount = self.data.get('personal_freedom_compensation_amount', '')
            self._fill_field_in_cell(cell, '请求赔偿人身自由赔偿金', amount)
    
    def _fill_other_compensation(self, row):
        """填充其他赔偿请求"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        other = self.data.get('other_compensation_requests', '')
        self._fill_text_in_cell(cell, other)
    
    def _fill_legal_basis(self, row):
        """填充法律依据"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        basis = self.data.get('legal_basis_and_reasons', '')
        self._fill_text_in_cell(cell, basis)
    
    def _fill_other_explanations(self, row):
        """填充其他说明"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        explanations = self.data.get('other_explanations', '')
        self._fill_text_in_cell(cell, explanations)
    
    def _fill_similar_cases(self, row):
        """填充同类案件"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        
        has_cases = self.data.get('has_similar_cases', False)
        self._fill_checkbox_in_cell(cell, '是□', has_cases)
        self._fill_checkbox_in_cell(cell, '否□', not has_cases)
        
        if has_cases:
            case_info = self.data.get('similar_case_number', '')
            self._fill_field_in_cell(cell, '案号 /', case_info)
            self._fill_field_in_cell(cell, '案例名称：', case_info)
    
    def _fill_evidence(self, row):
        """填充证据清单"""
        cells = list(row.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'))
        if len(cells) < 2:
            return
        
        cell = cells[1]
        evidence = self.data.get('evidence_list', '')
        self._fill_text_in_cell(cell, evidence)
