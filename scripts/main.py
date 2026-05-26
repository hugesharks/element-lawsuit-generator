#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
要素式文书一键生成器 v3 - 主入口

核心改进：使用区域定位的填充策略，避免跨区域误填
"""

import os
import sys
import json
import argparse
from datetime import datetime
from typing import Optional, Dict, List, Any

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)

from file_parser import FileParser
from case_classifier import CaseClassifier
from template_manager import TemplateManager
from content_extractor import ContentExtractor
from template_filler import TemplateFiller


class ElementLawsuitGenerator:
    """要素式文书一键生成器"""

    def __init__(self, local_template_dir=None, local_example_dir=None, 
                 output_dir=None, config_dir=None):
        if config_dir is None:
            config_dir = os.path.join(BASE_DIR, 'configs')
        if output_dir is None:
            output_dir = os.path.join(BASE_DIR, 'output')
        
        self.config_dir = config_dir
        self.output_dir = output_dir
        self.local_template_dir = local_template_dir
        self.local_example_dir = local_example_dir
        
        os.makedirs(self.output_dir, exist_ok=True)
        
        self.parser = FileParser()
        self.classifier = CaseClassifier(config_dir)
        self.template_mgr = TemplateManager(
            cache_dir=os.path.join(BASE_DIR, 'template_cache'),
            config_dir=config_dir,
            local_template_dir=local_template_dir
        )
        self.extractor = ContentExtractor(config_dir)

    def generate(self, file_path, case_type_override=None, doc_type_override=None):
        """一键生成要素式文书"""
        result = {
            "success": False, "output_path": "", "case_type": "", "doc_type": "",
            "confidence": 0.0, "template_used": "", "extracted_data": {},
            "errors": [], "warnings": [],
        }
        
        try:
            print(f"[1/6] 解析文件: {file_path}")
            text = self.parser.parse(file_path)
            if not text or not text.strip():
                result["errors"].append("文件内容为空")
                return result
            
            print(f"[2/6] 识别案由和文书类型")
            if case_type_override:
                case_type = case_type_override
                doc_type = doc_type_override or self._infer_doc_type(text)
                confidence = 1.0
            else:
                case_type, doc_type, confidence = self.classifier.classify(text)
                if confidence < 0.3:
                    result["warnings"].append(f"案由识别置信度较低({confidence})，建议人工确认")
            
            result["case_type"] = case_type
            result["doc_type"] = doc_type
            result["confidence"] = confidence
            
            if case_type == "未知":
                result["errors"].append("无法识别案由，请手动指定")
                return result
            
            print(f"[3/6] 获取模板")
            template_filename = self.classifier.get_template_filename(case_type, doc_type)
            template_path = self.template_mgr.get_template(case_type, doc_type)
            if not template_path:
                result["errors"].append(f"未找到模板: {template_filename}")
                return result
            result["template_used"] = template_filename
            
            print(f"[4/6] 提取关键要素")
            extracted = self.extractor.extract(text, case_type, doc_type)
            result["extracted_data"] = extracted
            
            print(f"[5/6] 构建区域化填充数据")
            fill_data = self._build_fill_data(extracted, case_type, doc_type)
            
            print(f"[6/6] 填充模板")
            output_filename = self._generate_output_filename(case_type, doc_type)
            filler = TemplateFiller(template_path, self.output_dir)
            output_path = filler.fill(fill_data, output_filename)
            
            result["success"] = True
            result["output_path"] = output_path
            print(f"  → 生成成功: {output_path}")
            
        except Exception as e:
            result["errors"].append(f"生成失败: {str(e)}")
            import traceback
            traceback.print_exc()
        
        return result

    def _infer_doc_type(self, text):
        _, doc_type, _ = self.classifier.classify(text)
        return doc_type

    def _build_fill_data(self, extracted: Dict, case_type: str, doc_type: str) -> Dict:
        """
        构建区域化填充数据
        
        使用 section_fills 结构，每个 section 对应模板中的一个段落区域
        这样可以精确控制字段填充的位置，避免跨区域误填
        """
        fill_data = {
            'section_fills': [],
            'checkbox_ops': [],
            'text_replacements': {},
        }
        
        parties = extracted.get('parties', {})
        signature = extracted.get('signature', {})
        
        # === 原告自然人区域 ===
        plaintiff = parties.get('plaintiff', {})
        if plaintiff and plaintiff.get('type') == 'natural':
            fields = {}
            
            if plaintiff.get('name'):
                fields['姓名：'] = plaintiff['name']
            if plaintiff.get('birthdate'):
                fields['出生日期：'] = plaintiff['birthdate']
            if plaintiff.get('ethnicity'):
                fields['民族：'] = plaintiff['ethnicity']
            if plaintiff.get('work_unit'):
                fields['工作单位：'] = plaintiff['work_unit']
            if plaintiff.get('position'):
                fields['职务：'] = plaintiff['position']
            if plaintiff.get('phone'):
                fields['联系电话：'] = plaintiff['phone']
            if plaintiff.get('address'):
                fields['住所地（户籍所在地）：'] = plaintiff['address']
            if plaintiff.get('residence'):
                fields['经常居住地：'] = plaintiff['residence']
            if plaintiff.get('id_type'):
                fields['证件类型：'] = plaintiff['id_type']
            if plaintiff.get('id_number'):
                fields['证件号码：'] = plaintiff['id_number']
            
            fill_data['section_fills'].append({
                'section': '原告_自然人',
                'fields': fields,
                'checkboxes': {},  # 勾选框统一用checkbox_ops处理
            })
            
            # 原告性别勾选框 - 使用精确的checkbox_ops
            gender = plaintiff.get('gender', '')
            if gender:
                fill_data['checkbox_ops'].append({
                    'section': '原告_自然人',
                    'paragraph_contains': '性别',
                    'before_checkbox': gender,
                    'check': True,
                })
        
        # === 原告法人区域 ===
        if plaintiff and plaintiff.get('type') == 'legal':
            fields = {}
            if plaintiff.get('name'):
                fields['名称：'] = plaintiff['name']
            if plaintiff.get('address'):
                fields['住所地（主要办事机构所在地）：'] = plaintiff['address']
            if plaintiff.get('legal_person'):
                fields['法定代表人 / 负责人：'] = plaintiff['legal_person']
            if plaintiff.get('phone'):
                fields['联系电话：'] = plaintiff['phone']
            if plaintiff.get('credit_code'):
                fields['统一社会信用代码：'] = plaintiff['credit_code']
            
            fill_data['section_fills'].append({
                'section': '原告_法人',
                'fields': fields,
                'checkboxes': {},
            })
        
        # === 被告自然人区域 ===
        defendant = parties.get('defendant', {})
        if defendant and defendant.get('type') == 'natural':
            fields = {}
            
            if defendant.get('name'):
                fields['姓名：'] = defendant['name']
            if defendant.get('phone'):
                fields['联系电话：'] = defendant['phone']
            if defendant.get('address'):
                fields['住所地（户籍所在地）：'] = defendant['address']
            if defendant.get('residence'):
                fields['经常居住地：'] = defendant['residence']
            if defendant.get('id_number'):
                fields['证件号码：'] = defendant['id_number']
            if defendant.get('ethnicity'):
                fields['民族：'] = defendant['ethnicity']
            if defendant.get('work_unit'):
                fields['工作单位：'] = defendant['work_unit']
            if defendant.get('position'):
                fields['职务：'] = defendant['position']
            if defendant.get('id_type'):
                fields['证件类型：'] = defendant['id_type']
            if defendant.get('birthdate'):
                fields['出生日期：'] = defendant['birthdate']
            
            fill_data['section_fills'].append({
                'section': '被告_自然人',
                'fields': fields,
                'checkboxes': {},
            })
            
            # 被告性别勾选框
            gender = defendant.get('gender', '')
            if gender:
                fill_data['checkbox_ops'].append({
                    'section': '被告_自然人',
                    'paragraph_contains': '性别',
                    'before_checkbox': gender,
                    'check': True,
                })
        
        # === 被告法人区域 ===
        if defendant and defendant.get('type') == 'legal':
            fields = {}
            if defendant.get('name'):
                fields['名称：'] = defendant['name']
            if defendant.get('address'):
                fields['住所地（主要办事机构所在地）：'] = defendant['address']
            if defendant.get('phone'):
                fields['联系电话：'] = defendant['phone']
            if defendant.get('credit_code'):
                fields['统一社会信用代码：'] = defendant['credit_code']
            
            fill_data['section_fills'].append({
                'section': '被告_法人',
                'fields': fields,
                'checkboxes': {},
            })
        
        # === 委托诉讼代理人区域 ===
        agent = parties.get('agent', {})
        if agent:
            has_agent = agent.get('has_agent', False)
            fields = {}
            checkboxes = {}
            
            checkboxes['有'] = has_agent
            if not has_agent:
                checkboxes['无'] = True
            
            if has_agent:
                if agent.get('name'):
                    fields['姓名：'] = agent['name']
                if agent.get('unit'):
                    fields['单位：'] = agent['unit']
                if agent.get('position'):
                    fields['职务：'] = agent['position']
                if agent.get('phone'):
                    fields['联系电话：'] = agent['phone']
                
                authority = agent.get('authority', '')
                if authority == '特别授权':
                    checkboxes['特别授权'] = True
                elif authority == '一般授权':
                    checkboxes['一般授权'] = True
                else:
                    checkboxes['无'] = True
            
            fill_data['section_fills'].append({
                'section': '委托诉讼代理人',
                'fields': fields,
                'checkboxes': checkboxes,
            })
        
        # === 诉讼请求区域（案由特定字段）===
        case_specific = extracted.get('case_specific', {})
        if case_specific:
            request_fields = {}
            reason_fields = {}
            jurisdiction_fields = {}
            
            # 通用字段映射：提取数据key → 模板中的标签文本
            REQUEST_FIELD_MAP = {
                '本金金额': '截至',
                '利息金额': '截至',
                '标的总额': '标的总额',
                '诉讼请求概括': '诉讼请求',
            }
            
            REASON_FIELD_MAP = {
                '合同签订情况': '合同签订情况',
                '出借人': '出借人',
                '借款人': '借款人',
                '约定借款金额': '约定',
                '实际提供金额': '实际提供',
                '约定期限': '约定期限',
                '逾期时间': '逾期时间',
                '已还本金': '已还本金',
                '已还利息': '已还利息',
                '请求依据合同约定': '合同约定',
                '请求依据法律规定': '法律规定',
                '事实与理由概括': '事实与理由',
            }
            
            JURISDICTION_FIELD_MAP = {
                '管辖约定': '合同条款及内容',
            }
            
            for key, label in REQUEST_FIELD_MAP.items():
                if case_specific.get(key):
                    request_fields[label] = case_specific[key]
            
            for key, label in REASON_FIELD_MAP.items():
                if case_specific.get(key):
                    reason_fields[label] = case_specific[key]
                    
            for key, label in JURISDICTION_FIELD_MAP.items():
                if case_specific.get(key):
                    jurisdiction_fields[label] = case_specific[key]
            
            if request_fields:
                fill_data['section_fills'].append({
                    'section': '诉讼请求',
                    'fields': request_fields,
                    'checkboxes': {},
                })
            
            if reason_fields:
                fill_data['section_fills'].append({
                    'section': '事实与理由',
                    'fields': reason_fields,
                    'checkboxes': {},
                })
            
            if jurisdiction_fields:
                fill_data['section_fills'].append({
                    'section': '约定管辖和诉前保全',
                    'fields': jurisdiction_fields,
                    'checkboxes': {},
                })
        
        # === 勾选框操作（使用精确的 checkbox_ops）===
        # 从提取的数据中构建精确勾选操作
        checkbox_ops = self._build_checkbox_ops(extracted, case_type)
        fill_data['checkbox_ops'] = checkbox_ops
        
        # === 签名 ===
        if signature.get('signer'):
            fill_data['text_replacements']['具状人（签字、盖章）：'] = \
                f'具状人（签字、盖章）：{signature["signer"]}'
        
        return fill_data

    def _build_checkbox_ops(self, extracted: Dict, case_type: str) -> List[Dict]:
        """
        构建精确的勾选框操作
        
        关键：要区分"是/否"等通用勾选框在不同区域的含义
        """
        ops = []
        parties = extracted.get('parties', {})
        
        # 原告性别勾选
        plaintiff = parties.get('plaintiff', {})
        if plaintiff and plaintiff.get('type') == 'natural':
            gender = plaintiff.get('gender', '')
            if gender:
                ops.append({
                    'section': '原告_自然人',
                    'paragraph_contains': '性别',
                    'before_checkbox': gender,
                    'check': True,
                })
        
        # 被告性别勾选
        defendant = parties.get('defendant', {})
        if defendant and defendant.get('type') == 'natural':
            gender = defendant.get('gender', '')
            if gender:
                ops.append({
                    'section': '被告_自然人',
                    'paragraph_contains': '性别',
                    'before_checkbox': gender,
                    'check': True,
                })
        
        # 根据案由添加特定勾选框
        if case_type == '民间借贷纠纷':
            case_specific = extracted.get('case_specific', {})
            
            # 是否有管辖约定
            if case_specific.get('管辖约定'):
                ops.append({
                    'section': '',  # 全局搜索
                    'paragraph_contains': '有无仲裁',
                    'before_checkbox': '有',
                    'check': True,
                })
            
            # 是否诉前保全 - 否
            ops.append({
                'section': '',
                'paragraph_contains': '诉前保全',
                'before_checkbox': '否',
                'check': True,
            })
        
        return ops

    def _generate_output_filename(self, case_type, doc_type):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"{doc_type}-{case_type}-要素式_{timestamp}.docx"


def main():
    parser = argparse.ArgumentParser(description='要素式文书一键生成器')
    parser.add_argument('file', help='输入文书文件路径')
    parser.add_argument('--output-dir', '-o', default=None)
    parser.add_argument('--local-templates', '-t', default=None)
    parser.add_argument('--case-type', '-c', default=None)
    parser.add_argument('--doc-type', '-d', default=None)
    parser.add_argument('--config-dir', default=None)
    
    args = parser.parse_args()
    
    generator = ElementLawsuitGenerator(
        local_template_dir=args.local_templates,
        output_dir=args.output_dir,
        config_dir=args.config_dir,
    )
    
    result = generator.generate(
        args.file,
        case_type_override=args.case_type,
        doc_type_override=args.doc_type,
    )
    
    print("\n" + "=" * 60)
    print("生成结果")
    print("=" * 60)
    for k, v in result.items():
        if k not in ('extracted_data',):
            print(f"  {k}: {v}")
    
    return 0 if result['success'] else 1


if __name__ == '__main__':
    sys.exit(main())
