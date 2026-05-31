# -*- coding: utf-8 -*-
"""
离婚纠纷适配器

覆盖离婚纠纷案件（离婚、财产分割、债务、子女抚养、抚养费、探望权、损害赔偿等）。

模板结构（民事起诉状-离婚纠纷.docx）：
- T0: 当事人信息（原告自然人、委托代理人）
- T1: 被告信息 + 诉讼请求（离婚/财产分割/债务/子女抚养/抚养费）
- T2: 探望权续 + 损害赔偿续 + 诉前保全 + 事实与理由
- T3: 请求依据 + 证据清单 + 调解意愿

书签命名规范：{表前缀}_{序号}_{字段名}
- T0_原告_姓名、T0_原告_性别、T0_原告_出生日期...
- T1_被告_姓名、T1_诉讼请求_离婚...
- T2_事实理由_婚姻基本情况...
- T3_调解意愿_了解调解...
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter


class MarriageFamilyAdapter(CaseAdapter):
    """
    离婚纠纷适配器
    
    支持：
    - 解除婚姻关系
    - 夫妻共同财产分割
    - 夫妻共同债务分担
    - 子女直接抚养权
    - 子女抚养费
    - 探望权
    - 离婚损害赔偿 / 离婚经济补偿 / 离婚经济帮助
    - 诉前保全
    """
    
    def name(self) -> str:
        return "离婚纠纷"
    
    def get_template_name(self) -> str:
        return "民事起诉状-离婚纠纷.docx"
    
    def get_schema(self) -> dict:
        return {
            "fields": [
                # === 原告信息 ===
                {"name": "plaintiff_name", "label": "原告姓名", "type": "str", "required": True},
                {"name": "plaintiff_gender", "label": "原告性别", "type": "str", "required": True, "options": ["男", "女"]},
                {"name": "plaintiff_birth_date", "label": "原告出生日期", "type": "str", "required": True, "description": "格式：1988年3月15日"},
                {"name": "plaintiff_ethnicity", "label": "原告民族", "type": "str", "required": True, "description": "如：汉族"},
                {"name": "plaintiff_work_unit", "label": "原告工作单位", "type": "str", "required": False},
                {"name": "plaintiff_position", "label": "原告职务", "type": "str", "required": False},
                {"name": "plaintiff_phone", "label": "原告联系电话", "type": "str", "required": False},
                {"name": "plaintiff_address", "label": "原告住所地", "type": "str", "required": True, "description": "户籍所在地"},
                {"name": "plaintiff_residence", "label": "原告经常居住地", "type": "str", "required": False},
                {"name": "plaintiff_id_type", "label": "原告证件类型", "type": "str", "required": True, "default": "居民身份证"},
                {"name": "plaintiff_id_number", "label": "原告身份证号", "type": "str", "required": True},
                
                # === 被告信息 ===
                {"name": "defendant_name", "label": "被告姓名", "type": "str", "required": True},
                {"name": "defendant_gender", "label": "被告性别", "type": "str", "required": True, "options": ["男", "女"]},
                {"name": "defendant_birth_date", "label": "被告出生日期", "type": "str", "required": False},
                {"name": "defendant_ethnicity", "label": "被告民族", "type": "str", "required": False},
                {"name": "defendant_work_unit", "label": "被告工作单位", "type": "str", "required": False},
                {"name": "defendant_position", "label": "被告职务", "type": "str", "required": False},
                {"name": "defendant_phone", "label": "被告联系电话", "type": "str", "required": False},
                {"name": "defendant_address", "label": "被告住所地", "type": "str", "required": True},
                {"name": "defendant_residence", "label": "被告经常居住地", "type": "str", "required": False},
                {"name": "defendant_id_type", "label": "被告证件类型", "type": "str", "required": False},
                {"name": "defendant_id_number", "label": "被告身份证号", "type": "str", "required": False},
                
                # === 委托诉讼代理人 ===
                {"name": "has_agent", "label": "是否有委托诉讼代理人", "type": "bool", "required": False, "default": False},
                {"name": "agent_name", "label": "代理人姓名", "type": "str", "required": False},
                {"name": "agent_unit", "label": "代理人单位", "type": "str", "required": False},
                {"name": "agent_position", "label": "代理人职务", "type": "str", "required": False, "default": "律师"},
                {"name": "agent_phone", "label": "代理人联系电话", "type": "str", "required": False},
                {"name": "agent_authorization", "label": "代理权限", "type": "str", "required": False, "options": ["一般授权", "特别授权"]},
                
                # === 送达地址 ===
                {"name": "delivery_address", "label": "送达地址", "type": "str", "required": False},
                
                # === 诉讼请求 ===
                {"name": "divorce_request", "label": "是否请求解除婚姻关系", "type": "bool", "required": True, "default": True},
                {"name": "divorce_specific", "label": "解除婚姻关系具体主张", "type": "str", "required": False},
                
                # 财产分割
                {"name": "has_property", "label": "是否有夫妻共同财产", "type": "bool", "required": False, "default": False},
                {"name": "property_detail", "label": "财产分割明细", "type": "str", "required": False},
                
                # 债务分担
                {"name": "has_debt", "label": "是否有夫妻共同债务", "type": "bool", "required": False, "default": False},
                {"name": "debt_detail", "label": "债务分担明细", "type": "str", "required": False},
                
                # 子女抚养
                {"name": "has_children", "label": "是否有子女抚养问题", "type": "bool", "required": False, "default": False},
                {"name": "children", "label": "子女信息", "type": "list", "required": False, "children": [
                    {"name": "name", "label": "子女姓名", "type": "str"},
                    {"name": "age", "label": "子女年龄", "type": "int"},
                    {"name": "custody", "label": "抚养归属", "type": "str", "options": ["原告", "被告"]},
                    {"name": "monthly_support", "label": "月抚养费", "type": "float"},
                ]},
                
                # 抚养费
                {"name": "has_child_support", "label": "是否主张抚养费", "type": "bool", "required": False, "default": False},
                {"name": "child_support_payer", "label": "抚养费承担主体", "type": "str", "required": False, "options": ["原告", "被告"]},
                {"name": "child_support_detail", "label": "抚养费明细", "type": "str", "required": False},
                
                # 探望权
                {"name": "has_visitation", "label": "是否有探望权问题", "type": "bool", "required": False, "default": False},
                {"name": "visitation_holder", "label": "探望权行使主体", "type": "str", "required": False, "options": ["原告", "被告"]},
                {"name": "visitation_detail", "label": "探望权行使方式", "type": "str", "required": False},
                
                # 损害赔偿/补偿/帮助
                {"name": "has_damage_compensation", "label": "是否有损害赔偿/补偿/帮助", "type": "bool", "required": False, "default": False},
                {"name": "damage_type", "label": "损害类型", "type": "str", "required": False, "options": ["离婚损害赔偿", "离婚经济补偿", "离婚经济帮助"]},
                {"name": "damage_amount", "label": "损害金额", "type": "float", "required": False},
                
                # 诉讼费用
                {"name": "claim_litigation_fee", "label": "是否主张诉讼费用", "type": "bool", "required": True, "default": True},
                
                # 其他请求
                {"name": "other_requests", "label": "其他诉讼请求", "type": "str", "required": False},
                
                # === 诉前保全 ===
                {"name": "has_preservation", "label": "是否已申请诉前保全", "type": "bool", "required": False, "default": False},
                {"name": "preservation_court", "label": "保全法院", "type": "str", "required": False},
                {"name": "preservation_time", "label": "保全时间", "type": "str", "required": False},
                {"name": "preservation_case_number", "label": "保全案号", "type": "str", "required": False},
                
                # === 事实与理由 ===
                {"name": "marriage_start_date", "label": "结婚时间", "type": "str", "required": False},
                {"name": "children_status", "label": "生育子女情况", "type": "str", "required": False},
                {"name": "living_status", "label": "双方生活情况", "type": "str", "required": False},
                {"name": "divorce_reason", "label": "离婚事由", "type": "str", "required": False},
                {"name": "previous_divorce_suit", "label": "之前是否提起过离婚诉讼", "type": "str", "required": False},
                
                {"name": "property_facts", "label": "夫妻共同财产事实", "type": "str", "required": False},
                {"name": "debt_facts", "label": "夫妻共同债务事实", "type": "str", "required": False},
                
                {"name": "custody_facts", "label": "子女直接抚养事实", "type": "str", "required": False},
                {"name": "child_support_facts", "label": "抚养费事实", "type": "str", "required": False},
                {"name": "visitation_facts", "label": "探望权事实", "type": "str", "required": False},
                {"name": "damage_facts", "label": "赔偿/补偿/帮助事实", "type": "str", "required": False},
                
                {"name": "other_facts", "label": "其他事实", "type": "str", "required": False},
                
                # === 调解意愿 ===
                {"name": "understand_mediation", "label": "是否了解调解", "type": "bool", "required": False, "default": False},
                {"name": "understand_mediation_benefits", "label": "是否了解先行调解好处", "type": "bool", "required": False, "default": False},
                {"name": "consider_mediation", "label": "是否考虑先行调解", "type": "str", "required": False, "options": ["是", "否", "暂不确定"]},
            ]
        }
    
    def calculate(self, case_data: dict) -> dict:
        """
        离婚纠纷费用计算
        
        离婚纠纷主要涉及：
        - 财产分割（协商确定，不一定需要法院计算）
        - 子女抚养费（根据收入、当地生活水平等确定）
        - 损害赔偿/补偿/帮助
        """
        items = []
        total = 0.0
        
        # 损害赔偿/补偿/帮助
        if case_data.get('has_damage_compensation') and case_data.get('damage_amount'):
            amount = case_data['damage_amount']
            items.append({
                "name": case_data.get('damage_type', '损害赔偿'),
                "amount": amount,
                "formula": f"金额：{amount}元",
                "legal_basis": "《民法典》第1091条/第1088条/第1090条",
            })
            total += amount
        
        # 子女抚养费总额
        if case_data.get('children'):
            for child in case_data['children']:
                if child.get('monthly_support'):
                    items.append({
                        "name": f"子女{child.get('name', '')}抚养费",
                        "amount": child.get('monthly_support', 0),
                        "formula": f"每月{child.get('monthly_support')}元",
                        "legal_basis": "《民法典》第1085条",
                    })
        
        return {
            "items": items,
            "total": total,
        }
    
    def build_fill_map(self, case_data: dict, calc_result: dict) -> dict[str, str]:
        """生成书签填充映射表"""
        fill_map = {}
        
        # ============================================================
        # T0: 当事人信息
        # ============================================================
        
        # T0 Row 2: 原告信息 - Cell 1
        # Para 0: 姓名
        fill_map['T0_原告_姓名'] = case_data.get('plaintiff_name', '')
        # Para 2: 出生日期+民族 - 注意模板结构：出生日期占位后直接跟民族
        birth_date = case_data.get('plaintiff_birth_date', '')
        ethnicity = case_data.get('plaintiff_ethnicity', '')
        fill_map['T0_原告_出生日期'] = birth_date
        fill_map['T0_原告_民族'] = ethnicity
        # Para 3: 工作单位+职务+联系电话
        fill_map['T0_原告_工作单位'] = case_data.get('plaintiff_work_unit', '')
        fill_map['T0_原告_职务'] = case_data.get('plaintiff_position', '')
        fill_map['T0_原告_联系电话'] = case_data.get('plaintiff_phone', '')
        # Para 4: 住所地
        fill_map['T0_原告_住所地'] = case_data.get('plaintiff_address', '')
        # Para 5: 经常居住地
        fill_map['T0_原告_经常居住地'] = case_data.get('plaintiff_residence', '')
        # Para 6: 证件类型
        fill_map['T0_原告_证件类型'] = case_data.get('plaintiff_id_type', '居民身份证')
        # Para 7: 证件号码
        fill_map['T0_原告_证件号码'] = case_data.get('plaintiff_id_number', '')
        
        # T0 Row 3: 委托代理人 - Cell 1
        if case_data.get('has_agent'):
            # Para 1: 姓名
            fill_map['T0_代理人_姓名'] = case_data.get('agent_name', '')
            # Para 2: 单位+职务+联系电话
            fill_map['T0_代理人_单位'] = case_data.get('agent_unit', '')
            fill_map['T0_代理人_职务'] = case_data.get('agent_position', '律师')
            fill_map['T0_代理人_联系电话'] = case_data.get('agent_phone', '')
        
        # ============================================================
        # T1: 被告信息 + 诉讼请求
        # ============================================================
        
        # T1 Row 0: 被告信息 - Cell 1
        fill_map['T1_被告_姓名'] = case_data.get('defendant_name', '')
        # Para 2: 出生日期+民族
        def_birth = case_data.get('defendant_birth_date', '')
        def_ethnicity = case_data.get('defendant_ethnicity', '')
        fill_map['T1_被告_出生日期'] = def_birth
        fill_map['T1_被告_民族'] = def_ethnicity
        # Para 3: 工作单位+职务+联系电话
        fill_map['T1_被告_工作单位'] = case_data.get('defendant_work_unit', '')
        fill_map['T1_被告_职务'] = case_data.get('defendant_position', '')
        fill_map['T1_被告_联系电话'] = case_data.get('defendant_phone', '')
        # Para 4: 住所地
        fill_map['T1_被告_住所地'] = case_data.get('defendant_address', '')
        # Para 5: 经常居住地
        fill_map['T1_被告_经常居住地'] = case_data.get('defendant_residence', '')
        # Para 6: 证件类型
        fill_map['T1_被告_证件类型'] = case_data.get('defendant_id_type', '')
        # Para 7: 证件号码
        fill_map['T1_被告_证件号码'] = case_data.get('defendant_id_number', '')
        
        # T1 Row 3: 解除婚姻关系 - Cell 1
        fill_map['T1_离婚主张'] = case_data.get('divorce_specific', '')
        
        # T1 Row 4: 夫妻共同财产 - Cell 1
        fill_map['T1_财产明细'] = case_data.get('property_detail', '')
        
        # T1 Row 5: 夫妻共同债务 - Cell 1
        fill_map['T1_债务明细'] = case_data.get('debt_detail', '')
        
        # T1 Row 6: 子女直接抚养 - Cell 1
        # 填充子女信息
        children = case_data.get('children', [])
        if children:
            child = children[0]
            fill_map['T1_子女1姓名'] = child.get('name', '')
            fill_map['T1_子女1归属'] = child.get('custody', '')
        
        # T1 Row 7: 子女抚养费 - Cell 1
        fill_map['T1_抚养费明细'] = case_data.get('child_support_detail', '')
        
        # ============================================================
        # T2: 事实与理由
        # ============================================================
        
        # T2 Row 0: 探望权 - Cell 1
        fill_map['T2_探望权行使方式'] = case_data.get('visitation_detail', '')
        
        # T2 Row 1: 损害赔偿/补偿/帮助 - Cell 1
        if case_data.get('has_damage_compensation') and case_data.get('damage_amount'):
            fill_map['T2_损害金额'] = str(int(case_data['damage_amount']))
        
        # T2 Row 8: 婚姻关系基本情况 - Cell 1
        fill_map['T2_结婚时间'] = case_data.get('marriage_start_date', '')
        fill_map['T2_生育子女情况'] = case_data.get('children_status', '')
        fill_map['T2_双方生活情况'] = case_data.get('living_status', '')
        fill_map['T2_离婚事由'] = case_data.get('divorce_reason', '')
        fill_map['T2_之前离婚诉讼'] = case_data.get('previous_divorce_suit', '')
        
        # T2 Row 9: 夫妻共同财产情况 - Cell 1
        fill_map['T2_财产事实'] = case_data.get('property_facts', '')
        
        # T2 Row 10: 夫妻共同债务情况 - Cell 1
        fill_map['T2_债务事实'] = case_data.get('debt_facts', '')
        
        # T2 Row 11: 子女直接抚养情况 - Cell 1
        fill_map['T2_抚养事实'] = case_data.get('custody_facts', '')
        
        # T2 Row 12: 子女抚养费情况 - Cell 1
        fill_map['T2_抚养费事实'] = case_data.get('child_support_facts', '')
        
        # T2 Row 13: 探望权情况 - Cell 1
        fill_map['T2_探望权事实'] = case_data.get('visitation_facts', '')
        
        # T2 Row 14: 损害赔偿/补偿/帮助情况 - Cell 1
        fill_map['T2_损害事实'] = case_data.get('damage_facts', '')
        
        # T2 Row 15: 其他事实 - Cell 1
        fill_map['T2_其他事实'] = case_data.get('other_facts', '')
        
        # ============================================================
        # T3: 请求依据 + 调解意愿
        # ============================================================
        # T3 Row 0: 请求依据 - Cell 1（由测试脚本处理）
        
        # T3 Row 5: 调解意愿（考虑先行调解）- Cell 1
        # 勾选框由get_checkbox_map处理
        
        return fill_map
    
    def get_checkbox_map(self, case_data: dict) -> dict[str, bool]:
        """
        生成勾选框映射表
        
        键名格式必须与模板中的选项文本一致：
        - 键名最后一段会被解析为选项文本
        - 例如 "T3_调解_了解" → 选项文本 = "了解"
        - 例如 "T1_无财产" → 选项文本 = "无财产"
        """
        checkbox_map = {}
        
        # ============================================================
        # T0: 原告性别
        # ============================================================
        gender = case_data.get('plaintiff_gender', '')
        checkbox_map['T0_性别_男'] = (gender == '男')
        checkbox_map['T0_性别_女'] = (gender == '女')
        
        # ============================================================
        # T0: 委托代理人
        # ============================================================
        has_agent = case_data.get('has_agent', False)
        checkbox_map['T0_代理人_有'] = has_agent
        checkbox_map['T0_代理人_无'] = not has_agent
        
        # 代理权限
        if has_agent:
            auth = case_data.get('agent_authorization', '一般授权')
            checkbox_map['T0_代理权限_一般授权'] = (auth == '一般授权')
            checkbox_map['T0_代理权限_特别授权'] = (auth == '特别授权')
        
        # ============================================================
        # T1: 被告性别
        # ============================================================
        def_gender = case_data.get('defendant_gender', '')
        checkbox_map['T1_性别_男'] = (def_gender == '男')
        checkbox_map['T1_性别_女'] = (def_gender == '女')
        
        # ============================================================
        # T1: 诉讼请求勾选项
        # ============================================================
        
        # Row 4: 夫妻共同财产
        has_property = case_data.get('has_property', False)
        checkbox_map['T1_无财产'] = not has_property
        checkbox_map['T1_有财产'] = has_property
        
        # Row 4: 财产归属 - 模板中是"归属：原告□ / 被告□ / 其他□"
        # 键名最后一段是选项文本
        checkbox_map['T1_归属_原告'] = False  # 财产归属暂不处理
        checkbox_map['T1_归属_被告'] = False
        
        # Row 5: 夫妻共同债务
        has_debt = case_data.get('has_debt', False)
        checkbox_map['T1_无债务'] = not has_debt
        checkbox_map['T1_有债务'] = has_debt
        
        # Row 6: 子女直接抚养
        has_children = case_data.get('has_children', False)
        checkbox_map['T1_无此问题'] = not has_children
        checkbox_map['T1_有此问题'] = has_children
        
        # 子女归属 - 模板中是"归属：原告□ / 被告□"
        children = case_data.get('children', [])
        if children and has_children:
            child = children[0]
            custody = child.get('custody', '')
            checkbox_map['T1_归属_原告'] = (custody == '原告')
            checkbox_map['T1_归属_被告'] = (custody == '被告')
        
        # Row 7: 子女抚养费
        has_support = case_data.get('has_child_support', False)
        checkbox_map['T1_抚养费_无此问题'] = not has_support
        checkbox_map['T1_抚养费_有此问题'] = has_support
        
        # 抚养费承担主体 - 模板中是"抚养费承担主体：原告□ / 被告□"
        payer = case_data.get('child_support_payer', '')
        checkbox_map['T1_抚养费主体_原告'] = (payer == '原告')
        checkbox_map['T1_抚养费主体_被告'] = (payer == '被告')
        
        # ============================================================
        # T2: 探望权 + 损害赔偿 + 诉讼费用 + 诉前保全
        # ============================================================
        
        # Row 0: 探望权
        has_visitation = case_data.get('has_visitation', False)
        checkbox_map['T2_探望权_无此问题'] = not has_visitation
        checkbox_map['T2_探望权_有此问题'] = has_visitation
        
        # 探望权行使主体 - 模板中是"探望权行使主体：原告□ / 被告□"
        vis_holder = case_data.get('visitation_holder', '')
        checkbox_map['T2_探望权主体_原告'] = (vis_holder == '原告')
        checkbox_map['T2_探望权主体_被告'] = (vis_holder == '被告')
        
        # Row 1: 损害赔偿/补偿/帮助
        has_damage = case_data.get('has_damage_compensation', False)
        checkbox_map['T2_损害_无此问题'] = not has_damage
        checkbox_map['T2_损害_有此问题'] = has_damage
        
        # Row 2: 诉讼费用
        claim_fee = case_data.get('claim_litigation_fee', True)
        checkbox_map['T2_诉讼费用_是'] = claim_fee
        checkbox_map['T2_诉讼费用_否'] = not claim_fee
        
        # Row 5: 诉前保全
        has_preservation = case_data.get('has_preservation', False)
        checkbox_map['T2_诉前保全_是'] = has_preservation
        checkbox_map['T2_诉前保全_否'] = not has_preservation
        
        # ============================================================
        # T3: 调解意愿
        # ============================================================
        
        # Row 3: 了解调解 - 模板中是"了解□    不了解□"
        understand = case_data.get('understand_mediation', False)
        checkbox_map['T3_调解_了解'] = understand
        checkbox_map['T3_调解_不了解'] = not understand
        
        # Row 4: 了解先行调解好处（5个勾选框都映射同一个值）
        understand_benefits = case_data.get('understand_mediation_benefits', False)
        checkbox_map['T3_先行调解_了解1'] = understand_benefits
        checkbox_map['T3_先行调解_了解2'] = understand_benefits
        checkbox_map['T3_先行调解_了解3'] = understand_benefits
        checkbox_map['T3_先行调解_了解4'] = understand_benefits
        checkbox_map['T3_先行调解_了解5'] = understand_benefits
        
        # Row 5: 是否考虑先行调解 - 模板中是"是□ / 否□ / 暂不确定□"
        consider = case_data.get('consider_mediation', '')
        checkbox_map['T3_考虑_是'] = (consider == '是')
        checkbox_map['T3_考虑_否'] = (consider == '否')
        checkbox_map['T3_考虑_暂不确定'] = (consider == '暂不确定')
        
        return checkbox_map
    
    def get_template_info(self) -> dict:
        """返回模板配置信息"""
        return {
            "template_file": "民事起诉状-离婚纠纷.docx",
            "table_configs": [
                {
                    "index": 0,
                    "type": "plaintiff",
                    "rows": [
                        {"index": 2, "name": "原告信息", "cell_index": 1, "paragraph_count": 8},
                        {"index": 3, "name": "委托代理人", "cell_index": 1, "paragraph_count": 4},
                    ],
                },
                {
                    "index": 1,
                    "type": "defendant_claims",
                    "rows": [
                        {"index": 0, "name": "被告信息", "cell_index": 1},
                        {"index": 3, "name": "离婚请求", "cell_index": 1},
                        {"index": 4, "name": "财产分割", "cell_index": 1},
                        {"index": 5, "name": "债务分担", "cell_index": 1},
                        {"index": 6, "name": "子女抚养", "cell_index": 1},
                        {"index": 7, "name": "抚养费", "cell_index": 1},
                    ],
                },
                {
                    "index": 2,
                    "type": "facts_reasons",
                    "rows": [
                        {"index": 0, "name": "探望权", "cell_index": 1},
                        {"index": 1, "name": "损害赔偿", "cell_index": 1},
                        {"index": 2, "name": "诉讼费用", "cell_index": 1},
                        {"index": 5, "name": "诉前保全", "cell_index": 1, "use_direct_fill": True},
                        {"index": 8, "name": "婚姻基本情况", "cell_index": 1, "paragraph_count": 5},
                        {"index": 9, "name": "财产事实", "cell_index": 1},
                        {"index": 10, "name": "债务事实", "cell_index": 1},
                        {"index": 11, "name": "抚养事实", "cell_index": 1},
                        {"index": 12, "name": "抚养费事实", "cell_index": 1},
                        {"index": 13, "name": "探望权事实", "cell_index": 1},
                        {"index": 14, "name": "损害事实", "cell_index": 1},
                        {"index": 15, "name": "其他事实", "cell_index": 1},
                    ],
                },
                {
                    "index": 3,
                    "type": "legal_basis_mediation",
                    "rows": [
                        {"index": 0, "name": "请求依据", "cell_index": 1},
                        {"index": 5, "name": "调解意愿", "cell_index": 1},
                    ],
                },
            ],
        }
    
    def get_evidence_rules(self) -> list:
        """返回证据规则列表"""
        return [
            {
                "name": "身份证",
                "description": "证明当事人身份",
                "required": True,
            },
            {
                "name": "结婚证",
                "description": "证明婚姻关系存在",
                "required": True,
            },
            {
                "name": "户口本",
                "description": "证明户籍及子女信息",
                "required": True,
            },
            {
                "name": "财产证明",
                "description": "房产证、车辆行驶证、银行存款证明等",
                "required": False,
            },
            {
                "name": "债务凭证",
                "description": "借条、贷款合同等",
                "required": False,
            },
            {
                "name": "子女出生证明",
                "description": "证明子女身份信息",
                "required": False,
            },
            {
                "name": "收入证明",
                "description": "工资单、纳税证明等",
                "required": False,
            },
            {
                "name": "婚姻过错证据",
                "description": "如存在法定过错情形",
                "required": False,
            },
        ]
