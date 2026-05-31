# -*- coding: utf-8 -*-
"""
劳动争议纠纷适配器

覆盖劳动争议、工伤赔偿等案由的要素式文书生成。

模板结构（民事起诉状-劳动争议纠纷.docx）：
- 表格0：当事人信息（原告自然人、委托代理人）
- 表格1：诉讼请求（8个勾选项：工资、双倍工资、加班费、年休假、社保、经济补偿、赔偿金、诉讼费）
- 表格2：事实与理由（劳动合同签订、履行、解除、工伤、仲裁等）+ 调解意愿
- 表格3：调解续

书签命名规范：{区域前缀}{序号}_{字段名}
- T0_01_原告姓名
- T1_01_被告名称
- T2_01_工资支付明细
- T3_01_劳动合同签订情况
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter


# ============================================================
# 劳动争议赔偿标准（河北省）
# ============================================================
# 注：劳动争议赔偿标准因地区、年度而异，以下为通用参考
LABOR_STANDARD = {
    # 经济补偿金标准
    "economic_compensation": {
        "description": "每工作一年支付一个月工资",
        "legal_basis": "《劳动合同法》第47条",
    },
    # 赔偿金标准
    "compensation": {
        "description": "经济补偿金的二倍",
        "legal_basis": "《劳动合同法》第87条",
    },
    # 未签订书面劳动合同双倍工资
    "double_wage": {
        "description": "自用工之日起超过一个月不满一年未签书面合同，每月支付二倍工资",
        "max_months": 11,
        "legal_basis": "《劳动合同法》第82条",
    },
    # 加班费标准
    "overtime": {
        "workday": 1.5,   # 工作日1.5倍
        "weekend": 2.0,   # 周末2倍
        "holiday": 3.0,   # 法定节假日3倍
        "legal_basis": "《劳动法》第44条",
    },
    # 未休年休假工资
    "annual_leave": {
        "description": "按日工资收入的300%支付",
        "legal_basis": "《职工带薪年休假条例》第5条",
    },
    # 社会保险
    "social_insurance": {
        "description": "用人单位未依法缴纳社保造成的损失",
        "legal_basis": "《社会保险法》第63条",
    },
}


class LaborDisputeAdapter(CaseAdapter):
    """
    劳动争议纠纷适配器
    
    支持：
    - 工资支付争议
    - 未签书面劳动合同双倍工资
    - 加班费争议
    - 未休年休假工资
    - 社会保险损失
    - 解除/终止劳动合同经济补偿
    - 违法解除劳动合同赔偿金
    - 工伤赔偿争议
    """
    
    def name(self) -> str:
        return "劳动争议纠纷"
    
    def get_schema(self) -> dict:
        return {
            "fields": [
                # === 原告（劳动者）信息 ===
                {
                    "name": "plaintiff_name",
                    "label": "原告姓名",
                    "type": "str",
                    "required": True,
                    "description": "劳动者姓名",
                },
                {
                    "name": "plaintiff_gender",
                    "label": "原告性别",
                    "type": "str",
                    "required": True,
                    "options": ["男", "女"],
                    "description": "劳动者性别",
                },
                {
                    "name": "plaintiff_birthdate",
                    "label": "原告出生日期",
                    "type": "str",
                    "required": True,
                    "description": "格式：1990年1月1日",
                },
                {
                    "name": "plaintiff_ethnicity",
                    "label": "原告民族",
                    "type": "str",
                    "required": True,
                    "description": "如：汉族",
                },
                {
                    "name": "plaintiff_id_number",
                    "label": "原告身份证号",
                    "type": "str",
                    "required": True,
                    "description": "18位身份证号码",
                },
                {
                    "name": "plaintiff_address",
                    "label": "原告住所地",
                    "type": "str",
                    "required": True,
                    "description": "户籍所在地",
                },
                {
                    "name": "plaintiff_residence",
                    "label": "原告经常居住地",
                    "type": "str",
                    "required": False,
                    "description": "经常居住地（如与住所地不同）",
                },
                {
                    "name": "plaintiff_work_unit",
                    "label": "原告工作单位",
                    "type": "str",
                    "required": False,
                    "description": "原工作单位名称",
                },
                {
                    "name": "plaintiff_position",
                    "label": "原告职务",
                    "type": "str",
                    "required": False,
                    "description": "工作岗位",
                },
                {
                    "name": "plaintiff_phone",
                    "label": "原告联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
                },
                
                # === 被告（用人单位）信息 ===
                {
                    "name": "defendant_name",
                    "label": "被告名称",
                    "type": "str",
                    "required": True,
                    "description": "用人单位全称",
                },
                {
                    "name": "defendant_address",
                    "label": "被告住所地",
                    "type": "str",
                    "required": True,
                    "description": "主要办事机构所在地",
                },
                {
                    "name": "defendant_registration_address",
                    "label": "被告注册地",
                    "type": "str",
                    "required": False,
                    "description": "注册地/登记地",
                },
                {
                    "name": "defendant_legal_rep",
                    "label": "被告法定代表人",
                    "type": "str",
                    "required": False,
                    "description": "法定代表人/负责人",
                },
                {
                    "name": "defendant_position",
                    "label": "被告法定代表人职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "defendant_phone",
                    "label": "被告联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
                },
                {
                    "name": "defendant_credit_code",
                    "label": "被告统一社会信用代码",
                    "type": "str",
                    "required": False,
                    "description": "统一社会信用代码",
                },
                {
                    "name": "defendant_company_type",
                    "label": "被告公司类型",
                    "type": "str",
                    "required": False,
                    "options": [
                        "有限责任公司", "股份有限公司", "上市公司", "其他企业法人",
                        "事业单位", "社会团体", "基金会", "社会服务机构",
                        "机关法人", "农村集体经济组织法人", "城镇农村的合作经济组织法人",
                        "基层群众性自治组织法人", "个人独资企业", "合伙企业",
                        "不具有法人资格的专业服务机构",
                    ],
                    "description": "用人单位类型",
                },
                {
                    "name": "defendant_ownership",
                    "label": "被告所有制性质",
                    "type": "str",
                    "required": False,
                    "options": ["国有", "民营", "其他"],
                    "description": "所有制性质",
                },
                
                # === 委托诉讼代理人信息 ===
                {
                    "name": "has_agent",
                    "label": "是否有委托诉讼代理人",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否有委托诉讼代理人",
                },
                {
                    "name": "agent_name",
                    "label": "代理人姓名",
                    "type": "str",
                    "required": False,
                    "description": "委托诉讼代理人姓名",
                },
                {
                    "name": "agent_unit",
                    "label": "代理人单位",
                    "type": "str",
                    "required": False,
                    "description": "代理人所在单位",
                },
                {
                    "name": "agent_position",
                    "label": "代理人职务",
                    "type": "str",
                    "required": False,
                    "description": "代理人职务",
                },
                {
                    "name": "agent_phone",
                    "label": "代理人联系电话",
                    "type": "str",
                    "required": False,
                    "description": "代理人联系电话",
                },
                {
                    "name": "agent_authorization",
                    "label": "代理权限",
                    "type": "str",
                    "required": False,
                    "options": ["一般授权", "特别授权"],
                    "description": "代理权限类型",
                },
                
                # === 送达地址信息 ===
                {
                    "name": "delivery_address",
                    "label": "送达地址",
                    "type": "str",
                    "required": False,
                    "description": "送达地址",
                },
                {
                    "name": "delivery_recipient",
                    "label": "收件人",
                    "type": "str",
                    "required": False,
                    "description": "收件人",
                },
                {
                    "name": "delivery_phone",
                    "label": "收件人电话",
                    "type": "str",
                    "required": False,
                    "description": "收件人电话",
                },
                {
                    "name": "accept_electronic_service",
                    "label": "是否接受电子送达",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否接受电子送达",
                },
                {
                    "name": "electronic_service_method",
                    "label": "电子送达方式",
                    "type": "str",
                    "required": False,
                    "description": "电子送达方式（短信/微信/邮箱等）",
                },
                
                # === 劳动关系基本信息 ===
                {
                    "name": "employment_start_date",
                    "label": "入职时间",
                    "type": "str",
                    "required": True,
                    "description": "入职日期，格式：2020年1月1日",
                },
                {
                    "name": "employment_end_date",
                    "label": "离职时间",
                    "type": "str",
                    "required": False,
                    "description": "离职日期（如已离职）",
                },
                {
                    "name": "work_position",
                    "label": "工作岗位",
                    "type": "str",
                    "required": True,
                    "description": "工作岗位",
                },
                {
                    "name": "work_location",
                    "label": "工作地点",
                    "type": "str",
                    "required": False,
                    "description": "工作地点",
                },
                {
                    "name": "contract_signed",
                    "label": "是否签订书面劳动合同",
                    "type": "bool",
                    "required": True,
                    "description": "是否签订书面劳动合同",
                },
                {
                    "name": "contract_sign_date",
                    "label": "合同签订时间",
                    "type": "str",
                    "required": False,
                    "description": "合同签订日期",
                },
                {
                    "name": "contract_end_date",
                    "label": "合同到期时间",
                    "type": "str",
                    "required": False,
                    "description": "合同到期日期",
                },
                {
                    "name": "contract_type",
                    "label": "合同类型",
                    "type": "str",
                    "required": False,
                    "options": ["固定期限", "无固定期限", "以完成一定工作任务为期限"],
                    "description": "劳动合同类型",
                },
                {
                    "name": "work_system",
                    "label": "工时制度",
                    "type": "str",
                    "required": False,
                    "options": ["标准工时", "综合计算工时", "不定时工时"],
                    "description": "工时制度",
                },
                {
                    "name": "monthly_wage",
                    "label": "月工资标准",
                    "type": "float",
                    "required": True,
                    "description": "合同约定月工资",
                },
                {
                    "name": "wage_composition",
                    "label": "工资构成",
                    "type": "str",
                    "required": False,
                    "description": "工资构成（基本工资+绩效+补贴等）",
                },
                {
                    "name": "actual_monthly_wage",
                    "label": "实际月工资",
                    "type": "float",
                    "required": False,
                    "description": "实际领取月工资",
                },
                {
                    "name": "social_insurance_start",
                    "label": "社保缴纳开始时间",
                    "type": "str",
                    "required": False,
                    "description": "社保缴纳开始时间",
                },
                {
                    "name": "social_insurance_types",
                    "label": "社保险种",
                    "type": "str",
                    "required": False,
                    "description": "社保险种",
                },
                
                # === 诉讼请求相关 ===
                {
                    "name": "claim_wage",
                    "label": "是否主张工资支付",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_wage_detail",
                    "label": "工资支付明细",
                    "type": "str",
                    "required": False,
                    "description": "工资支付明细（金额、期间）",
                },
                {
                    "name": "claim_double_wage",
                    "label": "是否主张双倍工资",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_double_wage_detail",
                    "label": "双倍工资明细",
                    "type": "str",
                    "required": False,
                    "description": "双倍工资明细（月数、金额）",
                },
                {
                    "name": "claim_overtime",
                    "label": "是否主张加班费",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_overtime_detail",
                    "label": "加班费明细",
                    "type": "str",
                    "required": False,
                    "description": "加班费明细（加班时间、金额）",
                },
                {
                    "name": "claim_annual_leave",
                    "label": "是否主张未休年休假工资",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_annual_leave_detail",
                    "label": "未休年休假工资明细",
                    "type": "str",
                    "required": False,
                    "description": "未休年休假工资明细",
                },
                {
                    "name": "claim_social_insurance_loss",
                    "label": "是否主张社保损失",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_social_insurance_loss_detail",
                    "label": "社保损失明细",
                    "type": "str",
                    "required": False,
                    "description": "社保损失明细",
                },
                {
                    "name": "claim_economic_compensation",
                    "label": "是否主张经济补偿",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_economic_compensation_detail",
                    "label": "经济补偿明细",
                    "type": "str",
                    "required": False,
                    "description": "经济补偿明细（工作年限、月数、金额）",
                },
                {
                    "name": "claim_compensation",
                    "label": "是否主张赔偿金",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "claim_compensation_detail",
                    "label": "赔偿金明细",
                    "type": "str",
                    "required": False,
                    "description": "赔偿金明细",
                },
                {
                    "name": "claim_litigation_fee",
                    "label": "是否主张诉讼费用",
                    "type": "bool",
                    "required": False,
                    "default": True,
                },
                {
                    "name": "other_claims",
                    "label": "其他诉讼请求",
                    "type": "str",
                    "required": False,
                    "description": "其他诉讼请求",
                },
                {
                    "name": "total_amount",
                    "label": "标的总额",
                    "type": "float",
                    "required": False,
                    "description": "诉讼标的总额",
                },
                
                # === 诉前保全 ===
                {
                    "name": "has_preservation",
                    "label": "是否已申请诉前保全",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "preservation_court",
                    "label": "保全法院",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "preservation_time",
                    "label": "保全时间",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "preservation_case_number",
                    "label": "保全案号",
                    "type": "str",
                    "required": False,
                },
                
                # === 事实与理由 ===
                {
                    "name": "contract_signing_info",
                    "label": "劳动合同签订情况",
                    "type": "str",
                    "required": False,
                    "description": "合同主体、签订时间、地点、合同名称等",
                },
                {
                    "name": "contract_performance_info",
                    "label": "劳动合同履行情况",
                    "type": "str",
                    "required": False,
                    "description": "入职时间、工作岗位、工资构成、社保、加班等",
                },
                {
                    "name": "termination_info",
                    "label": "解除/终止劳动关系情况",
                    "type": "str",
                    "required": False,
                    "description": "解除原因、经济补偿/赔偿金数额等",
                },
                {
                    "name": "work_injury_info",
                    "label": "工伤情况",
                    "type": "str",
                    "required": False,
                    "description": "工伤时间、认定情况、伤残等级、费用等",
                },
                {
                    "name": "arbitration_info",
                    "label": "劳动仲裁情况",
                    "type": "str",
                    "required": False,
                    "description": "仲裁时间、请求、结果等",
                },
                {
                    "name": "other_info",
                    "label": "其他相关情况",
                    "type": "str",
                    "required": False,
                    "description": "如是否农民工等",
                },
                {
                    "name": "legal_basis",
                    "label": "诉请依据",
                    "type": "str",
                    "required": False,
                    "description": "法律及司法解释规定",
                },
                
                # === 调解意愿 ===
                {
                    "name": "understand_mediation",
                    "label": "是否了解调解",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "understand_mediation_benefits",
                    "label": "是否了解先行调解好处",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "consider_mediation",
                    "label": "是否考虑先行调解",
                    "type": "str",
                    "required": False,
                    "options": ["是", "否", "暂不确定"],
                },
                
                # === 工伤相关（如适用） ===
                {
                    "name": "work_injury_date",
                    "label": "工伤发生时间",
                    "type": "str",
                    "required": False,
                    "description": "工伤发生日期",
                },
                {
                    "name": "work_injury_recognition",
                    "label": "工伤认定情况",
                    "type": "str",
                    "required": False,
                    "description": "是否已认定工伤",
                },
                {
                    "name": "disability_level",
                    "label": "伤残等级",
                    "type": "str",
                    "required": False,
                    "description": "劳动能力鉴定伤残等级",
                },
                {
                    "name": "work_injury_medical_expenses",
                    "label": "工伤医疗费",
                    "type": "float",
                    "required": False,
                    "description": "工伤医疗费金额",
                },
                {
                    "name": "hospitalization_start",
                    "label": "住院开始时间",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "hospitalization_end",
                    "label": "住院结束时间",
                    "type": "str",
                    "required": False,
                },
            ],
            "party_fields": {
                "plaintiff": [
                    {"name": "姓名", "type": "str", "required": True},
                    {"name": "性别", "type": "str", "required": True, "options": ["男", "女"]},
                    {"name": "出生日期", "type": "str", "required": True},
                    {"name": "民族", "type": "str", "required": True},
                    {"name": "身份证号", "type": "str", "required": True},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "经常居住地", "type": "str", "required": False},
                    {"name": "工作单位", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "联系电话", "type": "str", "required": False},
                ],
                "defendant_company": [
                    {"name": "名称", "type": "str", "required": True},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "注册地", "type": "str", "required": False},
                    {"name": "法定代表人", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "联系电话", "type": "str", "required": False},
                    {"name": "统一社会信用代码", "type": "str", "required": False},
                    {"name": "类型", "type": "str", "required": False},
                    {"name": "所有制性质", "type": "str", "required": False},
                ],
            },
        }
    
    def calculate(self, case_data: dict) -> dict:
        """
        劳动争议赔偿计算
        
        注：劳动争议赔偿计算复杂，涉及多种情形。此方法提供基础计算框架，
        具体金额需根据案件实际情况确定。
        """
        items = []
        total = 0.0
        
        # 工资支付
        if case_data.get('claim_wage'):
            amount = case_data.get('claim_wage_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "工资支付",
                    "amount": amount,
                    "formula": case_data.get('claim_wage_detail', ''),
                    "legal_basis": "《劳动合同法》第30条",
                })
                total += amount
        
        # 未签书面劳动合同双倍工资
        if case_data.get('claim_double_wage'):
            amount = case_data.get('claim_double_wage_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "未签书面劳动合同双倍工资",
                    "amount": amount,
                    "formula": case_data.get('claim_double_wage_detail', ''),
                    "legal_basis": "《劳动合同法》第82条",
                })
                total += amount
        
        # 加班费
        if case_data.get('claim_overtime'):
            amount = case_data.get('claim_overtime_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "加班费",
                    "amount": amount,
                    "formula": case_data.get('claim_overtime_detail', ''),
                    "legal_basis": "《劳动法》第44条",
                })
                total += amount
        
        # 未休年休假工资
        if case_data.get('claim_annual_leave'):
            amount = case_data.get('claim_annual_leave_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "未休年休假工资",
                    "amount": amount,
                    "formula": case_data.get('claim_annual_leave_detail', ''),
                    "legal_basis": "《职工带薪年休假条例》第5条",
                })
                total += amount
        
        # 社保损失
        if case_data.get('claim_social_insurance_loss'):
            amount = case_data.get('claim_social_insurance_loss_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "未依法缴纳社会保险费造成的经济损失",
                    "amount": amount,
                    "formula": case_data.get('claim_social_insurance_loss_detail', ''),
                    "legal_basis": "《社会保险法》第63条",
                })
                total += amount
        
        # 经济补偿
        if case_data.get('claim_economic_compensation'):
            amount = case_data.get('claim_economic_compensation_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "解除劳动合同经济补偿",
                    "amount": amount,
                    "formula": case_data.get('claim_economic_compensation_detail', ''),
                    "legal_basis": "《劳动合同法》第46条、第47条",
                })
                total += amount
        
        # 赔偿金
        if case_data.get('claim_compensation'):
            amount = case_data.get('claim_compensation_amount', 0.0)
            if amount > 0:
                items.append({
                    "name": "违法解除劳动合同赔偿金",
                    "amount": amount,
                    "formula": case_data.get('claim_compensation_detail', ''),
                    "legal_basis": "《劳动合同法》第87条",
                })
                total += amount
        
        # 工伤相关费用
        work_injury_total = 0.0
        if case_data.get('work_injury_medical_expenses'):
            amount = case_data['work_injury_medical_expenses']
            items.append({
                "name": "工伤医疗费",
                "amount": amount,
                "formula": "工伤医疗费",
                "legal_basis": "《工伤保险条例》第30条",
            })
            work_injury_total += amount
        
        if work_injury_total > 0:
            total += work_injury_total
        
        return {
            "items": items,
            "total": total,
            "work_injury_total": work_injury_total,
        }
    
    def build_fill_map(self, case_data: dict, calc_result: dict) -> dict[str, str]:
        """生成书签填充映射表"""
        fill_map = {}
        
        # === 表格0：当事人信息 ===
        # 原告信息
        # 原告信息 - 与模板段落结构匹配
        # 段落0: 姓名
        fill_map['T0_01_原告姓名'] = case_data.get('plaintiff_name', '')
        # 段落1: 性别（勾选框处理）
        # 段落2: 出生日期+民族（需要特殊处理）
        # 段落3: 工作单位+职务+联系电话（需要特殊处理）
        # 段落4: 住所地
        fill_map['T0_02_原告住所地'] = case_data.get('plaintiff_address', '')
        # 段落5: 经常居住地
        fill_map['T0_03_原告经常居住地'] = case_data.get('plaintiff_residence', '')
        # 段落6: 证件类型
        fill_map['T0_04_原告证件类型'] = '居民身份证'
        # 段落7: 证件号码
        fill_map['T0_05_原告证件号码'] = case_data.get('plaintiff_id_number', '')
        
        # 多字段段落需要特殊处理（在测试脚本中手动插入书签）
        # 段落2: 出生日期+民族
        # 段落3: 工作单位+职务+联系电话
        # 这些字段需要在测试脚本中手动处理
        # 委托代理人
        if case_data.get('has_agent'):
            fill_map['T0_12_代理人姓名'] = case_data.get('agent_name', '')
            fill_map['T0_13_代理人单位'] = case_data.get('agent_unit', '')
            fill_map['T0_14_代理人职务'] = case_data.get('agent_position', '')
            fill_map['T0_15_代理人联系电话'] = case_data.get('agent_phone', '')
        
        # 送达地址
        fill_map['T0_16_送达地址'] = case_data.get('delivery_address', '')
        fill_map['T0_17_收件人'] = case_data.get('delivery_recipient', '')
        fill_map['T0_18_收件人电话'] = case_data.get('delivery_phone', '')
        
        # === 表格1：被告信息 + 诉讼请求 ===
        # 被告信息
        fill_map['T1_01_被告名称'] = case_data.get('defendant_name', '')
        fill_map['T1_02_被告住所地'] = case_data.get('defendant_address', '')
        fill_map['T1_03_被告注册地'] = case_data.get('defendant_registration_address', '')
        fill_map['T1_04_被告法定代表人'] = case_data.get('defendant_legal_rep', '')
        fill_map['T1_05_被告法定代表人职务'] = case_data.get('defendant_position', '')
        fill_map['T1_06_被告联系电话'] = case_data.get('defendant_phone', '')
        fill_map['T1_07_被告统一社会信用代码'] = case_data.get('defendant_credit_code', '')
        
        # 诉讼请求明细
        if case_data.get('claim_wage'):
            fill_map['T1_08_工资支付明细'] = case_data.get('claim_wage_detail', '')
        if case_data.get('claim_double_wage'):
            fill_map['T1_09_双倍工资明细'] = case_data.get('claim_double_wage_detail', '')
        if case_data.get('claim_overtime'):
            fill_map['T1_10_加班费明细'] = case_data.get('claim_overtime_detail', '')
        if case_data.get('claim_annual_leave'):
            fill_map['T1_11_未休年休假工资明细'] = case_data.get('claim_annual_leave_detail', '')
        if case_data.get('claim_social_insurance_loss'):
            fill_map['T1_12_社保损失明细'] = case_data.get('claim_social_insurance_loss_detail', '')
        if case_data.get('claim_economic_compensation'):
            fill_map['T1_13_经济补偿明细'] = case_data.get('claim_economic_compensation_detail', '')
        if case_data.get('claim_compensation'):
            fill_map['T1_14_赔偿金明细'] = case_data.get('claim_compensation_detail', '')
        if case_data.get('other_claims'):
            fill_map['T1_15_其他诉讼请求'] = case_data.get('other_claims', '')
        if case_data.get('total_amount'):
            fill_map['T1_16_标的总额'] = str(int(case_data['total_amount']))
        
        # === 表格2：事实与理由 ===
        # 诉前保全
        if case_data.get('has_preservation'):
            fill_map['T2_01_保全法院'] = case_data.get('preservation_court', '')
            fill_map['T2_02_保全时间'] = case_data.get('preservation_time', '')
            fill_map['T2_03_保全案号'] = case_data.get('preservation_case_number', '')
        
        # 事实与理由
        fill_map['T2_04_劳动合同签订情况'] = case_data.get('contract_signing_info', '')
        fill_map['T2_05_劳动合同履行情况'] = case_data.get('contract_performance_info', '')
        fill_map['T2_06_解除终止劳动关系情况'] = case_data.get('termination_info', '')
        fill_map['T2_07_工伤情况'] = case_data.get('work_injury_info', '')
        fill_map['T2_08_劳动仲裁情况'] = case_data.get('arbitration_info', '')
        fill_map['T2_09_其他相关情况'] = case_data.get('other_info', '')
        fill_map['T2_10_诉请依据'] = case_data.get('legal_basis', '')
        
        return fill_map
    
    def get_checkbox_map(self, case_data: dict) -> dict[str, bool]:
        """生成勾选框映射表"""
        checkbox_map = {}
        
        # === 表格0：原告性别 ===
        gender = case_data.get('plaintiff_gender', '')
        checkbox_map['T0_性别_男'] = (gender == '男')
        checkbox_map['T0_性别_女'] = (gender == '女')
        
        # 委托代理人
        has_agent = case_data.get('has_agent', False)
        checkbox_map['T0_代理人_有'] = has_agent
        checkbox_map['T0_代理人_无'] = not has_agent
        
        # 代理权限
        if has_agent:
            auth = case_data.get('agent_authorization', '')
            checkbox_map['T0_代理权限_一般授权'] = (auth == '一般授权')
            checkbox_map['T0_代理权限_特别授权'] = (auth == '特别授权')
        
        # 电子送达
        accept_electronic = case_data.get('accept_electronic_service', False)
        checkbox_map['T0_电子送达_是'] = accept_electronic
        checkbox_map['T0_电子送达_否'] = not accept_electronic
        
        # === 表格1：被告公司类型 ===
        company_type = case_data.get('defendant_company_type', '')
        company_types = [
            '有限责任公司', '股份有限公司', '上市公司', '其他企业法人',
            '事业单位', '社会团体', '基金会', '社会服务机构',
            '机关法人', '农村集体经济组织法人', '城镇农村的合作经济组织法人',
            '基层群众性自治组织法人', '个人独资企业', '合伙企业',
            '不具有法人资格的专业服务机构',
        ]
        for ct in company_types:
            checkbox_map[f'T1_公司类型_{ct}'] = (company_type == ct)
        
        # 所有制性质
        ownership = case_data.get('defendant_ownership', '')
        checkbox_map['T1_所有制_国有'] = (ownership == '国有')
        checkbox_map['T1_所有制_民营'] = (ownership == '民营')
        checkbox_map['T1_所有制_其他'] = (ownership == '其他')
        
        # 诉讼请求勾选项
        claims = [
            ('claim_wage', '工资支付'),
            ('claim_double_wage', '双倍工资'),
            ('claim_overtime', '加班费'),
            ('claim_annual_leave', '未休年休假工资'),
            ('claim_social_insurance_loss', '社保损失'),
            ('claim_economic_compensation', '经济补偿'),
            ('claim_compensation', '赔偿金'),
            ('claim_litigation_fee', '诉讼费用'),
        ]
        for claim_key, claim_name in claims:
            value = case_data.get(claim_key, False)
            checkbox_map[f'T1_{claim_name}_是'] = value
            checkbox_map[f'T1_{claim_name}_否'] = not value
        
        # === 表格2：诉前保全 ===
        has_preservation = case_data.get('has_preservation', False)
        checkbox_map['T2_诉前保全_是'] = has_preservation
        checkbox_map['T2_诉前保全_否'] = not has_preservation
        
        # 调解意愿
        understand_mediation = case_data.get('understand_mediation', False)
        checkbox_map['T2_了解调解_了解'] = understand_mediation
        checkbox_map['T2_了解调解_不了解'] = not understand_mediation
        
        understand_benefits = case_data.get('understand_mediation_benefits', False)
        checkbox_map['T2_了解先行调解_了解'] = understand_benefits
        checkbox_map['T2_了解先行调解_不了解'] = not understand_benefits
        
        # 是否考虑先行调解
        consider = case_data.get('consider_mediation', '')
        checkbox_map['T2_考虑调解_是'] = (consider == '是')
        checkbox_map['T2_考虑调解_否'] = (consider == '否')
        checkbox_map['T2_考虑调解_暂不确定'] = (consider == '暂不确定')
        
        return checkbox_map
    
    def get_template_info(self) -> dict:
        """返回模板配置信息"""
        return {
            "template_file": "民事起诉状-劳动争议纠纷.docx",
            "table_configs": [
                {
                    "index": 0,
                    "type": "plaintiff",
                    "fields": [
                        "原告姓名", "原告性别", "原告出生日期", "原告民族",
                        "原告工作单位", "原告职务", "原告联系电话",
                        "原告住所地", "原告经常居住地", "原告证件类型", "原告证件号码",
                    ],
                },
                {
                    "index": 1,
                    "type": "defendant",
                    "fields": [
                        "被告名称", "被告住所地", "被告注册地",
                        "被告法定代表人", "被告法定代表人职务", "被告联系电话",
                        "被告统一社会信用代码",
                    ],
                },
                {
                    "index": 2,
                    "type": "claims",
                    "fields": [
                        "工资支付明细", "双倍工资明细", "加班费明细",
                        "未休年休假工资明细", "社保损失明细", "经济补偿明细",
                        "赔偿金明细", "其他诉讼请求", "标的总额",
                    ],
                },
                {
                    "index": 3,
                    "type": "facts",
                    "fields": [
                        "保全法院", "保全时间", "保全案号",
                        "劳动合同签订情况", "劳动合同履行情况",
                        "解除终止劳动关系情况", "工伤情况",
                        "劳动仲裁情况", "其他相关情况", "诉请依据",
                    ],
                },
            ],
        }
    
    def get_evidence_rules(self) -> list:
        """返回证据规则列表"""
        return [
            {
                "name": "劳动合同",
                "description": "证明劳动关系存在",
                "required": True,
                "legal_basis": "《劳动合同法》第10条",
            },
            {
                "name": "工资支付凭证",
                "description": "工资条、银行流水等",
                "required": True,
                "legal_basis": "《工资支付暂行规定》第6条",
            },
            {
                "name": "解除/终止劳动合同证明",
                "description": "证明劳动关系解除",
                "required": False,
                "legal_basis": "《劳动合同法》第50条",
            },
            {
                "name": "社保缴纳记录",
                "description": "证明社保缴纳情况",
                "required": False,
                "legal_basis": "《社会保险法》第60条",
            },
            {
                "name": "工伤认定决定书",
                "description": "证明工伤事实",
                "required": False,
                "legal_basis": "《工伤保险条例》第20条",
            },
            {
                "name": "劳动能力鉴定结论",
                "description": "证明伤残等级",
                "required": False,
                "legal_basis": "《工伤保险条例》第25条",
            },
            {
                "name": "劳动仲裁裁决书",
                "description": "证明仲裁前置",
                "required": True,
                "legal_basis": "《劳动争议调解仲裁法》第5条",
            },
        ]