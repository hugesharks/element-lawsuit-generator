# -*- coding: utf-8 -*-
"""
民间借贷纠纷适配器

覆盖民间借贷纠纷的要素式文书生成。

模板结构（民事起诉状-民间借贷纠纷.docx）：
- 表格0：当事人信息（原告自然人、原告法人）
- 表格1：委托代理人 + 被告（自然人/法人）+ 第三人（自然人/法人）
- 表格2：诉讼请求（本金/利息/提前还款/担保/费用）+ 约定管辖/诉前保全
- 表格3：事实与理由 part 1（合同签订/借款金额/期限/利率/还款方式/逾期/担保）
- 表格4：事实与理由 part 2（抵押质押登记/保证合同/其他担保/法律依据）+ 调解意愿
- 表格5：是否考虑先行调解

书签命名规范：{表前缀}_{序号}_{字段名}
- T0_01_原告姓名
- T1_01_被告名称
- T2_01_本金
- T3_01_合同签订情况
- T4_01_抵押质押登记
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter


# ============================================================
# 民间借贷纠纷利息计算标准
# ============================================================
LENDING_INTEREST_RULES = {
    # 2020年8月20日之前的利率（不超过年利率24%）
    "before_aug_2020": {
        "description": "约定利率不超过年利率24%",
        "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2015) 第26条",
    },
    # 2020年8月20日之后的利率（LPR的4倍）
    "after_aug_2020": {
        "description": "合同成立时一年期贷款市场报价利率四倍",
        "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订) 第25条",
    },
    # 逾期利息
    "overdue": {
        "description": "逾期利息按合同约定或借款期间利率计算",
        "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订) 第29条",
    },
}


class CivilLendingAdapter(CaseAdapter):
    """
    民间借贷纠纷适配器
    
    支持：
    - 自然人之间借贷
    - 自然人与法人或其他组织之间借贷
    - 本金、利息、逾期利息计算
    - 物的担保（抵押、质押）
    - 保证担保（一般保证、连带责任保证）
    - 诉前保全
    """
    
    def name(self) -> str:
        return "民间借贷纠纷"
    
    def get_template_name(self) -> str:
        return "民事起诉状-民间借贷纠纷.docx"
    
    def get_schema(self) -> dict:
        return {
            "fields": [
                # === 原告（出借人）信息 ===
                {
                    "name": "plaintiff_type",
                    "label": "原告类型",
                    "type": "str",
                    "required": True,
                    "options": ["自然人", "法人"],
                    "description": "原告是自然人还是法人",
                },
                {
                    "name": "plaintiff_name",
                    "label": "原告姓名/名称",
                    "type": "str",
                    "required": True,
                    "description": "自然人填写姓名，法人填写名称",
                },
                {
                    "name": "plaintiff_gender",
                    "label": "原告性别",
                    "type": "str",
                    "required": False,
                    "options": ["男", "女"],
                    "description": "仅自然人需要",
                },
                {
                    "name": "plaintiff_birthdate",
                    "label": "原告出生日期",
                    "type": "str",
                    "required": False,
                    "description": "格式：1990年1月1日",
                },
                {
                    "name": "plaintiff_work_unit",
                    "label": "原告工作单位",
                    "type": "str",
                    "required": False,
                    "description": "工作单位",
                },
                {
                    "name": "plaintiff_position",
                    "label": "原告职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "plaintiff_address",
                    "label": "原告住所地",
                    "type": "str",
                    "required": True,
                    "description": "户籍所在地或主要办事机构所在地",
                },
                {
                    "name": "plaintiff_residence",
                    "label": "原告经常居住地",
                    "type": "str",
                    "required": False,
                    "description": "经常居住地（如与住所地不同）",
                },
                {
                    "name": "plaintiff_id_type",
                    "label": "原告证件类型",
                    "type": "str",
                    "required": True,
                    "default": "居民身份证",
                    "description": "证件类型",
                },
                {
                    "name": "plaintiff_id_number",
                    "label": "原告证件号码",
                    "type": "str",
                    "required": True,
                    "description": "身份证号或统一社会信用代码",
                },
                # 原告法人信息
                {
                    "name": "plaintiff_registration_address",
                    "label": "原告注册地",
                    "type": "str",
                    "required": False,
                    "description": "注册地/登记地",
                },
                {
                    "name": "plaintiff_legal_rep",
                    "label": "原告法定代表人",
                    "type": "str",
                    "required": False,
                    "description": "法定代表人/负责人",
                },
                {
                    "name": "plaintiff_legal_rep_position",
                    "label": "原告法定代表人职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "plaintiff_phone",
                    "label": "原告联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
                },
                {
                    "name": "plaintiff_company_type",
                    "label": "原告公司类型",
                    "type": "str",
                    "required": False,
                    "description": "公司类型",
                },
                {
                    "name": "plaintiff_ownership",
                    "label": "原告所有制性质",
                    "type": "str",
                    "required": False,
                    "options": ["国有", "民营", "其他"],
                    "description": "所有制性质",
                },
                
                # === 被告（借款人）信息 ===
                {
                    "name": "defendant_type",
                    "label": "被告类型",
                    "type": "str",
                    "required": True,
                    "options": ["自然人", "法人"],
                    "description": "被告是自然人还是法人",
                },
                {
                    "name": "defendant_name",
                    "label": "被告姓名/名称",
                    "type": "str",
                    "required": True,
                    "description": "自然人填写姓名，法人填写名称",
                },
                {
                    "name": "defendant_gender",
                    "label": "被告性别",
                    "type": "str",
                    "required": False,
                    "options": ["男", "女"],
                    "description": "仅自然人需要",
                },
                {
                    "name": "defendant_birthdate",
                    "label": "被告出生日期",
                    "type": "str",
                    "required": False,
                    "description": "格式：1990年1月1日",
                },
                {
                    "name": "defendant_work_unit",
                    "label": "被告工作单位",
                    "type": "str",
                    "required": False,
                    "description": "工作单位",
                },
                {
                    "name": "defendant_position",
                    "label": "被告职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "defendant_address",
                    "label": "被告住所地",
                    "type": "str",
                    "required": True,
                    "description": "户籍所在地或主要办事机构所在地",
                },
                {
                    "name": "defendant_residence",
                    "label": "被告经常居住地",
                    "type": "str",
                    "required": False,
                    "description": "经常居住地（如与住所地不同）",
                },
                {
                    "name": "defendant_id_type",
                    "label": "被告证件类型",
                    "type": "str",
                    "required": True,
                    "default": "居民身份证",
                    "description": "证件类型",
                },
                {
                    "name": "defendant_id_number",
                    "label": "被告证件号码",
                    "type": "str",
                    "required": True,
                    "description": "身份证号或统一社会信用代码",
                },
                # 被告法人信息
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
                    "name": "defendant_legal_rep_position",
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
                    "name": "defendant_company_type",
                    "label": "被告公司类型",
                    "type": "str",
                    "required": False,
                    "description": "公司类型",
                },
                {
                    "name": "defendant_ownership",
                    "label": "被告所有制性质",
                    "type": "str",
                    "required": False,
                    "options": ["国有", "民营", "其他"],
                    "description": "所有制性质",
                },
                
                # === 第三人信息 ===
                {
                    "name": "has_third_party",
                    "label": "是否有第三人",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否有第三人",
                },
                {
                    "name": "third_party_type",
                    "label": "第三人类型",
                    "type": "str",
                    "required": False,
                    "options": ["自然人", "法人"],
                    "description": "第三人类型",
                },
                {
                    "name": "third_party_name",
                    "label": "第三人姓名/名称",
                    "type": "str",
                    "required": False,
                    "description": "第三人姓名或名称",
                },
                {
                    "name": "third_party_gender",
                    "label": "第三人性别",
                    "type": "str",
                    "required": False,
                    "options": ["男", "女"],
                    "description": "仅自然人需要",
                },
                {
                    "name": "third_party_address",
                    "label": "第三人住所地",
                    "type": "str",
                    "required": False,
                    "description": "住所地",
                },
                {
                    "name": "third_party_id_number",
                    "label": "第三人证件号码",
                    "type": "str",
                    "required": False,
                    "description": "证件号码",
                },
                
                # === 委托诉讼代理人 ===
                {
                    "name": "has_agent",
                    "label": "是否有委托诉讼代理人",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "agent_name",
                    "label": "代理人姓名",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "agent_unit",
                    "label": "代理人单位",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "agent_position",
                    "label": "代理人职务",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "agent_phone",
                    "label": "代理人联系电话",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "agent_authorization",
                    "label": "代理权限",
                    "type": "str",
                    "required": False,
                    "options": ["一般授权", "特别授权"],
                },
                
                # === 合同信息 ===
                {
                    "name": "contract_name",
                    "label": "合同名称",
                    "type": "str",
                    "required": False,
                    "description": "借款合同名称",
                },
                {
                    "name": "contract_number",
                    "label": "合同编号",
                    "type": "str",
                    "required": False,
                    "description": "借款合同编号",
                },
                {
                    "name": "contract_sign_date",
                    "label": "合同签订日期",
                    "type": "str",
                    "required": True,
                    "description": "合同签订日期",
                },
                {
                    "name": "contract_sign_place",
                    "label": "合同签订地点",
                    "type": "str",
                    "required": False,
                    "description": "合同签订地点",
                },
                
                # === 借款信息 ===
                {
                    "name": "lender_name",
                    "label": "出借人",
                    "type": "str",
                    "required": True,
                    "description": "出借人姓名/名称",
                },
                {
                    "name": "borrower_name",
                    "label": "借款人",
                    "type": "str",
                    "required": True,
                    "description": "借款人姓名/名称",
                },
                {
                    "name": "loan_amount_agreed",
                    "label": "约定借款金额",
                    "type": "float",
                    "required": True,
                    "description": "合同约定的借款金额",
                },
                {
                    "name": "loan_amount_actual",
                    "label": "实际提供金额",
                    "type": "float",
                    "required": True,
                    "description": "实际提供的借款金额",
                },
                {
                    "name": "loan_provision_method",
                    "label": "提供方式",
                    "type": "str",
                    "required": True,
                    "options": ["现金", "转账", "其他"],
                    "description": "借款提供方式",
                },
                {
                    "name": "loan_provision_method_other",
                    "label": "其他提供方式",
                    "type": "str",
                    "required": False,
                    "description": "其他方式的具体说明",
                },
                
                # === 借款期限 ===
                {
                    "name": "loan_start_date",
                    "label": "借款起始日期",
                    "type": "str",
                    "required": True,
                    "description": "借款起始日期",
                },
                {
                    "name": "loan_end_date",
                    "label": "借款到期日期",
                    "type": "str",
                    "required": False,
                    "description": "借款到期日期",
                },
                {
                    "name": "is_matured",
                    "label": "是否到期",
                    "type": "bool",
                    "required": True,
                    "description": "借款是否已到期",
                },
                
                # === 利率 ===
                {
                    "name": "interest_rate",
                    "label": "年利率",
                    "type": "float",
                    "required": True,
                    "description": "合同约定的年利率（百分比，如12表示12%）",
                },
                {
                    "name": "interest_rate_type",
                    "label": "利率类型",
                    "type": "str",
                    "required": False,
                    "options": ["年", "季", "月"],
                    "default": "年",
                    "description": "利率是年利率、季利率还是月利率",
                },
                {
                    "name": "contract_article_number",
                    "label": "合同条款编号",
                    "type": "str",
                    "required": False,
                    "description": "合同条款编号",
                },
                
                # === 实际提供时间 ===
                {
                    "name": "actual_provision_date",
                    "label": "实际提供日期",
                    "type": "str",
                    "required": True,
                    "description": "实际提供借款的日期",
                },
                {
                    "name": "actual_provision_amount",
                    "label": "实际提供金额",
                    "type": "float",
                    "required": True,
                    "description": "实际提供金额",
                },
                
                # === 还款方式 ===
                {
                    "name": "repayment_method",
                    "label": "还款方式",
                    "type": "str",
                    "required": True,
                    "options": [
                        "到期一次性还本付息",
                        "按月计息、到期一次性还本",
                        "按季计息、到期一次性还本",
                        "按年计息、到期一次性还本",
                        "其他",
                    ],
                    "description": "还款方式",
                },
                {
                    "name": "repayment_method_other",
                    "label": "其他还款方式",
                    "type": "str",
                    "required": False,
                    "description": "其他还款方式说明",
                },
                
                # === 还款情况 ===
                {
                    "name": "repaid_principal",
                    "label": "已还本金",
                    "type": "float",
                    "required": False,
                    "default": 0.0,
                    "description": "已偿还的本金金额",
                },
                {
                    "name": "repaid_interest",
                    "label": "已还利息",
                    "type": "float",
                    "required": False,
                    "default": 0.0,
                    "description": "已偿还的利息金额",
                },
                {
                    "name": "interest_paid_to_date",
                    "label": "还息截止日期",
                    "type": "str",
                    "required": False,
                    "description": "利息支付截止日期",
                },
                
                # === 逾期 ===
                {
                    "name": "is_overdue",
                    "label": "是否逾期",
                    "type": "bool",
                    "required": True,
                    "description": "是否存在逾期还款",
                },
                {
                    "name": "overdue_start_date",
                    "label": "逾期起始日期",
                    "type": "str",
                    "required": False,
                    "description": "逾期开始日期",
                },
                {
                    "name": "overdue_months",
                    "label": "已逾期月数",
                    "type": "int",
                    "required": False,
                    "description": "至今已逾期月数",
                },
                
                # === 物的担保 ===
                {
                    "name": "has_property_guarantee",
                    "label": "是否签订物的担保合同",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否签订抵押、质押合同",
                },
                {
                    "name": "property_guarantee_sign_date",
                    "label": "担保合同签订时间",
                    "type": "str",
                    "required": False,
                    "description": "物的担保合同签订时间",
                },
                {
                    "name": "guarantor_name",
                    "label": "担保人",
                    "type": "str",
                    "required": False,
                    "description": "担保人姓名/名称",
                },
                {
                    "name": "guarantee_property",
                    "label": "担保物",
                    "type": "str",
                    "required": False,
                    "description": "担保物描述",
                },
                {
                    "name": "has_maximum_guarantee",
                    "label": "是否最高额担保",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否最高额抵押、质押",
                },
                {
                    "name": "guarantee_determination_date",
                    "label": "担保债权确定时间",
                    "type": "str",
                    "required": False,
                    "description": "担保债权确定时间",
                },
                {
                    "name": "guarantee_amount",
                    "label": "担保额度",
                    "type": "float",
                    "required": False,
                    "description": "担保额度",
                },
                
                # === 抵押质押登记 ===
                {
                    "name": "has_registration",
                    "label": "是否办理登记",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否办理抵押、质押登记",
                },
                {
                    "name": "registration_type",
                    "label": "登记类型",
                    "type": "str",
                    "required": False,
                    "options": ["正式登记", "预告登记"],
                    "description": "正式登记或预告登记",
                },
                
                # === 保证担保 ===
                {
                    "name": "has_surety",
                    "label": "是否签订保证合同",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是否签订保证合同",
                },
                {
                    "name": "surety_sign_date",
                    "label": "保证合同签订时间",
                    "type": "str",
                    "required": False,
                    "description": "保证合同签订时间",
                },
                {
                    "name": "surety_person",
                    "label": "保证人",
                    "type": "str",
                    "required": False,
                    "description": "保证人姓名/名称",
                },
                {
                    "name": "surety_main_content",
                    "label": "主要内容",
                    "type": "str",
                    "required": False,
                    "description": "保证合同主要内容",
                },
                {
                    "name": "surety_type",
                    "label": "保证方式",
                    "type": "str",
                    "required": False,
                    "options": ["一般保证", "连带责任保证"],
                    "description": "一般保证或连带责任保证",
                },
                
                # === 其他担保 ===
                {
                    "name": "has_other_guarantee",
                    "label": "是否有其他担保方式",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "other_guarantee_form",
                    "label": "其他担保形式",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "other_guarantee_sign_date",
                    "label": "其他担保签订时间",
                    "type": "str",
                    "required": False,
                },
                
                # === 其他说明 ===
                {
                    "name": "other_matters",
                    "label": "其他需要说明的内容",
                    "type": "str",
                    "required": False,
                    "description": "其他需要说明的内容",
                },
                
                # === 诉讼请求 ===
                {
                    "name": "principal_calc_date",
                    "label": "本金计算截止日期",
                    "type": "str",
                    "required": True,
                    "description": "本金计算截止日期",
                },
                {
                    "name": "principal_amount",
                    "label": "本金金额",
                    "type": "float",
                    "required": True,
                    "description": "尚欠本金金额",
                },
                {
                    "name": "interest_calc_date",
                    "label": "利息计算截止日期",
                    "type": "str",
                    "required": True,
                    "description": "利息计算截止日期",
                },
                {
                    "name": "interest_amount",
                    "label": "利息金额",
                    "type": "float",
                    "required": True,
                    "description": "尚欠利息金额",
                },
                {
                    "name": "request_interest_to_actual",
                    "label": "是否请求支付至实际清偿日",
                    "type": "bool",
                    "required": False,
                    "default": True,
                    "description": "是否请求支付至实际清偿之日",
                },
                
                # === 提前还款/解除合同 ===
                {
                    "name": "request_acceleration",
                    "label": "是否要求提前还款（加速到期）",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "request_rescission",
                    "label": "是否要求解除合同",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                
                # === 担保权利 ===
                {
                    "name": "request_guarantee_rights",
                    "label": "是否主张担保权利",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "guarantee_rights_content",
                    "label": "担保权利内容",
                    "type": "str",
                    "required": False,
                },
                
                # === 债权费用 ===
                {
                    "name": "request_creditor_expenses",
                    "label": "是否主张实现债权的费用",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "creditor_expenses_detail",
                    "label": "债权费用明细",
                    "type": "str",
                    "required": False,
                },
                
                # === 诉讼费用 ===
                {
                    "name": "request_litigation_fee",
                    "label": "是否主张诉讼费用",
                    "type": "bool",
                    "required": False,
                    "default": True,
                },
                
                # === 其他请求 ===
                {
                    "name": "other_requests",
                    "label": "其他诉讼请求",
                    "type": "str",
                    "required": False,
                },
                
                # === 标的总额 ===
                {
                    "name": "total_amount",
                    "label": "标的总额",
                    "type": "float",
                    "required": True,
                    "description": "诉讼标的总额",
                },
                
                # === 约定管辖 ===
                {
                    "name": "has_jurisdiction_agreement",
                    "label": "是否有管辖约定",
                    "type": "bool",
                    "required": False,
                    "default": False,
                },
                {
                    "name": "jurisdiction_clause",
                    "label": "管辖条款内容",
                    "type": "str",
                    "required": False,
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
                
                # === 请求依据 ===
                {
                    "name": "contract_basis",
                    "label": "合同约定依据",
                    "type": "str",
                    "required": False,
                },
                {
                    "name": "legal_basis",
                    "label": "法律规定依据",
                    "type": "str",
                    "required": False,
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
                
                # === 证据清单 ===
                {
                    "name": "evidence_list",
                    "label": "证据清单",
                    "type": "str",
                    "required": False,
                },
            ],
            "party_fields": {
                "plaintiff_person": [
                    {"name": "姓名", "type": "str", "required": True},
                    {"name": "性别", "type": "str", "required": True, "options": ["男", "女"]},
                    {"name": "出生日期", "type": "str", "required": False},
                    {"name": "工作单位", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "经常居住地", "type": "str", "required": False},
                    {"name": "证件类型", "type": "str", "required": True},
                    {"name": "证件号码", "type": "str", "required": True},
                    {"name": "联系电话", "type": "str", "required": False},
                ],
                "plaintiff_company": [
                    {"name": "名称", "type": "str", "required": True},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "注册地", "type": "str", "required": False},
                    {"name": "法定代表人", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "联系电话", "type": "str", "required": False},
                    {"name": "统一社会信用代码", "type": "str", "required": True},
                    {"name": "类型", "type": "str", "required": False},
                    {"name": "所有制性质", "type": "str", "required": False},
                ],
                "defendant_person": [
                    {"name": "姓名", "type": "str", "required": True},
                    {"name": "性别", "type": "str", "required": False, "options": ["男", "女"]},
                    {"name": "出生日期", "type": "str", "required": False},
                    {"name": "工作单位", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "经常居住地", "type": "str", "required": False},
                    {"name": "证件类型", "type": "str", "required": True},
                    {"name": "证件号码", "type": "str", "required": True},
                ],
                "defendant_company": [
                    {"name": "名称", "type": "str", "required": True},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "注册地", "type": "str", "required": False},
                    {"name": "法定代表人", "type": "str", "required": False},
                    {"name": "职务", "type": "str", "required": False},
                    {"name": "联系电话", "type": "str", "required": False},
                    {"name": "统一社会信用代码", "type": "str", "required": True},
                    {"name": "类型", "type": "str", "required": False},
                    {"name": "所有制性质", "type": "str", "required": False},
                ],
                "third_party_person": [
                    {"name": "姓名", "type": "str", "required": True},
                    {"name": "性别", "type": "str", "required": False, "options": ["男", "女"]},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "证件号码", "type": "str", "required": False},
                ],
                "third_party_company": [
                    {"name": "名称", "type": "str", "required": True},
                    {"name": "住所地", "type": "str", "required": True},
                    {"name": "法定代表人", "type": "str", "required": False},
                    {"name": "联系电话", "type": "str", "required": False},
                    {"name": "统一社会信用代码", "type": "str", "required": True},
                    {"name": "类型", "type": "str", "required": False},
                ],
            },
        }
    
    def calculate(self, case_data: dict) -> dict:
        """
        民间借贷纠纷计算
        
        计算本金、利息、逾期利息等
        """
        items = []
        total = 0.0
        
        # 本金
        principal = case_data.get('principal_amount', 0.0)
        if principal > 0:
            items.append({
                "name": "借款本金",
                "amount": principal,
                "formula": f"实际提供借款本金",
                "legal_basis": "《民法典》第667条",
            })
            total += principal
        
        # 利息
        interest = case_data.get('interest_amount', 0.0)
        if interest > 0:
            items.append({
                "name": "借款利息",
                "amount": interest,
                "formula": f"截至{case_data.get('interest_calc_date', '')}的利息",
                "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订) 第25条",
            })
            total += interest
        
        # 如果请求支付至实际清偿日
        if case_data.get('request_interest_to_actual', False):
            items.append({
                "name": "逾期利息（至实际清偿日）",
                "amount": 0,  # 金额待定
                "formula": "按合同约定利率计算至实际清偿之日",
                "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订) 第29条",
            })
        
        return {
            "items": items,
            "total": total,
        }
    
    def build_fill_map(self, case_data: dict, calc_result: dict) -> dict[str, str]:
        """生成书签填充映射表"""
        fill_map = {}
        
        # === T0: 原告信息 ===
        plaintiff_type = case_data.get('plaintiff_type', '自然人')
        
        if plaintiff_type == '自然人':
            # T0 Row 2: 原告自然人
            fill_map['T0_01_原告姓名'] = case_data.get('plaintiff_name', '')
            fill_map['T0_02_原告住所地'] = case_data.get('plaintiff_address', '')
            fill_map['T0_03_原告经常居住地'] = case_data.get('plaintiff_residence', '')
            fill_map['T0_04_原告证件类型'] = case_data.get('plaintiff_id_type', '居民身份证')
            fill_map['T0_05_原告证件号码'] = case_data.get('plaintiff_id_number', '')
            # 多字段段落（出生日期+民族、工作单位+职务+联系电话）需要特殊处理
            fill_map['T0_06_原告工作单位'] = case_data.get('plaintiff_work_unit', '')
            fill_map['T0_07_原告职务'] = case_data.get('plaintiff_position', '')
        else:
            # T0 Row 3: 原告法人
            fill_map['T0_01_原告名称'] = case_data.get('plaintiff_name', '')
            fill_map['T0_02_原告住所地'] = case_data.get('plaintiff_address', '')
            fill_map['T0_03_原告注册地'] = case_data.get('plaintiff_registration_address', '')
            fill_map['T0_04_原告法定代表人'] = case_data.get('plaintiff_legal_rep', '')
            fill_map['T0_05_原告法定代表人职务'] = case_data.get('plaintiff_legal_rep_position', '')
            fill_map['T0_06_原告联系电话'] = case_data.get('plaintiff_phone', '')
            fill_map['T0_07_原告统一社会信用代码'] = case_data.get('plaintiff_id_number', '')
            fill_map['T0_08_原告公司类型'] = case_data.get('plaintiff_company_type', '')
            fill_map['T0_09_原告所有制性质'] = case_data.get('plaintiff_ownership', '')
        
        # === T1: 委托代理人 ===
        if case_data.get('has_agent'):
            fill_map['T1_01_代理人姓名'] = case_data.get('agent_name', '')
            fill_map['T1_02_代理人单位'] = case_data.get('agent_unit', '')
            fill_map['T1_03_代理人职务'] = case_data.get('agent_position', '')
            fill_map['T1_04_代理人联系电话'] = case_data.get('agent_phone', '')
        
        # === T1: 被告信息 ===
        defendant_type = case_data.get('defendant_type', '自然人')
        
        if defendant_type == '自然人':
            # T1 Row 1: 被告自然人
            fill_map['T1_05_被告姓名'] = case_data.get('defendant_name', '')
            fill_map['T1_06_被告住所地'] = case_data.get('defendant_address', '')
            fill_map['T1_07_被告经常居住地'] = case_data.get('defendant_residence', '')
            fill_map['T1_08_被告证件类型'] = case_data.get('defendant_id_type', '居民身份证')
            fill_map['T1_09_被告证件号码'] = case_data.get('defendant_id_number', '')
            fill_map['T1_10_被告工作单位'] = case_data.get('defendant_work_unit', '')
            fill_map['T1_11_被告职务'] = case_data.get('defendant_position', '')
        else:
            # T1 Row 2: 被告法人
            fill_map['T1_05_被告名称'] = case_data.get('defendant_name', '')
            fill_map['T1_06_被告住所地'] = case_data.get('defendant_address', '')
            fill_map['T1_07_被告注册地'] = case_data.get('defendant_registration_address', '')
            fill_map['T1_08_被告法定代表人'] = case_data.get('defendant_legal_rep', '')
            fill_map['T1_09_被告法定代表人职务'] = case_data.get('defendant_legal_rep_position', '')
            fill_map['T1_10_被告联系电话'] = case_data.get('defendant_phone', '')
            fill_map['T1_11_被告统一社会信用代码'] = case_data.get('defendant_id_number', '')
            fill_map['T1_12_被告公司类型'] = case_data.get('defendant_company_type', '')
            fill_map['T1_13_被告所有制性质'] = case_data.get('defendant_ownership', '')
        
        # === T1: 第三人信息 ===
        if case_data.get('has_third_party'):
            tp_type = case_data.get('third_party_type', '自然人')
            if tp_type == '自然人':
                # T1 Row 3: 第三人自然人
                fill_map['T1_14_第三人姓名'] = case_data.get('third_party_name', '')
                fill_map['T1_15_第三人住所地'] = case_data.get('third_party_address', '')
                fill_map['T1_16_第三人证件号码'] = case_data.get('third_party_id_number', '')
            else:
                # T1 Row 4: 第三人法人（实际在T2 Row 0）
                fill_map['T1_17_第三人名称'] = case_data.get('third_party_name', '')
                fill_map['T1_18_第三人住所地'] = case_data.get('third_party_address', '')
                fill_map['T1_19_第三人法定代表人'] = ''
                fill_map['T1_20_第三人联系电话'] = ''
                fill_map['T1_21_第三人证件号码'] = case_data.get('third_party_id_number', '')
                fill_map['T1_22_第三人公司类型'] = ''
        
        # === T2: 诉讼请求 ===
        # T2 Row 3: 本金
        fill_map['T2_01_本金截止日期'] = case_data.get('principal_calc_date', '')
        fill_map['T2_02_本金金额'] = str(int(case_data.get('principal_amount', 0)))
        
        # T2 Row 4: 利息
        fill_map['T2_03_利息截止日期'] = case_data.get('interest_calc_date', '')
        fill_map['T2_04_利息金额'] = str(int(case_data.get('interest_amount', 0)))
        
        # T2 Row 5: 提前还款/解除合同
        # T2 Row 6: 担保权利
        if case_data.get('request_guarantee_rights'):
            fill_map['T2_05_担保权利内容'] = case_data.get('guarantee_rights_content', '')
        
        # T2 Row 7: 债权费用
        if case_data.get('request_creditor_expenses'):
            fill_map['T2_06_债权费用明细'] = case_data.get('creditor_expenses_detail', '')
        
        # T2 Row 9: 其他请求
        if case_data.get('other_requests'):
            fill_map['T2_07_其他请求'] = case_data.get('other_requests', '')
        
        # T2 Row 10: 标的总额
        fill_map['T2_08_标的总额'] = str(int(case_data.get('total_amount', 0)))
        
        # T2 Row 12: 约定管辖
        if case_data.get('has_jurisdiction_agreement'):
            fill_map['T2_09_管辖条款内容'] = case_data.get('jurisdiction_clause', '')
        
        # === T3: 事实与理由 Part 1 ===
        # T3 Row 2: 合同签订情况
        fill_map['T3_01_合同名称'] = case_data.get('contract_name', '借款合同')
        fill_map['T3_02_合同编号'] = case_data.get('contract_number', '')
        fill_map['T3_03_合同签订日期'] = case_data.get('contract_sign_date', '')
        fill_map['T3_04_合同签订地点'] = case_data.get('contract_sign_place', '')
        
        # T3 Row 3: 签订主体
        fill_map['T3_05_出借人'] = case_data.get('lender_name', '')
        fill_map['T3_06_借款人'] = case_data.get('borrower_name', '')
        
        # T3 Row 4: 借款金额
        fill_map['T3_07_约定借款金额'] = str(int(case_data.get('loan_amount_agreed', 0)))
        fill_map['T3_08_实际提供金额'] = str(int(case_data.get('loan_amount_actual', 0)))
        fill_map['T3_09_提供方式其他'] = case_data.get('loan_provision_method_other', '')
        
        # T3 Row 5: 借款期限
        fill_map['T3_10_借款到期日期'] = case_data.get('loan_end_date', '')
        
        # T3 Row 6: 借款利率
        rate_type = case_data.get('interest_rate_type', '年')
        fill_map['T3_11_年利率数值'] = str(case_data.get('interest_rate', ''))
        fill_map['T3_12_利率类型'] = rate_type
        fill_map['T3_13_合同条款编号'] = case_data.get('contract_article_number', '')
        
        # T3 Row 7: 借款提供时间
        fill_map['T3_14_实际提供日期'] = case_data.get('actual_provision_date', '')
        fill_map['T3_15_实际提供金额'] = str(int(case_data.get('actual_provision_amount', 0)))
        
        # T3 Row 8: 还款方式
        fill_map['T3_16_还款方式其他'] = case_data.get('repayment_method_other', '')
        
        # T3 Row 9: 还款情况
        fill_map['T3_17_已还本金'] = str(int(case_data.get('repaid_principal', 0)))
        fill_map['T3_18_已还利息'] = str(int(case_data.get('repaid_interest', 0)))
        fill_map['T3_19_还息截止日期'] = case_data.get('interest_paid_to_date', '')
        
        # T3 Row 10: 逾期
        if case_data.get('is_overdue'):
            fill_map['T3_20_逾期起始日期'] = case_data.get('overdue_start_date', '')
            fill_map['T3_21_已逾期月数'] = str(case_data.get('overdue_months', ''))
        
        # T3 Row 11: 物的担保
        if case_data.get('has_property_guarantee'):
            fill_map['T3_22_担保合同签订时间'] = case_data.get('property_guarantee_sign_date', '')
        
        # T3 Row 12: 担保人、担保物
        fill_map['T3_23_担保人'] = case_data.get('guarantor_name', '')
        fill_map['T3_24_担保物'] = case_data.get('guarantee_property', '')
        
        # T3 Row 13: 最高额担保
        if case_data.get('has_maximum_guarantee'):
            fill_map['T3_25_担保债权确定时间'] = case_data.get('guarantee_determination_date', '')
            fill_map['T3_26_担保额度'] = str(int(case_data.get('guarantee_amount', 0)))
        
        # === T4: 事实与理由 Part 2 ===
        # T4 Row 0: 抵押质押登记
        if case_data.get('has_registration'):
            fill_map['T4_01_登记类型'] = case_data.get('registration_type', '')
        
        # T4 Row 1: 保证合同
        if case_data.get('has_surety'):
            fill_map['T4_02_保证合同签订时间'] = case_data.get('surety_sign_date', '')
            fill_map['T4_03_保证人'] = case_data.get('surety_person', '')
            fill_map['T4_04_保证主要内容'] = case_data.get('surety_main_content', '')
        
        # T4 Row 2: 其他担保
        if case_data.get('has_other_guarantee'):
            fill_map['T4_05_其他担保形式'] = case_data.get('other_guarantee_form', '')
            fill_map['T4_06_其他担保签订时间'] = case_data.get('other_guarantee_sign_date', '')
        
        # T4 Row 3: 其他说明
        if case_data.get('other_matters'):
            fill_map['T4_07_其他说明'] = case_data.get('other_matters', '')
        
        # T4 Row 4: 请求依据
        fill_map['T4_08_合同约定依据'] = case_data.get('contract_basis', '')
        fill_map['T4_09_法律规定依据'] = case_data.get('legal_basis', '')
        
        # T4 Row 5: 证据清单
        if case_data.get('evidence_list'):
            fill_map['T4_10_证据清单'] = case_data.get('evidence_list', '')
        
        return fill_map
    
    def get_checkbox_map(self, case_data: dict) -> dict[str, bool]:
        """生成勾选框映射表"""
        checkbox_map = {}
        
        # === T0: 原告类型 ===
        plaintiff_type = case_data.get('plaintiff_type', '自然人')
        checkbox_map['T0_原告类型_自然人'] = (plaintiff_type == '自然人')
        checkbox_map['T0_原告类型_法人'] = (plaintiff_type == '法人')
        
        # T0: 原告性别
        plaintiff_gender = case_data.get('plaintiff_gender', '')
        checkbox_map['T0_性别_男'] = (plaintiff_gender == '男')
        checkbox_map['T0_性别_女'] = (plaintiff_gender == '女')
        
        # T0: 原告公司类型
        plaintiff_company_type = case_data.get('plaintiff_company_type', '')
        company_types = [
            '有限责任公司', '股份有限公司', '上市公司', '其他企业法人',
            '事业单位', '社会团体', '基金会', '社会服务机构',
            '机关法人', '农村集体经济组织法人', '城镇农村的合作经济组织法人',
            '基层群众性自治组织法人', '个人独资企业', '合伙企业',
            '不具有法人资格的专业服务机构',
        ]
        for ct in company_types:
            checkbox_map[f'T0_公司类型_{ct}'] = (plaintiff_company_type == ct)
        
        # T0: 原告所有制性质
        plaintiff_ownership = case_data.get('plaintiff_ownership', '')
        checkbox_map['T0_所有制_国有'] = (plaintiff_ownership == '国有')
        checkbox_map['T0_所有制_民营'] = (plaintiff_ownership == '民营')
        checkbox_map['T0_所有制_其他'] = (plaintiff_ownership == '其他')
        
        # === T1: 委托代理人 ===
        has_agent = case_data.get('has_agent', False)
        checkbox_map['T1_代理人_有'] = has_agent
        checkbox_map['T1_代理人_无'] = not has_agent
        
        if has_agent:
            auth = case_data.get('agent_authorization', '')
            checkbox_map['T1_代理权限_一般授权'] = (auth == '一般授权')
            checkbox_map['T1_代理权限_特别授权'] = (auth == '特别授权')
        
        # T1: 被告类型
        defendant_type = case_data.get('defendant_type', '自然人')
        checkbox_map['T1_被告类型_自然人'] = (defendant_type == '自然人')
        checkbox_map['T1_被告类型_法人'] = (defendant_type == '法人')
        
        # T1: 被告性别
        defendant_gender = case_data.get('defendant_gender', '')
        checkbox_map['T1_性别_男'] = (defendant_gender == '男')
        checkbox_map['T1_性别_女'] = (defendant_gender == '女')
        
        # T1: 被告公司类型
        defendant_company_type = case_data.get('defendant_company_type', '')
        for ct in company_types:
            checkbox_map[f'T1_公司类型_{ct}'] = (defendant_company_type == ct)
        
        # T1: 被告所有制性质
        defendant_ownership = case_data.get('defendant_ownership', '')
        checkbox_map['T1_所有制_国有'] = (defendant_ownership == '国有')
        checkbox_map['T1_所有制_民营'] = (defendant_ownership == '民营')
        checkbox_map['T1_所有制_其他'] = (defendant_ownership == '其他')
        
        # T1: 第三人类型
        has_third_party = case_data.get('has_third_party', False)
        checkbox_map['T1_第三人_有'] = has_third_party
        checkbox_map['T1_第三人_无'] = not has_third_party
        
        if has_third_party:
            tp_type = case_data.get('third_party_type', '自然人')
            checkbox_map['T1_第三人类型_自然人'] = (tp_type == '自然人')
            checkbox_map['T1_第三人类型_法人'] = (tp_type == '法人')
            
            tp_gender = case_data.get('third_party_gender', '')
            checkbox_map['T1_第三人性别_男'] = (tp_gender == '男')
            checkbox_map['T1_第三人性别_女'] = (tp_gender == '女')
        
        # === T2: 诉讼请求 ===
        # 利息支付至实际清偿日
        request_interest_to_actual = case_data.get('request_interest_to_actual', True)
        checkbox_map['T2_利息至清偿_是'] = request_interest_to_actual
        checkbox_map['T2_利息至清偿_否'] = not request_interest_to_actual
        
        # 提前还款/解除合同
        request_acceleration = case_data.get('request_acceleration', False)
        request_rescission = case_data.get('request_rescission', False)
        checkbox_map['T2_提前还款_是'] = request_acceleration
        checkbox_map['T2_提前还款_否'] = not request_acceleration
        checkbox_map['T2_解除合同_是'] = request_rescission
        checkbox_map['T2_解除合同_否'] = not request_rescission
        
        # 担保权利
        request_guarantee_rights = case_data.get('request_guarantee_rights', False)
        checkbox_map['T2_担保权利_是'] = request_guarantee_rights
        checkbox_map['T2_担保权利_否'] = not request_guarantee_rights
        
        # 债权费用
        request_creditor_expenses = case_data.get('request_creditor_expenses', False)
        checkbox_map['T2_债权费用_是'] = request_creditor_expenses
        checkbox_map['T2_债权费用_否'] = not request_creditor_expenses
        
        # 诉讼费用
        request_litigation_fee = case_data.get('request_litigation_fee', True)
        checkbox_map['T2_诉讼费用_是'] = request_litigation_fee
        checkbox_map['T2_诉讼费用_否'] = not request_litigation_fee
        
        # 约定管辖
        has_jurisdiction_agreement = case_data.get('has_jurisdiction_agreement', False)
        checkbox_map['T2_管辖约定_有'] = has_jurisdiction_agreement
        checkbox_map['T2_管辖约定_无'] = not has_jurisdiction_agreement
        
        # 诉前保全
        has_preservation = case_data.get('has_preservation', False)
        checkbox_map['T2_诉前保全_是'] = has_preservation
        checkbox_map['T2_诉前保全_否'] = not has_preservation
        
        # === T3: 事实与理由 ===
        # 借款提供方式
        provision_method = case_data.get('loan_provision_method', '')
        checkbox_map['T3_提供方式_现金'] = (provision_method == '现金')
        checkbox_map['T3_提供方式_转账'] = (provision_method == '转账')
        checkbox_map['T3_提供方式_其他'] = (provision_method == '其他')
        
        # 是否到期
        is_matured = case_data.get('is_matured', False)
        checkbox_map['T3_借款到期_是'] = is_matured
        checkbox_map['T3_借款到期_否'] = not is_matured
        
        # 还款方式
        repayment_method = case_data.get('repayment_method', '')
        checkbox_map['T3_还款方式_到期一次性'] = (repayment_method == '到期一次性还本付息')
        checkbox_map['T3_还款方式_按月计息'] = (repayment_method == '按月计息、到期一次性还本')
        checkbox_map['T3_还款方式_按季计息'] = (repayment_method == '按季计息、到期一次性还本')
        checkbox_map['T3_还款方式_按年计息'] = (repayment_method == '按年计息、到期一次性还本')
        checkbox_map['T3_还款方式_其他'] = (repayment_method == '其他')
        
        # 是否逾期
        is_overdue = case_data.get('is_overdue', False)
        checkbox_map['T3_逾期_是'] = is_overdue
        checkbox_map['T3_逾期_否'] = not is_overdue
        
        # 物的担保
        has_property_guarantee = case_data.get('has_property_guarantee', False)
        checkbox_map['T3_物的担保_是'] = has_property_guarantee
        checkbox_map['T3_物的担保_否'] = not has_property_guarantee
        
        # 最高额担保
        has_maximum_guarantee = case_data.get('has_maximum_guarantee', False)
        checkbox_map['T3_最高额担保_是'] = has_maximum_guarantee
        checkbox_map['T3_最高额担保_否'] = not has_maximum_guarantee
        
        # === T4: 事实与理由 Part 2 ===
        # 抵押质押登记
        has_registration = case_data.get('has_registration', False)
        checkbox_map['T4_抵押质押登记_是'] = has_registration
        checkbox_map['T4_抵押质押登记_否'] = not has_registration
        
        if has_registration:
            reg_type = case_data.get('registration_type', '')
            checkbox_map['T4_登记类型_正式登记'] = (reg_type == '正式登记')
            checkbox_map['T4_登记类型_预告登记'] = (reg_type == '预告登记')
        
        # 保证合同
        has_surety = case_data.get('has_surety', False)
        checkbox_map['T4_保证合同_是'] = has_surety
        checkbox_map['T4_保证合同_否'] = not has_surety
        
        if has_surety:
            surety_type = case_data.get('surety_type', '')
            checkbox_map['T4_保证方式_一般保证'] = (surety_type == '一般保证')
            checkbox_map['T4_保证方式_连带责任保证'] = (surety_type == '连带责任保证')
        
        # 其他担保
        has_other_guarantee = case_data.get('has_other_guarantee', False)
        checkbox_map['T4_其他担保_是'] = has_other_guarantee
        checkbox_map['T4_其他担保_否'] = not has_other_guarantee
        
        # 调解意愿
        understand_mediation = case_data.get('understand_mediation', False)
        checkbox_map['T4_了解调解_了解'] = understand_mediation
        checkbox_map['T4_了解调解_不了解'] = not understand_mediation
        
        understand_benefits = case_data.get('understand_mediation_benefits', False)
        checkbox_map['T4_了解先行调解_了解'] = understand_benefits
        checkbox_map['T4_了解先行调解_不了解'] = understand_benefits  # 注意：这里应该是not understand_benefits
        
        # === T5: 先行调解 ===
        consider_mediation = case_data.get('consider_mediation', '')
        checkbox_map['T5_考虑先行调解_是'] = (consider_mediation == '是')
        checkbox_map['T5_考虑先行调解_否'] = (consider_mediation == '否')
        checkbox_map['T5_考虑先行调解_暂不确定'] = (consider_mediation == '暂不确定')
        
        return checkbox_map
    
    def get_template_info(self) -> dict:
        """返回模板配置信息"""
        return {
            "template_file": "民事起诉状-民间借贷纠纷.docx",
            "table_configs": [
                {
                    "index": 0,
                    "type": "plaintiff",
                    "description": "原告信息",
                    "rows": {
                        "plaintiff_person": 2,
                        "plaintiff_company": 3,
                    },
                },
                {
                    "index": 1,
                    "type": "defendant_thirdparty_agent",
                    "description": "被告、第三人、代理人信息",
                    "rows": {
                        "agent": 0,
                        "defendant_person": 1,
                        "defendant_company": 2,
                        "third_party_person": 3,
                    },
                },
                {
                    "index": 2,
                    "type": "claims",
                    "description": "诉讼请求",
                },
                {
                    "index": 3,
                    "type": "facts_part1",
                    "description": "事实与理由 Part 1",
                },
                {
                    "index": 4,
                    "type": "facts_part2",
                    "description": "事实与理由 Part 2",
                },
                {
                    "index": 5,
                    "type": "mediation",
                    "description": "调解意愿",
                },
            ],
        }
    
    def get_evidence_rules(self) -> list:
        """返回证据规则列表"""
        return [
            {
                "name": "借款合同/借据",
                "description": "证明借贷关系存在的书面证据",
                "required": True,
                "legal_basis": "《民法典》第668条",
            },
            {
                "name": "转账凭证/收据",
                "description": "证明借款实际提供的证据",
                "required": True,
                "legal_basis": "《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订) 第9条",
            },
            {
                "name": "还款凭证",
                "description": "证明已还款项的证据",
                "required": False,
                "legal_basis": "《民事诉讼法》第67条",
            },
            {
                "name": "担保合同/登记证明",
                "description": "证明担保关系的证据",
                "required": False,
                "legal_basis": "《民法典》第388条、第401条、第425条",
            },
            {
                "name": "催款通知",
                "description": "证明诉讼时效中断的证据",
                "required": False,
                "legal_basis": "《民法典》第195条",
            },
        ]
