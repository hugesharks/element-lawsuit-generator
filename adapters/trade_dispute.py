# -*- coding: utf-8 -*-
"""
买卖合同纠纷适配器

覆盖买卖合同纠纷的要素式文书生成。

模板结构（民事起诉状-买卖合同纠纷.docx）：
- 表格0：原告信息（自然人/法人）
- 表格1：委托代理人 + 被告（自然人/法人）+ 第三人（自然人/法人）
- 表格2：诉讼请求（给付价款/利息违约金/违约损失/瑕疵责任/继续履行/担保/债权费用/诉讼费等）
- 表格3：约定管辖 + 诉前保全 + 事实与理由（合同签订/标的物/价款/交货/质量/违约金等）
- 表格4：履行情况（价款交付/迟延履行/催促/质量争议/解除合同等）+ 担保（抵押质押/保证合同/保证方式等）
- 表格5：其他说明 + 证据 + 调解意愿

书签命名规范：T{表格索引}_{序号}_{字段名}
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter


class TradeDisputeAdapter(CaseAdapter):
    """
    买卖合同纠纷适配器
    
    支持：
    - 原告为卖方或买方
    - 给付价款请求
    - 迟延给付利息/违约金
    - 违约损失赔偿
    - 标的物瑕疵责任
    - 继续履行或解除合同
    - 担保权利主张
    - 债权费用
    - 物的担保（抵押、质押）
    - 保证担保（一般保证、连带责任保证）
    - 诉前保全
    """
    
    def name(self) -> str:
        return "买卖合同纠纷"
    
    def get_template_name(self) -> str:
        return "民事起诉状-买卖合同纠纷.docx"
    
    def get_schema(self) -> dict:
        return {
            "fields": [
                # === 原告信息 ===
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
                    "name": "plaintiff_ethnicity",
                    "label": "原告民族",
                    "type": "str",
                    "required": False,
                    "description": "如：汉族",
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
                    "name": "plaintiff_phone",
                    "label": "原告联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
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
                
                # === 代理人信息 ===
                {
                    "name": "has_agent",
                    "label": "是否有代理人",
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
                    "description": "代理人姓名",
                },
                {
                    "name": "agent_unit",
                    "label": "代理人单位",
                    "type": "str",
                    "required": False,
                    "description": "代理人所属单位",
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
                
                # === 被告信息 ===
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
                    "name": "defendant_ethnicity",
                    "label": "被告民族",
                    "type": "str",
                    "required": False,
                    "description": "如：汉族",
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
                    "name": "defendant_phone",
                    "label": "被告联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
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
                    "required": False,
                    "default": "居民身份证",
                    "description": "证件类型",
                },
                {
                    "name": "defendant_id_number",
                    "label": "被告证件号码",
                    "type": "str",
                    "required": False,
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
                    "description": "第三人是自然人还是法人",
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
                    "name": "third_party_birthdate",
                    "label": "第三人出生日期",
                    "type": "str",
                    "required": False,
                    "description": "格式：1990年1月1日",
                },
                {
                    "name": "third_party_ethnicity",
                    "label": "第三人民族",
                    "type": "str",
                    "required": False,
                    "description": "如：汉族",
                },
                {
                    "name": "third_party_work_unit",
                    "label": "第三人工作单位",
                    "type": "str",
                    "required": False,
                    "description": "工作单位",
                },
                {
                    "name": "third_party_position",
                    "label": "第三人职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "third_party_phone",
                    "label": "第三人联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
                },
                {
                    "name": "third_party_address",
                    "label": "第三人住所地",
                    "type": "str",
                    "required": False,
                    "description": "户籍所在地或主要办事机构所在地",
                },
                {
                    "name": "third_party_residence",
                    "label": "第三人经常居住地",
                    "type": "str",
                    "required": False,
                    "description": "经常居住地（如与住所地不同）",
                },
                {
                    "name": "third_party_id_type",
                    "label": "第三人证件类型",
                    "type": "str",
                    "required": False,
                    "default": "居民身份证",
                    "description": "证件类型",
                },
                {
                    "name": "third_party_id_number",
                    "label": "第三人证件号码",
                    "type": "str",
                    "required": False,
                    "description": "身份证号或统一社会信用代码",
                },
                # 第三人法人信息
                {
                    "name": "third_party_registration_address",
                    "label": "第三人注册地",
                    "type": "str",
                    "required": False,
                    "description": "注册地/登记地",
                },
                {
                    "name": "third_party_legal_rep",
                    "label": "第三人法定代表人",
                    "type": "str",
                    "required": False,
                    "description": "法定代表人/负责人",
                },
                {
                    "name": "third_party_legal_rep_position",
                    "label": "第三人法定代表人职务",
                    "type": "str",
                    "required": False,
                    "description": "职务",
                },
                {
                    "name": "third_party_company_type",
                    "label": "第三人公司类型",
                    "type": "str",
                    "required": False,
                    "description": "公司类型",
                },
                {
                    "name": "third_party_ownership",
                    "label": "第三人所有制性质",
                    "type": "str",
                    "required": False,
                    "options": ["国有", "民营", "其他"],
                    "description": "所有制性质",
                },
                
                # === 原告诉讼角色 ===
                {
                    "name": "plaintiff_role",
                    "label": "原告角色",
                    "type": "str",
                    "required": True,
                    "options": ["卖方", "买方"],
                    "description": "原告是卖方还是买方",
                },
                
                # === 诉讼请求 ===
                {
                    "name": "claim_price_amount",
                    "label": "给付价款金额",
                    "type": "float",
                    "required": True,
                    "description": "请求给付的价款金额（元）",
                },
                {
                    "name": "claim_interest_until_date",
                    "label": "利息截止日期",
                    "type": "str",
                    "required": False,
                    "description": "利息计算截止日期",
                },
                {
                    "name": "claim_interest_amount",
                    "label": "利息金额",
                    "type": "float",
                    "required": False,
                    "description": "迟延给付价款的利息金额",
                },
                {
                    "name": "claim_penalty_amount",
                    "label": "违约金金额",
                    "type": "float",
                    "required": False,
                    "description": "违约金金额",
                },
                {
                    "name": "claim_interest_penalty_from_date",
                    "label": "逾期利息违约金起始日期",
                    "type": "str",
                    "required": False,
                    "description": "之后逾期利息、违约金起始日期",
                },
                {
                    "name": "claim_interest_penalty_base",
                    "label": "逾期利息违约金计算基数",
                    "type": "float",
                    "required": False,
                    "description": "逾期利息、违约金计算基数",
                },
                {
                    "name": "claim_interest_penalty_rate",
                    "label": "逾期利息违约金标准",
                    "type": "str",
                    "required": False,
                    "description": "如：年利率4.35%",
                },
                {
                    "name": "claim_interest_penalty_to_actual",
                    "label": "是否请求支付至实际清偿之日",
                    "type": "bool",
                    "required": False,
                    "description": "是否请求支付至实际清偿之日止",
                },
                {
                    "name": "claim_loss_type",
                    "label": "违约类型",
                    "type": "str",
                    "required": False,
                    "options": ["迟延履行", "不履行", "其他"],
                    "description": "违约类型",
                },
                {
                    "name": "claim_loss_amount",
                    "label": "赔偿金金额",
                    "type": "float",
                    "required": False,
                    "description": "赔偿因卖方违约所受的损失金额",
                },
                {
                    "name": "claim_loss_detail",
                    "label": "违约具体情形",
                    "type": "str",
                    "required": False,
                    "description": "违约具体情形描述",
                },
                {
                    "name": "claim_loss_basis",
                    "label": "损失计算依据",
                    "type": "str",
                    "required": False,
                    "description": "损失计算依据",
                },
                {
                    "name": "claim_defect_responsibility",
                    "label": "是否对标的物瑕疵承担责任",
                    "type": "str",
                    "required": False,
                    "options": ["是", "否"],
                    "description": "是否对标的物的瑕疵承担责任",
                },
                {
                    "name": "claim_defect_way",
                    "label": "瑕疵责任方式",
                    "type": "str",
                    "required": False,
                    "description": "修理/重作/更换/退货/减少价款等",
                },
                {
                    "name": "claim_defect_other",
                    "label": "瑕疵责任其他方式",
                    "type": "str",
                    "required": False,
                    "description": "其他瑕疵责任方式",
                },
                {
                    "name": "claim_continue_perform",
                    "label": "要求继续履行或解除",
                    "type": "str",
                    "required": False,
                    "options": ["继续履行", "解除合同", "无"],
                    "description": "继续履行或解除合同",
                },
                {
                    "name": "claim_continue_fulfill_type",
                    "label": "继续履行义务类型",
                    "type": "str",
                    "required": False,
                    "description": "付款/供货等",
                },
                {
                    "name": "claim_continue_fulfill_deadline",
                    "label": "继续履行期限",
                    "type": "str",
                    "required": False,
                    "description": "日内履行完毕",
                },
                {
                    "name": "claim_contract_dissolved_date",
                    "label": "合同解除日期",
                    "type": "str",
                    "required": False,
                    "description": "确认买卖合同已解除的日期",
                },
                {
                    "name": "claim_guarantee_rights",
                    "label": "是否主张担保权利",
                    "type": "bool",
                    "required": False,
                    "description": "是否主张担保权利",
                },
                {
                    "name": "claim_guarantee_rights_content",
                    "label": "担保权利内容",
                    "type": "str",
                    "required": False,
                    "description": "担保权利具体内容",
                },
                {
                    "name": "claim_creditor_expenses",
                    "label": "是否主张债权费用",
                    "type": "bool",
                    "required": False,
                    "description": "是否主张实现债权的费用",
                },
                {
                    "name": "claim_creditor_expenses_detail",
                    "label": "债权费用明细",
                    "type": "str",
                    "required": False,
                    "description": "债权费用明细",
                },
                {
                    "name": "claim_litigation_fee",
                    "label": "是否主张诉讼费用",
                    "type": "bool",
                    "required": False,
                    "description": "是否主张诉讼费用由被告承担",
                },
                {
                    "name": "claim_other_requests",
                    "label": "其他请求",
                    "type": "str",
                    "required": False,
                    "description": "其他诉讼请求",
                },
                {
                    "name": "total_amount",
                    "label": "标的总额",
                    "type": "float",
                    "required": True,
                    "description": "标的总额（元）",
                },
                
                # === 约定管辖和诉前保全 ===
                {
                    "name": "has_jurisdiction_agreement",
                    "label": "是否有管辖约定",
                    "type": "bool",
                    "required": False,
                    "description": "是否有仲裁、法院管辖约定",
                },
                {
                    "name": "jurisdiction_clause",
                    "label": "管辖条款内容",
                    "type": "str",
                    "required": False,
                    "description": "合同条款及内容",
                },
                {
                    "name": "has_preservation",
                    "label": "是否已诉前保全",
                    "type": "bool",
                    "required": False,
                    "description": "是否已经诉前保全",
                },
                {
                    "name": "preservation_court",
                    "label": "保全法院",
                    "type": "str",
                    "required": False,
                    "description": "保全法院名称",
                },
                {
                    "name": "preservation_time",
                    "label": "保全时间",
                    "type": "str",
                    "required": False,
                    "description": "保全时间",
                },
                {
                    "name": "preservation_case_number",
                    "label": "保全案号",
                    "type": "str",
                    "required": False,
                    "description": "保全案号",
                },
                
                # === 事实与理由 ===
                {
                    "name": "contract_name",
                    "label": "合同名称",
                    "type": "str",
                    "required": False,
                    "description": "合同名称",
                },
                {
                    "name": "contract_number",
                    "label": "合同编号",
                    "type": "str",
                    "required": False,
                    "description": "合同编号",
                },
                {
                    "name": "contract_sign_date",
                    "label": "合同签订日期",
                    "type": "str",
                    "required": False,
                    "description": "合同签订日期",
                },
                {
                    "name": "contract_sign_place",
                    "label": "合同签订地点",
                    "type": "str",
                    "required": False,
                    "description": "合同签订地点",
                },
                {
                    "name": "seller_name",
                    "label": "出卖人（卖方）",
                    "type": "str",
                    "required": True,
                    "description": "出卖人名称",
                },
                {
                    "name": "buyer_name",
                    "label": "买受人（买方）",
                    "type": "str",
                    "required": True,
                    "description": "买受人名称",
                },
                {
                    "name": "goods_description",
                    "label": "标的物情况",
                    "type": "str",
                    "required": True,
                    "description": "标的物名称、规格、质量、数量等",
                },
                {
                    "name": "unit_price",
                    "label": "单价",
                    "type": "float",
                    "required": False,
                    "description": "单价（元）",
                },
                {
                    "name": "total_price",
                    "label": "总价",
                    "type": "float",
                    "required": False,
                    "description": "总价（元）",
                },
                {
                    "name": "payment_method",
                    "label": "支付方式",
                    "type": "str",
                    "required": False,
                    "options": ["现金", "转账", "票据", "其他"],
                    "description": "支付方式",
                },
                {
                    "name": "payment_bill_type",
                    "label": "票据类型",
                    "type": "str",
                    "required": False,
                    "description": "票据类型（如支票、银行汇票等）",
                },
                {
                    "name": "payment_schedule",
                    "label": "支付时间安排",
                    "type": "str",
                    "required": False,
                    "options": ["一次性", "分期"],
                    "description": "一次性或分期支付",
                },
                {
                    "name": "payment_installment_detail",
                    "label": "分期方式",
                    "type": "str",
                    "required": False,
                    "description": "分期方式说明",
                },
                {
                    "name": "delivery_terms",
                    "label": "交货条款",
                    "type": "str",
                    "required": False,
                    "description": "交货时间、地点、方式、风险承担、安装、调试、验收等",
                },
                {
                    "name": "quality_standard",
                    "label": "质量标准及检验方式",
                    "type": "str",
                    "required": False,
                    "description": "合同约定的质量标准及检验方式、质量异议期限",
                },
                {
                    "name": "contract_penalty_article",
                    "label": "违约金合同条款",
                    "type": "str",
                    "required": False,
                    "description": "违约金合同条款编号",
                },
                {
                    "name": "contract_penalty_amount",
                    "label": "违约金金额",
                    "type": "float",
                    "required": False,
                    "description": "违约金金额",
                },
                {
                    "name": "contract_deposit_article",
                    "label": "定金合同条款",
                    "type": "str",
                    "required": False,
                    "description": "定金合同条款编号",
                },
                {
                    "name": "contract_deposit_amount",
                    "label": "定金金额",
                    "type": "float",
                    "required": False,
                    "description": "定金金额",
                },
                {
                    "name": "contract_delay_penalty_rate",
                    "label": "迟延履行违约金比例",
                    "type": "str",
                    "required": False,
                    "description": "迟延履行违约金%/日",
                },
                {
                    "name": "contract_delay_penalty_article",
                    "label": "迟延履行违约金条款",
                    "type": "str",
                    "required": False,
                    "description": "迟延履行违约金合同条款编号",
                },
                
                # === 履行情况 ===
                {
                    "name": "payment_delivery_status",
                    "label": "价款支付及标的物交付情况",
                    "type": "str",
                    "required": False,
                    "description": "价款支付及标的物交付情况说明",
                },
                {
                    "name": "has_delay",
                    "label": "是否存在迟延履行",
                    "type": "bool",
                    "required": False,
                    "description": "是否存在迟延履行",
                },
                {
                    "name": "delay_time",
                    "label": "迟延时间",
                    "type": "str",
                    "required": False,
                    "description": "迟延时间",
                },
                {
                    "name": "delay_type",
                    "label": "迟延类型",
                    "type": "str",
                    "required": False,
                    "options": ["逾期付款", "逾期交货"],
                    "description": "逾期付款或逾期交货",
                },
                {
                    "name": "has_reminder",
                    "label": "是否催促过履行",
                    "type": "bool",
                    "required": False,
                    "description": "是否催促过履行",
                },
                {
                    "name": "reminder_date",
                    "label": "催促日期",
                    "type": "str",
                    "required": False,
                    "description": "催促日期",
                },
                {
                    "name": "reminder_method",
                    "label": "催促方式",
                    "type": "str",
                    "required": False,
                    "description": "催促方式",
                },
                {
                    "name": "has_quality_dispute",
                    "label": "买卖合同标的有无质量争议",
                    "type": "bool",
                    "required": False,
                    "description": "买卖合同标的有无质量争议",
                },
                {
                    "name": "quality_dispute_detail",
                    "label": "质量争议具体情况",
                    "type": "str",
                    "required": False,
                    "description": "质量争议具体情况",
                },
                {
                    "name": "has_non_conformity",
                    "label": "标的物是否符合约定",
                    "type": "bool",
                    "required": False,
                    "description": "标的物质量规格或履行方式是否存在不符合约定的情况",
                },
                {
                    "name": "non_conformity_detail",
                    "label": "不符合约定具体情况",
                    "type": "str",
                    "required": False,
                    "description": "不符合约定具体情况",
                },
                {
                    "name": "has_quality_negotiation",
                    "label": "是否就质量问题协商",
                    "type": "bool",
                    "required": False,
                    "description": "是否曾就标的物质量问题进行协商",
                },
                {
                    "name": "quality_negotiation_detail",
                    "label": "质量协商具体情况",
                    "type": "str",
                    "required": False,
                    "description": "质量协商具体情况",
                },
                {
                    "name": "has_dissolution_notice",
                    "label": "是否通知解除合同",
                    "type": "bool",
                    "required": False,
                    "description": "是否通知解除合同",
                },
                {
                    "name": "dissolution_notice_detail",
                    "label": "解除合同具体情况",
                    "type": "str",
                    "required": False,
                    "description": "解除合同具体情况",
                },
                {
                    "name": "payable_interest",
                    "label": "应付利息",
                    "type": "float",
                    "required": False,
                    "description": "被告应当支付的利息（元）",
                },
                {
                    "name": "payable_penalty",
                    "label": "应付违约金",
                    "type": "float",
                    "required": False,
                    "description": "被告应当支付的违约金（元）",
                },
                {
                    "name": "payable_compensation",
                    "label": "应付赔偿金",
                    "type": "float",
                    "required": False,
                    "description": "被告应当支付的赔偿金（元）",
                },
                {
                    "name": "total_payable",
                    "label": "共计",
                    "type": "float",
                    "required": False,
                    "description": "利息、违约金、赔偿金共计（元）",
                },
                {
                    "name": "payable_calculation_method",
                    "label": "计算方式",
                    "type": "str",
                    "required": False,
                    "description": "利息、违约金、赔偿金计算方式",
                },
                
                # === 担保信息 ===
                {
                    "name": "has_property_guarantee",
                    "label": "是否签订物的担保合同",
                    "type": "bool",
                    "required": False,
                    "description": "是否签订物的担保（抵押、质押）合同",
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
                    "description": "担保人姓名或名称",
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
                    "description": "是否最高额担保（抵押、质押）",
                },
                {
                    "name": "guarantee_determination_date",
                    "label": "担保债权确定时间",
                    "type": "str",
                    "required": False,
                    "description": "担保债权的确定时间",
                },
                {
                    "name": "guarantee_amount",
                    "label": "担保额度",
                    "type": "float",
                    "required": False,
                    "description": "担保额度（元）",
                },
                {
                    "name": "has_registration",
                    "label": "是否办理抵押质押登记",
                    "type": "bool",
                    "required": False,
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
                {
                    "name": "has_surety",
                    "label": "是否签订保证合同",
                    "type": "bool",
                    "required": False,
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
                    "description": "保证人姓名或名称",
                },
                {
                    "name": "surety_main_content",
                    "label": "保证主要内容",
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
                {
                    "name": "has_other_guarantee",
                    "label": "其他担保方式",
                    "type": "bool",
                    "required": False,
                    "description": "是否有其他担保方式",
                },
                {
                    "name": "other_guarantee_form",
                    "label": "其他担保形式",
                    "type": "str",
                    "required": False,
                    "description": "其他担保形式",
                },
                {
                    "name": "other_guarantee_sign_date",
                    "label": "其他担保签订时间",
                    "type": "str",
                    "required": False,
                    "description": "其他担保合同签订时间",
                },
                
                # === 请求依据 ===
                {
                    "name": "contract_basis",
                    "label": "合同约定依据",
                    "type": "str",
                    "required": False,
                    "description": "合同约定",
                },
                {
                    "name": "legal_basis",
                    "label": "法律规定依据",
                    "type": "str",
                    "required": False,
                    "description": "法律规定",
                },
                
                # === 其他 ===
                {
                    "name": "other_matters",
                    "label": "其他需要说明的内容",
                    "type": "str",
                    "required": False,
                    "description": "其他需要说明的内容",
                },
                {
                    "name": "evidence_list",
                    "label": "证据清单",
                    "type": "str",
                    "required": False,
                    "description": "证据清单",
                },
                
                # === 调解意愿 ===
                {
                    "name": "understand_mediation",
                    "label": "是否了解调解",
                    "type": "bool",
                    "required": True,
                    "default": True,
                    "description": "是否了解调解作为非诉讼纠纷解决方式",
                },
                {
                    "name": "understand_mediation_benefits",
                    "label": "是否了解先行调解好处",
                    "type": "bool",
                    "required": True,
                    "default": True,
                    "description": "是否了解先行调解解决纠纷的好处",
                },
                {
                    "name": "consider_mediation",
                    "label": "是否考虑先行调解",
                    "type": "str",
                    "required": True,
                    "options": ["是", "否", "暂不确定"],
                    "description": "是否考虑先行调解",
                },
            ]
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
                fill_map['T1_17_第三人工作单位'] = case_data.get('third_party_work_unit', '')
                fill_map['T1_18_第三人职务'] = case_data.get('third_party_position', '')
            else:
                # T1 Row 4: 第三人法人
                fill_map['T1_14_第三人名称'] = case_data.get('third_party_name', '')
                fill_map['T1_15_第三人住所地'] = case_data.get('third_party_address', '')
                fill_map['T1_16_第三人注册地'] = case_data.get('third_party_registration_address', '')
                fill_map['T1_17_第三人法定代表人'] = case_data.get('third_party_legal_rep', '')
                fill_map['T1_18_第三人联系电话'] = case_data.get('third_party_phone', '')
                fill_map['T1_19_第三人统一社会信用代码'] = case_data.get('third_party_id_number', '')
                fill_map['T1_20_第三人公司类型'] = case_data.get('third_party_company_type', '')
                fill_map['T1_21_第三人所有制性质'] = case_data.get('third_party_ownership', '')
        
        # === T2: 诉讼请求 ===
        # T2 Row 3: 给付价款
        fill_map['T2_01_给付价款'] = str(int(case_data.get('claim_price_amount', 0)))
        
        # T2 Row 4: 迟延给付价款的利息（违约金）
        fill_map['T2_02_利息截止日期'] = case_data.get('claim_interest_until_date', '')
        fill_map['T2_03_利息金额'] = str(int(case_data.get('claim_interest_amount', 0)))
        fill_map['T2_04_违约金金额'] = str(int(case_data.get('claim_penalty_amount', 0)))
        fill_map['T2_05_逾期利息违约金起始日期'] = case_data.get('claim_interest_penalty_from_date', '')
        fill_map['T2_06_逾期利息违约金基数'] = str(int(case_data.get('claim_interest_penalty_base', 0)))
        fill_map['T2_07_逾期利息违约金标准'] = case_data.get('claim_interest_penalty_rate', '')
        
        # T2 Row 5: 违约损失
        fill_map['T2_08_赔偿金金额'] = str(int(case_data.get('claim_loss_amount', 0)))
        fill_map['T2_09_违约具体情形'] = case_data.get('claim_loss_detail', '')
        fill_map['T2_10_损失计算依据'] = case_data.get('claim_loss_basis', '')
        
        # T2 Row 6: 瑕疵责任
        fill_map['T2_11_瑕疵责任其他方式'] = case_data.get('claim_defect_other', '')
        
        # T2 Row 7: 继续履行或解除合同
        fill_map['T2_12_继续履行期限'] = case_data.get('claim_continue_fulfill_deadline', '')
        fill_map['T2_13_继续履行义务类型'] = case_data.get('claim_continue_fulfill_type', '')
        fill_map['T2_14_合同解除日期'] = case_data.get('claim_contract_dissolved_date', '')
        
        # T2 Row 8: 担保权利
        if case_data.get('claim_guarantee_rights'):
            fill_map['T2_15_担保权利内容'] = case_data.get('claim_guarantee_rights_content', '')
        
        # T2 Row 9: 债权费用
        if case_data.get('claim_creditor_expenses'):
            fill_map['T2_16_债权费用明细'] = case_data.get('claim_creditor_expenses_detail', '')
        
        # T2 Row 11: 其他请求
        if case_data.get('claim_other_requests'):
            fill_map['T2_17_其他请求'] = case_data.get('claim_other_requests', '')
        
        # T2 Row 12: 标的总额
        fill_map['T2_18_标的总额'] = str(int(case_data.get('total_amount', 0)))
        
        # === T3: 约定管辖和诉前保全 ===
        # T3 Row 1: 约定管辖
        if case_data.get('has_jurisdiction_agreement'):
            fill_map['T3_01_管辖条款内容'] = case_data.get('jurisdiction_clause', '')
        
        # T3 Row 2: 诉前保全
        if case_data.get('has_preservation'):
            fill_map['T3_02_保全法院'] = case_data.get('preservation_court', '')
            fill_map['T3_03_保全时间'] = case_data.get('preservation_time', '')
            fill_map['T3_04_保全案号'] = case_data.get('preservation_case_number', '')
        
        # T3 Row 5: 合同签订情况
        fill_map['T3_05_合同名称'] = case_data.get('contract_name', '')
        fill_map['T3_06_合同编号'] = case_data.get('contract_number', '')
        fill_map['T3_07_合同签订日期'] = case_data.get('contract_sign_date', '')
        fill_map['T3_08_合同签订地点'] = case_data.get('contract_sign_place', '')
        
        # T3 Row 6: 合同主体
        fill_map['T3_09_出卖人'] = case_data.get('seller_name', '')
        fill_map['T3_10_买受人'] = case_data.get('buyer_name', '')
        
        # T3 Row 7: 标的物情况
        fill_map['T3_11_标的物情况'] = case_data.get('goods_description', '')
        
        # T3 Row 8: 价格及支付方式
        fill_map['T3_12_单价'] = str(int(case_data.get('unit_price', 0)))
        fill_map['T3_13_总价'] = str(int(case_data.get('total_price', 0)))
        fill_map['T3_14_票据类型'] = case_data.get('payment_bill_type', '')
        fill_map['T3_15_分期方式'] = case_data.get('payment_installment_detail', '')
        
        # T3 Row 9: 交货条款
        fill_map['T3_16_交货条款'] = case_data.get('delivery_terms', '')
        
        # T3 Row 10: 质量标准
        fill_map['T3_17_质量标准'] = case_data.get('quality_standard', '')
        
        # T3 Row 11: 违约金/定金
        fill_map['T3_18_违约金条款'] = case_data.get('contract_penalty_article', '')
        fill_map['T3_19_违约金金额'] = str(int(case_data.get('contract_penalty_amount', 0)))
        fill_map['T3_20_定金条款'] = case_data.get('contract_deposit_article', '')
        fill_map['T3_21_定金金额'] = str(int(case_data.get('contract_deposit_amount', 0)))
        fill_map['T3_22_迟延履行违约金比例'] = case_data.get('contract_delay_penalty_rate', '')
        fill_map['T3_23_迟延履行违约金条款'] = case_data.get('contract_delay_penalty_article', '')
        
        # === T4: 履行情况 ===
        # T4 Row 0: 价款支付及标的物交付情况
        fill_map['T4_01_价款支付及交付情况'] = case_data.get('payment_delivery_status', '')
        
        # T4 Row 2: 催促履行
        if case_data.get('has_reminder'):
            fill_map['T4_02_催促日期'] = case_data.get('reminder_date', '')
            fill_map['T4_03_催促方式'] = case_data.get('reminder_method', '')
        
        # T4 Row 3: 质量争议
        if case_data.get('has_quality_dispute'):
            fill_map['T4_04_质量争议具体情况'] = case_data.get('quality_dispute_detail', '')
        
        # T4 Row 4: 不符合约定
        if case_data.get('has_non_conformity'):
            fill_map['T4_05_不符合约定具体情况'] = case_data.get('non_conformity_detail', '')
        
        # T4 Row 5: 质量协商
        if case_data.get('has_quality_negotiation'):
            fill_map['T4_06_质量协商具体情况'] = case_data.get('quality_negotiation_detail', '')
        
        # T4 Row 6: 解除合同
        if case_data.get('has_dissolution_notice'):
            fill_map['T4_07_解除合同具体情况'] = case_data.get('dissolution_notice_detail', '')
        
        # T4 Row 7: 利息违约金赔偿金
        fill_map['T4_08_应付利息'] = str(int(case_data.get('payable_interest', 0)))
        fill_map['T4_09_应付违约金'] = str(int(case_data.get('payable_penalty', 0)))
        fill_map['T4_10_应付赔偿金'] = str(int(case_data.get('payable_compensation', 0)))
        fill_map['T4_11_共计金额'] = str(int(case_data.get('total_payable', 0)))
        fill_map['T4_12_计算方式'] = case_data.get('payable_calculation_method', '')
        
        # T4 Row 8: 物的担保
        if case_data.get('has_property_guarantee'):
            fill_map['T4_13_担保合同签订时间'] = case_data.get('property_guarantee_sign_date', '')
        
        # T4 Row 9: 担保人、担保物
        fill_map['T4_14_担保人'] = case_data.get('guarantor_name', '')
        fill_map['T4_15_担保物'] = case_data.get('guarantee_property', '')
        
        # T4 Row 10: 最高额担保
        if case_data.get('has_maximum_guarantee'):
            fill_map['T4_16_担保债权确定时间'] = case_data.get('guarantee_determination_date', '')
            fill_map['T4_17_担保额度'] = str(int(case_data.get('guarantee_amount', 0)))
        
        # T4 Row 12: 保证合同
        if case_data.get('has_surety'):
            fill_map['T4_18_保证合同签订时间'] = case_data.get('surety_sign_date', '')
            fill_map['T4_19_保证人'] = case_data.get('surety_person', '')
            fill_map['T4_20_保证主要内容'] = case_data.get('surety_main_content', '')
        
        # T4 Row 14: 其他担保方式
        if case_data.get('has_other_guarantee'):
            fill_map['T4_21_其他担保形式'] = case_data.get('other_guarantee_form', '')
            fill_map['T4_22_其他担保签订时间'] = case_data.get('other_guarantee_sign_date', '')
        
        # T4 Row 15: 请求承担责任的依据
        fill_map['T4_23_合同约定依据'] = case_data.get('contract_basis', '')
        fill_map['T4_24_法律规定依据'] = case_data.get('legal_basis', '')
        
        # === T5: 其他 ===
        # T5 Row 0: 其他需要说明的内容
        if case_data.get('other_matters'):
            fill_map['T5_01_其他需要说明的内容'] = case_data.get('other_matters', '')
        
        # T5 Row 1: 证据清单
        if case_data.get('evidence_list'):
            fill_map['T5_02_证据清单'] = case_data.get('evidence_list', '')
        
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
        
        # === T1: 代理人 ===
        has_agent = case_data.get('has_agent', False)
        checkbox_map['T1_代理人_有'] = has_agent
        checkbox_map['T1_代理人_无'] = not has_agent
        
        # T1: 代理权限
        if has_agent:
            auth = case_data.get('agent_authorization', '')
            checkbox_map['T1_代理权限_一般授权'] = (auth == '一般授权')
            checkbox_map['T1_代理权限_特别授权'] = (auth == '特别授权')
        
        # === T1: 被告类型 ===
        defendant_type = case_data.get('defendant_type', '自然人')
        checkbox_map['T1_被告类型_自然人'] = (defendant_type == '自然人')
        checkbox_map['T1_被告类型_法人'] = (defendant_type == '法人')
        
        # T1: 被告性别
        defendant_gender = case_data.get('defendant_gender', '')
        checkbox_map['T1_被告性别_男'] = (defendant_gender == '男')
        checkbox_map['T1_被告性别_女'] = (defendant_gender == '女')
        
        # T1: 被告公司类型
        defendant_company_type = case_data.get('defendant_company_type', '')
        company_types = [
            '有限责任公司', '股份有限公司', '上市公司', '其他企业法人',
            '事业单位', '社会团体', '基金会', '社会服务机构',
            '机关法人', '农村集体经济组织法人', '城镇农村的合作经济组织法人',
            '基层群众性自治组织法人', '个人独资企业', '合伙企业',
            '不具有法人资格的专业服务机构',
        ]
        for ct in company_types:
            checkbox_map[f'T1_被告公司类型_{ct}'] = (defendant_company_type == ct)
        
        # T1: 被告所有制性质
        defendant_ownership = case_data.get('defendant_ownership', '')
        checkbox_map['T1_被告所有制_国有'] = (defendant_ownership == '国有')
        checkbox_map['T1_被告所有制_民营'] = (defendant_ownership == '民营')
        checkbox_map['T1_被告所有制_其他'] = (defendant_ownership == '其他')
        
        # === T1: 第三人 ===
        has_third_party = case_data.get('has_third_party', False)
        checkbox_map['T1_第三人_有'] = has_third_party
        checkbox_map['T1_第三人_无'] = not has_third_party
        
        if has_third_party:
            tp_type = case_data.get('third_party_type', '自然人')
            checkbox_map['T1_第三人类型_自然人'] = (tp_type == '自然人')
            checkbox_map['T1_第三人类型_法人'] = (tp_type == '法人')
            
            # 第三人性别
            tp_gender = case_data.get('third_party_gender', '')
            checkbox_map['T1_第三人性别_男'] = (tp_gender == '男')
            checkbox_map['T1_第三人性别_女'] = (tp_gender == '女')
            
            # 第三人公司类型
            tp_company_type = case_data.get('third_party_company_type', '')
            for ct in company_types:
                checkbox_map[f'T1_第三人公司类型_{ct}'] = (tp_company_type == ct)
            
            # 第三人所有制性质
            tp_ownership = case_data.get('third_party_ownership', '')
            checkbox_map['T1_第三人所有制_国有'] = (tp_ownership == '国有')
            checkbox_map['T1_第三人所有制_民营'] = (tp_ownership == '民营')
            checkbox_map['T1_第三人所有制_其他'] = (tp_ownership == '其他')
        
        # === T2: 诉讼请求 ===
        plaintiff_role = case_data.get('plaintiff_role', '卖方')
        checkbox_map['T2_原告角色_卖方'] = (plaintiff_role == '卖方')
        checkbox_map['T2_原告角色_买方'] = (plaintiff_role == '买方')
        
        # 是否请求支付至实际清偿之日
        claim_interest_penalty_to_actual = case_data.get('claim_interest_penalty_to_actual', False)
        checkbox_map['T2_支付至实际清偿_是'] = claim_interest_penalty_to_actual
        checkbox_map['T2_支付至实际清偿_否'] = not claim_interest_penalty_to_actual
        
        # 违约类型
        claim_loss_type = case_data.get('claim_loss_type', '')
        checkbox_map['T2_违约类型_迟延履行'] = (claim_loss_type == '迟延履行')
        checkbox_map['T2_违约类型_不履行'] = (claim_loss_type == '不履行')
        checkbox_map['T2_违约类型_其他'] = (claim_loss_type == '其他')
        
        # 瑕疵责任
        claim_defect_responsibility = case_data.get('claim_defect_responsibility', '')
        checkbox_map['T2_瑕疵责任_是'] = (claim_defect_responsibility == '是')
        checkbox_map['T2_瑕疵责任_否'] = (claim_defect_responsibility == '否')
        
        # 瑕疵责任方式
        claim_defect_way = case_data.get('claim_defect_way', '')
        if claim_defect_way:
            defect_ways = ['修理', '重作', '更换', '退货', '减少价款或者报酬']
            for way in defect_ways:
                checkbox_map[f'T2_瑕疵方式_{way}'] = (way in claim_defect_way)
        
        # 继续履行或解除合同
        claim_continue_perform = case_data.get('claim_continue_perform', '')
        checkbox_map['T2_继续履行_继续履行'] = (claim_continue_perform == '继续履行')
        checkbox_map['T2_解除合同_解除合同'] = (claim_continue_perform == '解除合同')
        checkbox_map['T2_确认解除_是'] = (claim_continue_perform == '解除合同' and bool(case_data.get('claim_contract_dissolved_date')))
        
        # 继续履行义务类型
        claim_continue_fulfill_type = case_data.get('claim_continue_fulfill_type', '')
        if claim_continue_fulfill_type:
            fulfill_types = ['付款', '供货']
            for ft in fulfill_types:
                checkbox_map[f'T2_履行义务_{ft}'] = (ft in claim_continue_fulfill_type)
        
        # 担保权利
        claim_guarantee_rights = case_data.get('claim_guarantee_rights', False)
        checkbox_map['T2_担保权利_是'] = claim_guarantee_rights
        checkbox_map['T2_担保权利_否'] = not claim_guarantee_rights
        
        # 债权费用
        claim_creditor_expenses = case_data.get('claim_creditor_expenses', False)
        checkbox_map['T2_债权费用_是'] = claim_creditor_expenses
        checkbox_map['T2_债权费用_否'] = not claim_creditor_expenses
        
        # 诉讼费用
        claim_litigation_fee = case_data.get('claim_litigation_fee', False)
        checkbox_map['T2_诉讼费用_是'] = claim_litigation_fee
        checkbox_map['T2_诉讼费用_否'] = not claim_litigation_fee
        
        # === T3: 约定管辖和诉前保全 ===
        has_jurisdiction_agreement = case_data.get('has_jurisdiction_agreement', False)
        checkbox_map['T3_管辖约定_有'] = has_jurisdiction_agreement
        checkbox_map['T3_管辖约定_无'] = not has_jurisdiction_agreement
        
        has_preservation = case_data.get('has_preservation', False)
        checkbox_map['T3_诉前保全_是'] = has_preservation
        checkbox_map['T3_诉前保全_否'] = not has_preservation
        
        # === T3: 事实与理由 - 支付方式 ===
        payment_method = case_data.get('payment_method', '')
        payment_methods = ['现金', '转账', '票据', '其他']
        for pm in payment_methods:
            checkbox_map[f'T3_支付方式_{pm}'] = (payment_method == pm)
        
        # 支付时间安排
        payment_schedule = case_data.get('payment_schedule', '')
        checkbox_map['T3_支付安排_一次性'] = (payment_schedule == '一次性')
        checkbox_map['T3_支付安排_分期'] = (payment_schedule == '分期')
        
        # === T4: 履行情况 ===
        has_delay = case_data.get('has_delay', False)
        checkbox_map['T4_迟延履行_是'] = has_delay
        checkbox_map['T4_迟延履行_否'] = not has_delay
        
        if has_delay:
            delay_type = case_data.get('delay_type', '')
            checkbox_map['T4_迟延类型_逾期付款'] = (delay_type == '逾期付款')
            checkbox_map['T4_迟延类型_逾期交货'] = (delay_type == '逾期交货')
        
        has_reminder = case_data.get('has_reminder', False)
        checkbox_map['T4_催促履行_是'] = has_reminder
        checkbox_map['T4_催促履行_否'] = not has_reminder
        
        has_quality_dispute = case_data.get('has_quality_dispute', False)
        checkbox_map['T4_质量争议_有'] = has_quality_dispute
        checkbox_map['T4_质量争议_无'] = not has_quality_dispute
        
        has_non_conformity = case_data.get('has_non_conformity', False)
        checkbox_map['T4_不符合约定_是'] = has_non_conformity
        checkbox_map['T4_不符合约定_否'] = not has_non_conformity
        
        has_quality_negotiation = case_data.get('has_quality_negotiation', False)
        checkbox_map['T4_质量协商_是'] = has_quality_negotiation
        checkbox_map['T4_质量协商_否'] = not has_quality_negotiation
        
        has_dissolution_notice = case_data.get('has_dissolution_notice', False)
        checkbox_map['T4_解除合同_是'] = has_dissolution_notice
        checkbox_map['T4_解除合同_否'] = not has_dissolution_notice
        
        # === T4: 担保 ===
        has_property_guarantee = case_data.get('has_property_guarantee', False)
        checkbox_map['T4_物的担保_是'] = has_property_guarantee
        checkbox_map['T4_物的担保_否'] = not has_property_guarantee
        
        has_maximum_guarantee = case_data.get('has_maximum_guarantee', False)
        checkbox_map['T4_最高额担保_是'] = has_maximum_guarantee
        checkbox_map['T4_最高额担保_否'] = not has_maximum_guarantee
        
        has_registration = case_data.get('has_registration', False)
        checkbox_map['T4_抵押质押登记_是'] = has_registration
        checkbox_map['T4_抵押质押登记_否'] = not has_registration
        
        if has_registration:
            registration_type = case_data.get('registration_type', '')
            checkbox_map['T4_登记类型_正式登记'] = (registration_type == '正式登记')
            checkbox_map['T4_登记类型_预告登记'] = (registration_type == '预告登记')
        
        has_surety = case_data.get('has_surety', False)
        checkbox_map['T4_保证合同_是'] = has_surety
        checkbox_map['T4_保证合同_否'] = not has_surety
        
        # 保证方式
        surety_type = case_data.get('surety_type', '')
        checkbox_map['T4_保证方式_一般保证'] = (surety_type == '一般保证')
        checkbox_map['T4_保证方式_连带责任保证'] = (surety_type == '连带责任保证')
        
        has_other_guarantee = case_data.get('has_other_guarantee', False)
        checkbox_map['T4_其他担保_是'] = has_other_guarantee
        checkbox_map['T4_其他担保_否'] = not has_other_guarantee
        
        # === T5: 调解意愿 ===
        understand_mediation = case_data.get('understand_mediation', True)
        checkbox_map['T5_了解调解_了解'] = understand_mediation
        checkbox_map['T5_了解调解_不了解'] = not understand_mediation
        
        understand_mediation_benefits = case_data.get('understand_mediation_benefits', True)
        checkbox_map['T5_了解先行调解_了解'] = understand_mediation_benefits
        checkbox_map['T5_了解先行调解_不了解'] = not understand_mediation_benefits
        
        consider_mediation = case_data.get('consider_mediation', '')
        checkbox_map['T5_考虑先行调解_是'] = (consider_mediation == '是')
        checkbox_map['T5_考虑先行调解_否'] = (consider_mediation == '否')
        checkbox_map['T5_考虑先行调解_暂不确定'] = (consider_mediation == '暂不确定')
        
        return checkbox_map
    
    def get_field_order(self) -> list:
        """返回字段顺序配置，用于书签插入顺序"""
        return [
            # T0: 原告信息 - 自然人（多字段段落需特殊处理）
            {
                'table': 0,
                'row': 2,
                'cell': 1,
                'paragraph_indices': [0, 2, 3, 4, 5, 6, 7],
                'fields': [
                    ('T0_01_原告姓名', 'paragraph', 0),
                    ('T0_06_原告工作单位', 'paragraph', 3),
                    ('T0_07_原告职务', 'paragraph', 3),
                    ('T0_02_原告住所地', 'paragraph', 4),
                    ('T0_03_原告经常居住地', 'paragraph', 5),
                    ('T0_04_原告证件类型', 'paragraph', 6),
                    ('T0_05_原告证件号码', 'paragraph', 7),
                ]
            },
            # T0: 原告信息 - 法人
            {
                'table': 0,
                'row': 3,
                'cell': 1,
                'paragraph_indices': [0, 1, 2, 3, 4, 5, 6],
                'fields': [
                    ('T0_01_原告名称', 'paragraph', 0),
                    ('T0_02_原告住所地', 'paragraph', 1),
                    ('T0_03_原告注册地', 'paragraph', 2),
                    ('T0_04_原告法定代表人', 'paragraph', 3),
                    ('T0_05_原告法定代表人职务', 'paragraph', 3),
                    ('T0_06_原告联系电话', 'paragraph', 3),
                    ('T0_07_原告统一社会信用代码', 'paragraph', 4),
                ]
            },
            # T1: 委托代理人
            {
                'table': 1,
                'row': 0,
                'cell': 1,
                'paragraph_indices': [1, 2, 2, 2],
                'fields': [
                    ('T1_01_代理人姓名', 'paragraph', 1),
                    ('T1_02_代理人单位', 'paragraph', 2),
                    ('T1_03_代理人职务', 'paragraph', 2),
                    ('T1_04_代理人联系电话', 'paragraph', 2),
                ]
            },
            # T1: 被告自然人
            {
                'table': 1,
                'row': 1,
                'cell': 1,
                'paragraph_indices': [0, 2, 3, 4, 5, 6, 7],
                'fields': [
                    ('T1_05_被告姓名', 'paragraph', 0),
                    ('T1_10_被告工作单位', 'paragraph', 3),
                    ('T1_11_被告职务', 'paragraph', 3),
                    ('T1_06_被告住所地', 'paragraph', 4),
                    ('T1_07_被告经常居住地', 'paragraph', 5),
                    ('T1_08_被告证件类型', 'paragraph', 6),
                    ('T1_09_被告证件号码', 'paragraph', 7),
                ]
            },
            # T1: 被告法人
            {
                'table': 1,
                'row': 2,
                'cell': 1,
                'paragraph_indices': [0, 1, 2, 3, 4, 5, 6],
                'fields': [
                    ('T1_05_被告名称', 'paragraph', 0),
                    ('T1_06_被告住所地', 'paragraph', 1),
                    ('T1_07_被告注册地', 'paragraph', 2),
                    ('T1_08_被告法定代表人', 'paragraph', 3),
                    ('T1_09_被告法定代表人职务', 'paragraph', 3),
                    ('T1_10_被告联系电话', 'paragraph', 3),
                    ('T1_11_被告统一社会信用代码', 'paragraph', 4),
                ]
            },
            # T1: 第三人自然人
            {
                'table': 1,
                'row': 3,
                'cell': 1,
                'paragraph_indices': [0, 2, 3, 4, 5, 6, 7],
                'fields': [
                    ('T1_14_第三人姓名', 'paragraph', 0),
                    ('T1_17_第三人工作单位', 'paragraph', 3),
                    ('T1_18_第三人职务', 'paragraph', 3),
                    ('T1_15_第三人住所地', 'paragraph', 4),
                    ('T1_16_第三人证件号码', 'paragraph', 7),
                ]
            },
            # T1: 第三人法人
            {
                'table': 1,
                'row': 4,
                'cell': 1,
                'paragraph_indices': [0, 1, 2, 3, 4],
                'fields': [
                    ('T1_14_第三人名称', 'paragraph', 0),
                    ('T1_15_第三人住所地', 'paragraph', 1),
                    ('T1_16_第三人注册地', 'paragraph', 2),
                    ('T1_17_第三人法定代表人', 'paragraph', 3),
                    ('T1_18_第三人联系电话', 'paragraph', 3),
                    ('T1_19_第三人统一社会信用代码', 'paragraph', 4),
                ]
            },
            # T2: 诉讼请求 - 给付价款
            {
                'table': 2,
                'row': 3,
                'cell': 1,
                'fields': [('T2_01_给付价款', 'cell')]
            },
            # T2: 诉讼请求 - 迟延利息违约金
            {
                'table': 2,
                'row': 4,
                'cell': 1,
                'paragraph_indices': [0, 0, 0, 2, 2, 3, 4],
                'fields': [
                    ('T2_02_利息截止日期', 'paragraph', 0),
                    ('T2_03_利息金额', 'paragraph', 0),
                    ('T2_04_违约金金额', 'paragraph', 0),
                    ('T2_05_逾期利息违约金起始日期', 'paragraph', 2),
                    ('T2_06_逾期利息违约金基数', 'paragraph', 2),
                    ('T2_07_逾期利息违约金标准', 'paragraph', 3),
                ]
            },
            # T2: 诉讼请求 - 违约损失
            {
                'table': 2,
                'row': 5,
                'cell': 1,
                'paragraph_indices': [0, 1, 3, 4],
                'fields': [
                    ('T2_08_赔偿金金额', 'paragraph', 0),
                    ('T2_09_违约具体情形', 'paragraph', 3),
                    ('T2_10_损失计算依据', 'paragraph', 4),
                ]
            },
            # T2: 诉讼请求 - 瑕疵责任
            {
                'table': 2,
                'row': 6,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T2_11_瑕疵责任其他方式', 'paragraph', 1),
                ]
            },
            # T2: 诉讼请求 - 继续履行或解除
            {
                'table': 2,
                'row': 7,
                'cell': 1,
                'paragraph_indices': [0, 0, 1],
                'fields': [
                    ('T2_12_继续履行期限', 'paragraph', 0),
                    ('T2_13_继续履行义务类型', 'paragraph', 0),
                    ('T2_14_合同解除日期', 'paragraph', 1),
                ]
            },
            # T2: 诉讼请求 - 担保权利
            {
                'table': 2,
                'row': 8,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T2_15_担保权利内容', 'paragraph', 1),
                ]
            },
            # T2: 诉讼请求 - 债权费用
            {
                'table': 2,
                'row': 9,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T2_16_债权费用明细', 'paragraph', 1),
                ]
            },
            # T2: 诉讼请求 - 其他请求
            {
                'table': 2,
                'row': 11,
                'cell': 0,
                'fields': [('T2_17_其他请求', 'cell')]
            },
            # T2: 诉讼请求 - 标的总额
            {
                'table': 2,
                'row': 12,
                'cell': 1,
                'fields': [('T2_18_标的总额', 'cell')]
            },
            # T3: 约定管辖
            {
                'table': 3,
                'row': 1,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T3_01_管辖条款内容', 'paragraph', 1),
                ]
            },
            # T3: 诉前保全
            {
                'table': 3,
                'row': 2,
                'cell': 1,
                'paragraph_indices': [0, 1, 2],
                'fields': [
                    ('T3_02_保全法院', 'paragraph', 0),
                    ('T3_03_保全时间', 'paragraph', 0),
                    ('T3_04_保全案号', 'paragraph', 1),
                ]
            },
            # T3: 合同签订情况
            {
                'table': 3,
                'row': 5,
                'cell': 1,
                'fields': [('T3_05_合同名称', 'cell')]
            },
            # T3: 合同主体
            {
                'table': 3,
                'row': 6,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T3_09_出卖人', 'paragraph', 0),
                    ('T3_10_买受人', 'paragraph', 1),
                ]
            },
            # T3: 标的物情况
            {
                'table': 3,
                'row': 7,
                'cell': 1,
                'fields': [('T3_11_标的物情况', 'cell')]
            },
            # T3: 价格及支付方式
            {
                'table': 3,
                'row': 8,
                'cell': 1,
                'paragraph_indices': [0, 0, 1, 1, 2],
                'fields': [
                    ('T3_12_单价', 'paragraph', 0),
                    ('T3_13_总价', 'paragraph', 0),
                    ('T3_14_票据类型', 'paragraph', 1),
                    ('T3_15_分期方式', 'paragraph', 2),
                ]
            },
            # T3: 交货条款
            {
                'table': 3,
                'row': 9,
                'cell': 1,
                'fields': [('T3_16_交货条款', 'cell')]
            },
            # T3: 质量标准
            {
                'table': 3,
                'row': 10,
                'cell': 1,
                'fields': [('T3_17_质量标准', 'cell')]
            },
            # T3: 违约金/定金
            {
                'table': 3,
                'row': 11,
                'cell': 1,
                'paragraph_indices': [0, 0, 0, 1, 1, 2, 2],
                'fields': [
                    ('T3_18_违约金条款', 'paragraph', 0),
                    ('T3_19_违约金金额', 'paragraph', 0),
                    ('T3_20_定金条款', 'paragraph', 1),
                    ('T3_21_定金金额', 'paragraph', 1),
                    ('T3_22_迟延履行违约金比例', 'paragraph', 2),
                    ('T3_23_迟延履行违约金条款', 'paragraph', 2),
                ]
            },
            # T4: 价款支付及交付情况
            {
                'table': 4,
                'row': 0,
                'cell': 1,
                'fields': [('T4_01_价款支付及交付情况', 'cell')]
            },
            # T4: 催促履行
            {
                'table': 4,
                'row': 2,
                'cell': 1,
                'paragraph_indices': [0, 0, 1],
                'fields': [
                    ('T4_02_催促日期', 'paragraph', 0),
                    ('T4_03_催促方式', 'paragraph', 0),
                ]
            },
            # T4: 质量争议
            {
                'table': 4,
                'row': 3,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_04_质量争议具体情况', 'paragraph', 1),
                ]
            },
            # T4: 不符合约定
            {
                'table': 4,
                'row': 4,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_05_不符合约定具体情况', 'paragraph', 1),
                ]
            },
            # T4: 质量协商
            {
                'table': 4,
                'row': 5,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_06_质量协商具体情况', 'paragraph', 1),
                ]
            },
            # T4: 解除合同
            {
                'table': 4,
                'row': 6,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_07_解除合同具体情况', 'paragraph', 1),
                ]
            },
            # T4: 利息违约金赔偿金
            {
                'table': 4,
                'row': 7,
                'cell': 1,
                'paragraph_indices': [0, 0, 0, 1, 1, 2, 3],
                'fields': [
                    ('T4_08_应付利息', 'paragraph', 0),
                    ('T4_09_应付违约金', 'paragraph', 0),
                    ('T4_10_应付赔偿金', 'paragraph', 0),
                    ('T4_11_共计金额', 'paragraph', 2),
                    ('T4_12_计算方式', 'paragraph', 3),
                ]
            },
            # T4: 物的担保
            {
                'table': 4,
                'row': 8,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_13_担保合同签订时间', 'paragraph', 1),
                ]
            },
            # T4: 担保人、担保物
            {
                'table': 4,
                'row': 9,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_14_担保人', 'paragraph', 0),
                    ('T4_15_担保物', 'paragraph', 1),
                ]
            },
            # T4: 最高额担保
            {
                'table': 4,
                'row': 10,
                'cell': 1,
                'paragraph_indices': [0, 1, 2],
                'fields': [
                    ('T4_16_担保债权确定时间', 'paragraph', 1),
                    ('T4_17_担保额度', 'paragraph', 2),
                ]
            },
            # T4: 保证合同
            {
                'table': 4,
                'row': 12,
                'cell': 1,
                'paragraph_indices': [0, 1, 1, 1],
                'fields': [
                    ('T4_18_保证合同签订时间', 'paragraph', 1),
                    ('T4_19_保证人', 'paragraph', 1),
                    ('T4_20_保证主要内容', 'paragraph', 1),
                ]
            },
            # T4: 其他担保方式
            {
                'table': 4,
                'row': 14,
                'cell': 1,
                'paragraph_indices': [0, 1, 1],
                'fields': [
                    ('T4_21_其他担保形式', 'paragraph', 1),
                    ('T4_22_其他担保签订时间', 'paragraph', 1),
                ]
            },
            # T4: 请求依据
            {
                'table': 4,
                'row': 15,
                'cell': 1,
                'paragraph_indices': [0, 1],
                'fields': [
                    ('T4_23_合同约定依据', 'paragraph', 0),
                    ('T4_24_法律规定依据', 'paragraph', 1),
                ]
            },
            # T5: 其他需要说明的内容
            {
                'table': 5,
                'row': 0,
                'cell': 1,
                'fields': [('T5_01_其他需要说明的内容', 'cell')]
            },
            # T5: 证据清单
            {
                'table': 5,
                'row': 1,
                'cell': 1,
                'fields': [('T5_02_证据清单', 'cell')]
            },
        ]
    
    def calculate(self, case_data: dict) -> dict:
        """
        买卖合同纠纷计算（违约金、利息等）

        Args:
            case_data: 案件数据字典

        Returns:
            计算结果字典
        """
        items = []

        # 1. 给付价款
        price_amount = float(case_data.get('claim_price_amount', 0))
        if price_amount > 0:
            items.append({
                "name": "给付价款",
                "amount": price_amount,
                "formula": f"合同约定价款",
                "legal_basis": "《民法典》第五百零九条、第五百七十九条",
            })

        # 2. 迟延利息/违约金
        interest_amount = float(case_data.get('claim_interest_amount', 0))
        penalty_amount = float(case_data.get('claim_penalty_amount', 0))

        if interest_amount > 0 or penalty_amount > 0:
            interest_penalty_basis = float(case_data.get('claim_interest_penalty_base', 0))
            interest_penalty_rate = case_data.get('claim_interest_penalty_rate', '')

            if interest_amount > 0:
                items.append({
                    "name": "迟延给付价款利息",
                    "amount": interest_amount,
                    "formula": f"以{interest_penalty_basis}元为基数计算",
                    "legal_basis": f"《民法典》第五百七十九条；{interest_penalty_rate}",
                })

            if penalty_amount > 0:
                items.append({
                    "name": "违约金",
                    "amount": penalty_amount,
                    "formula": f"合同约定",
                    "legal_basis": "《民法典》第五百八十五条",
                })

        # 3. 违约损失
        loss_amount = float(case_data.get('claim_loss_amount', 0))
        if loss_amount > 0:
            items.append({
                "name": "违约损失赔偿",
                "amount": loss_amount,
                "formula": case_data.get('claim_loss_basis', '实际损失'),
                "legal_basis": "《民法典》第五百八十四条",
            })

        # 4. 债权费用
        if case_data.get('claim_creditor_expenses'):
            expenses_detail = case_data.get('claim_creditor_expenses_detail', '')
            items.append({
                "name": "实现债权费用",
                "amount": 0,
                "formula": expenses_detail,
                "legal_basis": "《民法典》第五百八十一条",
            })

        # 计算总额
        total = sum(item.get('amount', 0) for item in items)

        return {
            "items": items,
            "total": total,
        }

    def get_evidence_rules(self) -> list:
        """返回证据规则列表"""
        return [
            {
                "name": "买卖合同",
                "description": "书面买卖合同，证明合同关系成立及约定内容",
                "required": True,
                "legal_basis": "《民法典》第四百六十九条",
            },
            {
                "name": "送货单/收货单",
                "description": "证明卖方已交付货物",
                "required": True,
                "legal_basis": "《民法典》第五百九十六条",
            },
            {
                "name": "付款凭证",
                "description": "银行转账记录、收据等，证明付款情况",
                "required": True,
                "legal_basis": "《民法典》第五百七十九条",
            },
            {
                "name": "发票",
                "description": "增值税发票或普通发票",
                "required": False,
                "legal_basis": "《发票管理办法》",
            },
            {
                "name": "催款函/律师函",
                "description": "证明原告已催告被告履行",
                "required": False,
                "legal_basis": "《民法典》第五百零九条",
            },
            {
                "name": "往来函件/聊天记录",
                "description": "证明双方协商过程",
                "required": False,
                "legal_basis": "《民法典》第一百三十七条",
            },
            {
                "name": "质量检验报告",
                "description": "标的物质量鉴定（如涉及质量争议）",
                "required": False,
                "legal_basis": "《民法典》第六百二十条、第六百二十一条",
            },
        ]

    def analyze_case_data(self, case_data: dict) -> dict:
        """
        分析案件数据，转换为fill_map格式

        这个方法用于将外部输入的案件数据转换为适配器内部使用的fill_map格式。
        子类可以重写此方法来实现自定义的数据转换逻辑。

        Args:
            case_data: 原始案件数据字典

        Returns:
            处理后的案件数据字典
        """
        result = case_data.copy()

        # 确保数值字段被正确格式化为整数
        numeric_fields = [
            'claim_price_amount', 'claim_interest_amount', 'claim_penalty_amount',
            'claim_interest_penalty_base', 'claim_loss_amount', 'unit_price',
            'total_price', 'contract_penalty_amount', 'contract_deposit_amount',
            'payable_interest', 'payable_penalty', 'payable_compensation',
            'total_payable', 'guarantee_amount', 'total_amount'
        ]

        for field in numeric_fields:
            if field in result and result[field]:
                try:
                    result[field] = int(float(result[field]))
                except (ValueError, TypeError):
                    pass

        return result
