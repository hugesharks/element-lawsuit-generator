# -*- coding: utf-8 -*-
"""
建设工程施工合同纠纷适配器

覆盖建设工程施工合同纠纷的要素式文书生成。

模板结构（民事起诉状-建设工程施工合同纠纷.docx）：
- 表格0：原告信息（自然人/法人）
- 表格1：委托代理人 + 被告（自然人/法人）+ 第三人（自然人/法人）
- 表格2：诉讼请求（3列结构！左侧标题 + 中间标签 + 右侧内容）
- 表格3：诉讼请求7-15项 + 约定管辖/诉前保全/鉴定申请 + 事实与理由
- 表格4：合同详情 + 工程款 + 质量/交付/停窝工 + 优先受偿权 + 请求依据
- 表格5：证据清单 + 调解意愿

书签命名规范：T{表格索引}_{序号}_{字段名}
勾选框命名规范：T{表格索引}_{内容描述}
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter


class ConstructionContractAdapter(CaseAdapter):
    """
    建设工程施工合同纠纷适配器
    
    支持：
    - 原告为承包人/施工人或发包人
    - 支付工程款请求
    - 迟延支付工程款利息/违约金
    - 建设工程价款优先受偿权
    - 突破合同相对性（挂靠、转包等）
    - 赔偿损失（停窝工损失等）
    - 退还超付工程款
    - 建设工程修复责任
    - 合同无效确认
    - 继续履行或解除合同
    - 债权实现费用
    - 诉讼费用
    """
    
    def name(self) -> str:
        return "建设工程施工合同纠纷"
    
    def get_template_name(self) -> str:
        return "民事起诉状-建设工程施工合同纠纷.docx"
    
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
                    "description": "职务",
                },
                {
                    "name": "agent_phone",
                    "label": "代理人联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
                },
                {
                    "name": "agent_authorization",
                    "label": "代理权限",
                    "type": "str",
                    "required": False,
                    "default": "一般授权",
                    "description": "一般授权/特别授权",
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
                    "description": "自然人填写姓名，法人填写名称",
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
                    "name": "third_party_address",
                    "label": "第三人住所地",
                    "type": "str",
                    "required": False,
                    "description": "住所地",
                },
                {
                    "name": "third_party_id_number",
                    "label": "第三证件号码",
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
                    "label": "第三方法定代表人",
                    "type": "str",
                    "required": False,
                    "description": "法定代表人/负责人",
                },
                {
                    "name": "third_party_phone",
                    "label": "第三人联系电话",
                    "type": "str",
                    "required": False,
                    "description": "联系电话",
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
                
                # === 原告角色 ===
                {
                    "name": "plaintiff_role",
                    "label": "原告角色",
                    "type": "str",
                    "required": True,
                    "options": ["承包人或施工人", "发包人"],
                    "description": "原告是承包人/施工人还是发包人",
                },
                
                # === 诉讼请求（承包人/施工人） ===
                # 1. 支付工程款
                {
                    "name": "claim_project_payment",
                    "label": "支付工程款金额",
                    "type": "float",
                    "required": False,
                    "description": "请求支付的工程款金额（元）",
                },
                # 2. 迟延支付工程款的利息（违约金）
                {
                    "name": "claim_interest_until_date",
                    "label": "利息截止日期",
                    "type": "str",
                    "required": False,
                    "description": "截至日期",
                },
                {
                    "name": "claim_interest_amount",
                    "label": "利息金额",
                    "type": "float",
                    "required": False,
                    "description": "迟延支付工程款的利息金额",
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
                    "label": "逾期利息起始日期",
                    "type": "str",
                    "required": False,
                    "description": "起始日期",
                },
                {
                    "name": "claim_interest_penalty_base",
                    "label": "逾期利息基数",
                    "type": "float",
                    "required": False,
                    "description": "逾期利息基数",
                },
                {
                    "name": "claim_interest_penalty_rate",
                    "label": "逾期利息标准",
                    "type": "str",
                    "required": False,
                    "description": "如：全国银行间同业拆借中心公布的贷款市场报价利率",
                },
                {
                    "name": "claim_interest_to_actual",
                    "label": "是否请求支付至实际清偿之日止",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                # 3. 建设工程价款优先受偿权
                {
                    "name": "claim_priority_right",
                    "label": "是否主张建设工程价款优先受偿权",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_priority_content",
                    "label": "优先受偿权内容",
                    "type": "str",
                    "required": False,
                    "description": "优先受偿权的具体内容",
                },
                # 4. 突破合同相对性
                {
                    "name": "claim_bypass_liability",
                    "label": "是否请求突破合同相对性",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_bypass_parties",
                    "label": "责任主体名称",
                    "type": "str",
                    "required": False,
                    "description": "发包人、其他转包方、分包方等主体",
                },
                # 5. 赔偿损失
                {
                    "name": "claim_loss",
                    "label": "是否要求赔偿损失",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_loss_amount",
                    "label": "赔偿金金额",
                    "type": "float",
                    "required": False,
                    "description": "赔偿金金额",
                },
                {
                    "name": "claim_loss_type",
                    "label": "责任类型",
                    "type": "str",
                    "required": False,
                    "options": ["停窝工损失", "其他"],
                    "description": "责任类型",
                },
                {
                    "name": "claim_loss_detail",
                    "label": "具体情形",
                    "type": "str",
                    "required": False,
                    "description": "损失具体情形",
                },
                {
                    "name": "claim_loss_basis",
                    "label": "损失计算依据",
                    "type": "str",
                    "required": False,
                    "description": "损失计算依据",
                },
                
                # === 诉讼请求（发包人） ===
                # 6. 退还超付工程款
                {
                    "name": "claim_overpayment",
                    "label": "是否退还超付工程款",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_overpayment_amount",
                    "label": "超付金额",
                    "type": "float",
                    "required": False,
                    "description": "超付金额",
                },
                
                # === 诉讼请求（共同项7-15） ===
                # 7. 支付超付工程款利息
                {
                    "name": "claim_overpayment_interest",
                    "label": "是否支付超付工程款利息",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_overpayment_interest_until_date",
                    "label": "超付利息截止日期",
                    "type": "str",
                    "required": False,
                    "description": "截至日期",
                },
                {
                    "name": "claim_overpayment_interest_amount",
                    "label": "超付利息金额",
                    "type": "float",
                    "required": False,
                    "description": "利息金额",
                },
                {
                    "name": "claim_overpayment_interest_from_date",
                    "label": "超付利息起始日期",
                    "type": "str",
                    "required": False,
                    "description": "起始日期",
                },
                {
                    "name": "claim_overpayment_interest_base",
                    "label": "超付利息基数",
                    "type": "float",
                    "required": False,
                    "description": "基数",
                },
                {
                    "name": "claim_overpayment_interest_rate",
                    "label": "超付利息计算方式",
                    "type": "str",
                    "required": False,
                    "description": "计算方式",
                },
                # 8. 修复责任
                {
                    "name": "claim_repair",
                    "label": "是否对建设工程承担修复责任",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_repair_method",
                    "label": "修复方式",
                    "type": "str",
                    "required": False,
                    "description": "修复/付修复费用/减少工程款/其他",
                },
                {
                    "name": "claim_repair_amount",
                    "label": "修复金额/数额",
                    "type": "float",
                    "required": False,
                    "description": "数额",
                },
                # 9. 赔偿损失（发包人）
                {
                    "name": "claim_defect_loss",
                    "label": "是否要求赔偿损失（发包人）",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_defect_loss_amount",
                    "label": "赔偿金金额（发包人）",
                    "type": "float",
                    "required": False,
                    "description": "赔偿金金额",
                },
                {
                    "name": "claim_defect_loss_type",
                    "label": "责任类型（发包人）",
                    "type": "str",
                    "required": False,
                    "options": ["工程质量不符合约定", "迟延交付工程", "拒绝履行", "其他"],
                    "description": "责任类型",
                },
                {
                    "name": "claim_defect_loss_detail",
                    "label": "具体情形（发包人）",
                    "type": "str",
                    "required": False,
                    "description": "具体情形",
                },
                # 10. 合同无效
                {
                    "name": "claim_contract_invalid",
                    "label": "是否请求确认合同无效",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_contract_invalid_reason",
                    "label": "合同无效的理由",
                    "type": "str",
                    "required": False,
                    "description": "合同无效的理由",
                },
                # 11. 继续履行或解除
                {
                    "name": "claim_continue_or_dissolve",
                    "label": "继续履行还是解除",
                    "type": "str",
                    "required": False,
                    "options": ["继续履行", "解除合同"],
                    "description": "继续履行/解除合同",
                },
                {
                    "name": "claim_continue_deadline",
                    "label": "继续履行期限",
                    "type": "str",
                    "required": False,
                    "description": "如：30日内",
                },
                {
                    "name": "claim_continue_obligations",
                    "label": "继续履行义务",
                    "type": "str",
                    "required": False,
                    "description": "付款/竣工/其他义务",
                },
                {
                    "name": "claim_dissolve_date",
                    "label": "解除合同日期",
                    "type": "str",
                    "required": False,
                    "description": "确认合同解除的日期",
                },
                # 12. 债权实现费用
                {
                    "name": "claim_creditor_expenses",
                    "label": "是否主张实现债权费用",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "claim_creditor_expenses_detail",
                    "label": "费用明细",
                    "type": "str",
                    "required": False,
                    "description": "律师费、差旅费等",
                },
                # 13. 诉讼费用
                {
                    "name": "claim_litigation_fee",
                    "label": "是否主张诉讼费用",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                # 14. 其他请求
                {
                    "name": "claim_other",
                    "label": "其他请求",
                    "type": "str",
                    "required": False,
                    "description": "其他诉讼请求",
                },
                # 15. 标的总额
                {
                    "name": "total_amount",
                    "label": "标的总额",
                    "type": "float",
                    "required": True,
                    "description": "诉讼标的金额",
                },
                
                # === 约定管辖和诉前保全 ===
                {
                    "name": "has_jurisdiction_agreement",
                    "label": "是否有管辖约定",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
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
                    "label": "是否已经诉前保全",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "preservation_court",
                    "label": "保全法院",
                    "type": "str",
                    "required": False,
                    "description": "保全法院",
                },
                {
                    "name": "preservation_date",
                    "label": "保全时间",
                    "type": "str",
                    "required": False,
                    "description": "保全时间",
                },
                {
                    "name": "preservation_case_no",
                    "label": "保全案号",
                    "type": "str",
                    "required": False,
                    "description": "保全案号",
                },
                {
                    "name": "has_appraisal",
                    "label": "是否申请鉴定",
                    "type": "bool",
                    "required": False,
                    "default": False,
                    "description": "是/否",
                },
                {
                    "name": "appraisal_matters",
                    "label": "鉴定事项",
                    "type": "str",
                    "required": False,
                    "description": "鉴定事项",
                },
                
                # === 事实与理由 ===
                # 1. 合同签订情况
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
                    "label": "签订时间",
                    "type": "str",
                    "required": False,
                    "description": "签订时间",
                },
                {
                    "name": "contract_sign_place",
                    "label": "签订地点",
                    "type": "str",
                    "required": False,
                    "description": "签订地点",
                },
                {
                    "name": "contract_bidding",
                    "label": "是否招投标",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                # 2. 签订主体
                {
                    "name": "employer_name",
                    "label": "发包人",
                    "type": "str",
                    "required": False,
                    "description": "发包人名称",
                },
                {
                    "name": "contractor_name",
                    "label": "承包人",
                    "type": "str",
                    "required": False,
                    "description": "承包人名称",
                },
                {
                    "name": "资质出借企业",
                    "label": "出借资质的建筑企业",
                    "type": "str",
                    "required": False,
                    "description": "出借资质的建筑企业",
                },
                # 3. 建设工程情况
                {
                    "name": "project_name",
                    "label": "工程名称",
                    "type": "str",
                    "required": False,
                    "description": "工程名称",
                },
                {
                    "name": "project_location",
                    "label": "工程地点",
                    "type": "str",
                    "required": False,
                    "description": "工程所在地点",
                },
                {
                    "name": "project_scope",
                    "label": "施工范围",
                    "type": "str",
                    "required": False,
                    "description": "施工范围",
                },
                {
                    "name": "project_quality_standard",
                    "label": "质量标准",
                    "type": "str",
                    "required": False,
                    "description": "质量标准",
                },
                # 4. 工程款及支付方式
                {
                    "name": "project_price_type",
                    "label": "工程款计价方式",
                    "type": "str",
                    "required": False,
                    "options": ["综合单价", "固定单价", "其他"],
                    "description": "综合单价/固定单价/其他",
                },
                {
                    "name": "project_unit_price",
                    "label": "综合单价",
                    "type": "float",
                    "required": False,
                    "description": "综合单价金额",
                },
                {
                    "name": "project_fixed_price",
                    "label": "固定单价",
                    "type": "float",
                    "required": False,
                    "description": "固定单价金额",
                },
                {
                    "name": "project_other_price",
                    "label": "其他计价",
                    "type": "str",
                    "required": False,
                    "description": "其他计价方式说明",
                },
                # 5. 工期
                {
                    "name": "project_start_date",
                    "label": "开工时间",
                    "type": "str",
                    "required": False,
                    "description": "开工时间",
                },
                {
                    "name": "project_end_date",
                    "label": "竣工时间",
                    "type": "str",
                    "required": False,
                    "description": "竣工时间",
                },
                {
                    "name": "project_duration",
                    "label": "工期天数",
                    "type": "int",
                    "required": False,
                    "description": "工期天数",
                },
                # 6. 工程质量标准及竣工验收
                {
                    "name": "quality_acceptance_standard",
                    "label": "工程质量标准及竣工验收程序",
                    "type": "str",
                    "required": False,
                    "description": "工程质量标准及竣工验收程序",
                },
                # 7. 违约金/保证金
                {
                    "name": "has_contract_penalty",
                    "label": "是否有违约金约定",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "contract_penalty_amount",
                    "label": "违约金金额",
                    "type": "float",
                    "required": False,
                    "description": "违约金金额",
                },
                {
                    "name": "contract_penalty_article",
                    "label": "违约金条款",
                    "type": "str",
                    "required": False,
                    "description": "如：第X条",
                },
                {
                    "name": "has_contract_deposit",
                    "label": "是否有保证金约定",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "contract_deposit_amount",
                    "label": "保证金金额",
                    "type": "float",
                    "required": False,
                    "description": "保证金金额",
                },
                {
                    "name": "contract_deposit_article",
                    "label": "保证金条款",
                    "type": "str",
                    "required": False,
                    "description": "如：第X条",
                },
                {
                    "name": "has_delay_penalty",
                    "label": "是否有迟延履行违约金",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "delay_penalty_rate",
                    "label": "迟延履行违约金比例",
                    "type": "str",
                    "required": False,
                    "description": "如：0.5%/日",
                },
                {
                    "name": "delay_penalty_article",
                    "label": "迟延履行违约金条款",
                    "type": "str",
                    "required": False,
                    "description": "如：第X条",
                },
                # 8. 工程款支付情况
                {
                    "name": "project_total_price",
                    "label": "工程总价",
                    "type": "float",
                    "required": False,
                    "description": "工程总价",
                },
                {
                    "name": "project_paid_amount",
                    "label": "已支付工程款",
                    "type": "float",
                    "required": False,
                    "description": "已支付工程款",
                },
                {
                    "name": "project_owed_amount",
                    "label": "欠/超付工程款",
                    "type": "float",
                    "required": False,
                    "description": "欠/超付工程款金额",
                },
                {
                    "name": "project_owed_interest",
                    "label": "欠/超付工程款利息",
                    "type": "float",
                    "required": False,
                    "description": "利息金额",
                },
                # 9. 建设工程质量情况
                {
                    "name": "quality_qualified",
                    "label": "工程质量是否合格",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "quality_issues",
                    "label": "工程质量问题",
                    "type": "str",
                    "required": False,
                    "description": "工程质量问题描述",
                },
                {
                    "name": "quality_loss_amount",
                    "label": "工程质量造成损失",
                    "type": "float",
                    "required": False,
                    "description": "损失金额",
                },
                # 10. 建设工程交付情况
                {
                    "name": "delivery_delayed",
                    "label": "工程是否迟延交付",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "actual_delivery_date",
                    "label": "实际交付时间",
                    "type": "str",
                    "required": False,
                    "description": "实际交付时间",
                },
                {
                    "name": "delivery_delay_loss",
                    "label": "迟延交付造成损失",
                    "type": "float",
                    "required": False,
                    "description": "损失金额",
                },
                # 11. 停窝工情况
                {
                    "name": "has_suspension",
                    "label": "是否停窝工",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "suspension_loss",
                    "label": "停窝工造成损失",
                    "type": "float",
                    "required": False,
                    "description": "损失金额",
                },
                # 12. 优先受偿权主张情况
                {
                    "name": "priority_claimed",
                    "label": "是否主张过建设工程价款优先受偿权",
                    "type": "bool",
                    "required": False,
                    "description": "是/否",
                },
                {
                    "name": "priority_claim_date",
                    "label": "主张时间",
                    "type": "str",
                    "required": False,
                    "description": "主张时间",
                },
                {
                    "name": "priority_claim_method",
                    "label": "主张方式",
                    "type": "str",
                    "required": False,
                    "description": "通过什么方式主张",
                },
                # 13. 其他需要说明的内容
                {
                    "name": "other_matters",
                    "label": "其他需要说明的内容",
                    "type": "str",
                    "required": False,
                    "description": "其他需要说明的内容",
                },
                # 14. 请求依据
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
                
                # === 调解意愿 ===
                {
                    "name": "understand_mediation",
                    "label": "是否了解调解",
                    "type": "bool",
                    "required": False,
                    "default": True,
                    "description": "了解/不了解",
                },
                {
                    "name": "understand_mediation_benefits",
                    "label": "是否了解先行调解好处",
                    "type": "bool",
                    "required": False,
                    "default": True,
                    "description": "了解/不了解",
                },
                {
                    "name": "consider_mediation",
                    "label": "是否考虑先行调解",
                    "type": "str",
                    "required": False,
                    "options": ["是", "否", "暂不确定"],
                    "description": "是/否/暂不确定",
                },
            ]
        }
    
    def analyze_case_data(self, case_data: dict) -> dict:
        """预处理案件数据"""
        # 确定原告类型（自然人/法人）
        if 'plaintiff_type' not in case_data:
            plaintiff_id = case_data.get('plaintiff_id_number', '')
            if len(plaintiff_id) == 18 and plaintiff_id[:6].isdigit():
                case_data['plaintiff_type'] = '自然人'
            else:
                case_data['plaintiff_type'] = '法人'
        
        # 确定被告类型
        if 'defendant_type' not in case_data:
            defendant_id = case_data.get('defendant_id_number', '')
            if len(defendant_id) == 18 and defendant_id[:6].isdigit():
                case_data['defendant_type'] = '自然人'
            else:
                case_data['defendant_type'] = '法人'
        
        # 确定第三人类型
        if case_data.get('has_third_party') and 'third_party_type' not in case_data:
            third_party_id = case_data.get('third_party_id_number', '')
            if len(third_party_id) == 18 and third_party_id[:6].isdigit():
                case_data['third_party_type'] = '自然人'
            else:
                case_data['third_party_type'] = '法人'
        
        return case_data
    
    def calculate(self, case_data: dict) -> dict:
        """计算诉讼请求相关金额"""
        return {
            "items": [],
            "total": case_data.get('total_amount', 0),
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
                fill_map['T1_16_第三证件号码'] = case_data.get('third_party_id_number', '')
                fill_map['T1_17_第三人工作单位'] = case_data.get('third_party_work_unit', '')
                fill_map['T1_18_第三人职务'] = case_data.get('third_party_position', '')
            else:
                # T1 Row 4: 第三人法人
                fill_map['T1_14_第三人名称'] = case_data.get('third_party_name', '')
                fill_map['T1_15_第三人住所地'] = case_data.get('third_party_address', '')
                fill_map['T1_16_第三人注册地'] = case_data.get('third_party_registration_address', '')
                fill_map['T1_17_第三方法定代表人'] = case_data.get('third_party_legal_rep', '')
                fill_map['T1_18_第三人联系电话'] = case_data.get('third_party_phone', '')
                fill_map['T1_19_第三人统一社会信用代码'] = case_data.get('third_party_id_number', '')
                fill_map['T1_20_第三人公司类型'] = case_data.get('third_party_company_type', '')
                fill_map['T1_21_第三人所有制性质'] = case_data.get('third_party_ownership', '')
        
        # === T2: 诉讼请求 (3列结构！) ===
        # T2 Row 3: 1. 支付工程款 - C2列填写金额
        fill_map['T2_01_支付工程款金额'] = str(int(case_data.get('claim_project_payment', 0)))
        
        # T2 Row 4: 2. 迟延支付工程款的利息（违约金）
        fill_map['T2_02_利息截止日期'] = case_data.get('claim_interest_until_date', '')
        fill_map['T2_03_利息金额'] = str(int(case_data.get('claim_interest_amount', 0)))
        fill_map['T2_04_违约金金额'] = str(int(case_data.get('claim_penalty_amount', 0)))
        fill_map['T2_05_逾期利息起始日期'] = case_data.get('claim_interest_penalty_from_date', '')
        fill_map['T2_06_逾期利息基数'] = str(int(case_data.get('claim_interest_penalty_base', 0)))
        fill_map['T2_07_逾期利息标准'] = case_data.get('claim_interest_penalty_rate', '')
        
        # T2 Row 5: 3. 建设工程价款优先受偿权
        fill_map['T2_08_优先受偿权内容'] = case_data.get('claim_priority_content', '')
        
        # T2 Row 6: 4. 突破合同相对性
        fill_map['T2_09_责任主体名称'] = case_data.get('claim_bypass_parties', '')
        
        # T2 Row 7: 5. 赔偿损失
        fill_map['T2_10_赔偿金金额'] = str(int(case_data.get('claim_loss_amount', 0)))
        fill_map['T2_11_损失具体情形'] = case_data.get('claim_loss_detail', '')
        fill_map['T2_12_损失计算依据'] = case_data.get('claim_loss_basis', '')
        
        # T2 Row 8: 6. 退还超付工程款
        fill_map['T2_13_超付金额'] = str(int(case_data.get('claim_overpayment_amount', 0)))
        
        # === T3: 诉讼请求 7-15项 ===
        # T3 Row 0: 7. 是否支付超付工程款利息
        fill_map['T3_01_超付利息截止日期'] = case_data.get('claim_overpayment_interest_until_date', '')
        fill_map['T3_02_超付利息金额'] = str(int(case_data.get('claim_overpayment_interest_amount', 0)))
        fill_map['T3_03_超付利息起始日期'] = case_data.get('claim_overpayment_interest_from_date', '')
        fill_map['T3_04_超付利息基数'] = str(int(case_data.get('claim_overpayment_interest_base', 0)))
        fill_map['T3_05_超付利息计算方式'] = case_data.get('claim_overpayment_interest_rate', '')
        
        # T3 Row 1: 8. 是否对建设工程承担修复责任
        fill_map['T3_06_修复数额'] = str(int(case_data.get('claim_repair_amount', 0)))
        
        # T3 Row 2: 9. 是否要求赔偿损失（发包人）
        fill_map['T3_07_赔偿金金额'] = str(int(case_data.get('claim_defect_loss_amount', 0)))
        fill_map['T3_08_具体情形'] = case_data.get('claim_defect_loss_detail', '')
        
        # T3 Row 3: 10. 请求确认建设工程施工合同无效
        fill_map['T3_09_合同无效理由'] = case_data.get('claim_contract_invalid_reason', '')
        
        # T3 Row 4: 11. 要求继续履行或是解除合同
        fill_map['T3_10_继续履行期限'] = case_data.get('claim_continue_deadline', '')
        fill_map['T3_11_继续履行义务'] = case_data.get('claim_continue_obligations', '')
        fill_map['T3_12_解除合同日期'] = case_data.get('claim_dissolve_date', '')
        
        # T3 Row 5: 12. 是否主张实现债权的费用
        fill_map['T3_13_债权费用明细'] = case_data.get('claim_creditor_expenses_detail', '')
        
        # T3 Row 7: 14. 其他请求
        fill_map['T3_14_其他请求'] = case_data.get('claim_other', '')
        
        # T3 Row 8: 15. 标的总额
        fill_map['T3_15_标的总额'] = str(int(case_data.get('total_amount', 0)))
        
        # === T3: 约定管辖和诉前保全 ===
        # T3 Row 10: 1. 有无仲裁、法院管辖约定
        fill_map['T3_16_管辖条款内容'] = case_data.get('jurisdiction_clause', '')
        
        # T3 Row 11: 2. 是否已经诉前保全
        fill_map['T3_17_保全法院'] = case_data.get('preservation_court', '')
        fill_map['T3_18_保全时间'] = case_data.get('preservation_date', '')
        fill_map['T3_19_保全案号'] = case_data.get('preservation_case_no', '')
        
        # T3 Row 12: 3. 是否申请鉴定
        fill_map['T3_20_鉴定事项'] = case_data.get('appraisal_matters', '')
        
        # === T4: 合同详情和工程款详情 ===
        # T4 Row 0: 1. 合同的签订情况
        contract_info = []
        if case_data.get('contract_name'):
            contract_info.append(f"合同名称：{case_data['contract_name']}")
        if case_data.get('contract_number'):
            contract_info.append(f"合同编号：{case_data['contract_number']}")
        if case_data.get('contract_sign_date'):
            contract_info.append(f"签订时间：{case_data['contract_sign_date']}")
        if case_data.get('contract_sign_place'):
            contract_info.append(f"签订地点：{case_data['contract_sign_place']}")
        if case_data.get('contract_bidding'):
            contract_info.append("经过招投标程序")
        fill_map['T4_01_合同签订情况'] = '；'.join(contract_info)
        
        # T4 Row 1: 2. 签订主体
        fill_map['T4_02_发包人'] = case_data.get('employer_name', '')
        fill_map['T4_03_承包人'] = case_data.get('contractor_name', '')
        fill_map['T4_04_出借资质企业'] = case_data.get('资质出借企业', '')
        
        # T4 Row 2: 3. 建设工程情况
        project_info = []
        if case_data.get('project_name'):
            project_info.append(f"工程名称：{case_data['project_name']}")
        if case_data.get('project_location'):
            project_info.append(f"工程地点：{case_data['project_location']}")
        if case_data.get('project_scope'):
            project_info.append(f"施工范围：{case_data['project_scope']}")
        if case_data.get('project_quality_standard'):
            project_info.append(f"质量标准：{case_data['project_quality_standard']}")
        fill_map['T4_05_建设工程情况'] = '；'.join(project_info)
        
        # T4 Row 3: 4. 合同约定的工程款及支付方式
        price_type = case_data.get('project_price_type', '')
        if price_type == '综合单价':
            fill_map['T4_06_综合单价'] = str(int(case_data.get('project_unit_price', 0)))
        elif price_type == '固定单价':
            fill_map['T4_07_固定单价'] = str(int(case_data.get('project_fixed_price', 0)))
        else:
            fill_map['T4_08_其他计价方式'] = case_data.get('project_other_price', '')
        
        # T4 Row 4: 5. 工期
        fill_map['T4_09_开工时间'] = case_data.get('project_start_date', '')
        fill_map['T4_10_竣工时间'] = case_data.get('project_end_date', '')
        fill_map['T4_11_工期天数'] = str(case_data.get('project_duration', ''))
        
        # T4 Row 5: 6. 合同约定的工程质量标准及竣工验收程序
        fill_map['T4_12_质量验收标准'] = case_data.get('quality_acceptance_standard', '')
        
        # T4 Row 6: 7. 合同约定的违约金（保证金）
        fill_map['T4_13_违约金金额'] = str(int(case_data.get('contract_penalty_amount', 0)))
        fill_map['T4_14_违约金条款'] = case_data.get('contract_penalty_article', '')
        fill_map['T4_15_保证金金额'] = str(int(case_data.get('contract_deposit_amount', 0)))
        fill_map['T4_16_保证金条款'] = case_data.get('contract_deposit_article', '')
        fill_map['T4_17_迟延履行违约金比例'] = case_data.get('delay_penalty_rate', '')
        fill_map['T4_18_迟延履行违约金条款'] = case_data.get('delay_penalty_article', '')
        
        # T4 Row 7: 8. 工程款支付情况
        fill_map['T4_19_工程总价'] = str(int(case_data.get('project_total_price', 0)))
        fill_map['T4_20_已支付工程款'] = str(int(case_data.get('project_paid_amount', 0)))
        fill_map['T4_21_欠超付工程款'] = str(int(case_data.get('project_owed_amount', 0)))
        fill_map['T4_22_欠超付工程款利息'] = str(int(case_data.get('project_owed_interest', 0)))
        
        # T4 Row 8: 9. 建设工程质量情况
        fill_map['T4_23_工程质量问题'] = case_data.get('quality_issues', '')
        fill_map['T4_24_工程质量造成损失'] = str(int(case_data.get('quality_loss_amount', 0)))
        
        # T4 Row 9: 10. 建设工程交付情况
        fill_map['T4_25_实际交付时间'] = case_data.get('actual_delivery_date', '')
        fill_map['T4_26_迟延交付造成损失'] = str(int(case_data.get('delivery_delay_loss', 0)))
        
        # T4 Row 10: 11. 停窝工等情况
        fill_map['T4_27_停窝工造成损失'] = str(int(case_data.get('suspension_loss', 0)))
        
        # T4 Row 11: 12. 是否主张过建设工程价款优先受偿权
        fill_map['T4_28_优先受偿权主张时间'] = case_data.get('priority_claim_date', '')
        fill_map['T4_29_优先受偿权主张方式'] = case_data.get('priority_claim_method', '')
        
        # T4 Row 12: 13. 其他需要说明的内容
        fill_map['T4_30_其他说明内容'] = case_data.get('other_matters', '')
        
        # T4 Row 13: 14. 请求依据
        fill_map['T4_31_合同约定依据'] = case_data.get('contract_basis', '')
        fill_map['T4_32_法律规定依据'] = case_data.get('legal_basis', '')
        
        return fill_map
    
    def get_checkbox_map(self, case_data: dict) -> dict[str, bool]:
        """生成勾选框映射表"""
        checkbox_map = {}
        
        # === T0: 原告类型勾选 ===
        plaintiff_type = case_data.get('plaintiff_type', '自然人')
        checkbox_map['T0_原告类型_自然人'] = (plaintiff_type == '自然人')
        checkbox_map['T0_原告类型_法人'] = (plaintiff_type == '法人')
        
        # 原告性别
        plaintiff_gender = case_data.get('plaintiff_gender', '')
        checkbox_map['T0_性别_男'] = (plaintiff_gender == '男')
        checkbox_map['T0_性别_女'] = (plaintiff_gender == '女')
        
        # 原告公司类型
        plaintiff_company_type = case_data.get('plaintiff_company_type', '')
        checkbox_map['T0_公司类型_有限责任公司'] = ('有限' in plaintiff_company_type)
        checkbox_map['T0_公司类型_股份有限公司'] = ('股份' in plaintiff_company_type)
        checkbox_map['T0_公司类型_上市公司'] = ('上市' in plaintiff_company_type)
        
        # 原告所有制性质
        plaintiff_ownership = case_data.get('plaintiff_ownership', '')
        checkbox_map['T0_所有制_国有'] = (plaintiff_ownership == '国有')
        checkbox_map['T0_所有制_民营'] = (plaintiff_ownership == '民营')
        checkbox_map['T0_所有制_其他'] = (plaintiff_ownership == '其他')
        
        # === T1: 代理人 ===
        has_agent = case_data.get('has_agent', False)
        checkbox_map['T1_代理人_有'] = has_agent
        checkbox_map['T1_代理人_无'] = not has_agent
        
        # === T1: 被告类型勾选 ===
        defendant_type = case_data.get('defendant_type', '自然人')
        checkbox_map['T1_被告类型_自然人'] = (defendant_type == '自然人')
        checkbox_map['T1_被告类型_法人'] = (defendant_type == '法人')
        
        # 被告性别
        defendant_gender = case_data.get('defendant_gender', '')
        checkbox_map['T1_性别_男'] = (defendant_gender == '男')
        checkbox_map['T1_性别_女'] = (defendant_gender == '女')
        
        # 被告公司类型
        defendant_company_type = case_data.get('defendant_company_type', '')
        checkbox_map['T1_公司类型_有限责任公司'] = ('有限' in defendant_company_type)
        checkbox_map['T1_公司类型_股份有限公司'] = ('股份' in defendant_company_type)
        checkbox_map['T1_公司类型_上市公司'] = ('上市' in defendant_company_type)
        
        # 被告所有制性质
        defendant_ownership = case_data.get('defendant_ownership', '')
        checkbox_map['T1_所有制_国有'] = (defendant_ownership == '国有')
        checkbox_map['T1_所有制_民营'] = (defendant_ownership == '民营')
        checkbox_map['T1_所有制_其他'] = (defendant_ownership == '其他')
        
        # === T1: 第三人类型勾选 ===
        if case_data.get('has_third_party'):
            tp_type = case_data.get('third_party_type', '自然人')
            checkbox_map['T1_第三人类型_自然人'] = (tp_type == '自然人')
            checkbox_map['T1_第三人类型_法人'] = (tp_type == '法人')
            
            third_party_gender = case_data.get('third_party_gender', '')
            checkbox_map['T1_第三人性别_男'] = (third_party_gender == '男')
            checkbox_map['T1_第三人性别_女'] = (third_party_gender == '女')
            
            tp_company_type = case_data.get('third_party_company_type', '')
            checkbox_map['T1_第三人公司类型_有限责任公司'] = ('有限' in tp_company_type)
            checkbox_map['T1_第三人公司类型_股份有限公司'] = ('股份' in tp_company_type)
            
            tp_ownership = case_data.get('third_party_ownership', '')
            checkbox_map['T1_第三人所有制_国有'] = (tp_ownership == '国有')
            checkbox_map['T1_第三人所有制_民营'] = (tp_ownership == '民营')
            checkbox_map['T1_第三人所有制_其他'] = (tp_ownership == '其他')
        
        # === T2: 诉讼请求勾选 ===
        # 利息/违约金是否请求支付至实际清偿之日止
        interest_to_actual = case_data.get('claim_interest_to_actual', False)
        checkbox_map['T2_利息至清偿_是'] = interest_to_actual
        checkbox_map['T2_利息至清偿_否'] = not interest_to_actual
        
        # 建设工程价款优先受偿权
        claim_priority = case_data.get('claim_priority_right', False)
        checkbox_map['T2_优先受偿_是'] = claim_priority
        checkbox_map['T2_优先受偿_否'] = not claim_priority
        
        # 突破合同相对性
        claim_bypass = case_data.get('claim_bypass_liability', False)
        checkbox_map['T2_突破相对性_是'] = claim_bypass
        checkbox_map['T2_突破相对性_否'] = not claim_bypass
        
        # 赔偿损失
        claim_loss = case_data.get('claim_loss', False)
        checkbox_map['T2_赔偿损失_是'] = claim_loss
        checkbox_map['T2_赔偿损失_否'] = not claim_loss
        
        # 赔偿损失类型
        claim_loss_type = case_data.get('claim_loss_type', '')
        checkbox_map['T2_损失类型_停窝工损失'] = (claim_loss_type == '停窝工损失')
        checkbox_map['T2_损失类型_其他'] = (claim_loss_type == '其他')
        
        # 退还超付工程款
        claim_overpayment = case_data.get('claim_overpayment', False)
        checkbox_map['T2_退还超付_是'] = claim_overpayment
        checkbox_map['T2_退还超付_否'] = not claim_overpayment
        
        # === T3: 诉讼请求 7-15项勾选 ===
        # 7. 支付超付工程款利息
        overpayment_interest = case_data.get('claim_overpayment_interest', False)
        checkbox_map['T3_超付利息_是'] = overpayment_interest
        checkbox_map['T3_超付利息_否'] = not overpayment_interest
        
        # 8. 修复责任
        claim_repair = case_data.get('claim_repair', False)
        checkbox_map['T3_修复责任_是'] = claim_repair
        checkbox_map['T3_修复责任_否'] = not claim_repair
        
        # 修复方式
        repair_method = case_data.get('claim_repair_method', '')
        checkbox_map['T3_修复方式_修复'] = ('修复' in repair_method)
        checkbox_map['T3_修复方式_付修复费用'] = ('费用' in repair_method)
        checkbox_map['T3_修复方式_减少工程款'] = ('减少' in repair_method)
        checkbox_map['T3_修复方式_其他'] = ('其他' in repair_method)
        
        # 9. 赔偿损失（发包人）
        defect_loss = case_data.get('claim_defect_loss', False)
        checkbox_map['T3_发包人赔偿_是'] = defect_loss
        checkbox_map['T3_发包人赔偿_否'] = not defect_loss
        
        # 发包人损失类型
        defect_loss_type = case_data.get('claim_defect_loss_type', '')
        checkbox_map['T3_发包人损失类型_工程质量不符合约定'] = ('工程质量' in defect_loss_type)
        checkbox_map['T3_发包人损失类型_迟延交付工程'] = ('迟延交付' in defect_loss_type)
        checkbox_map['T3_发包人损失类型_拒绝履行'] = ('拒绝履行' in defect_loss_type)
        checkbox_map['T3_发包人损失类型_其他'] = ('其他' in defect_loss_type)
        
        # 10. 合同无效
        contract_invalid = case_data.get('claim_contract_invalid', False)
        checkbox_map['T3_合同无效_是'] = contract_invalid
        checkbox_map['T3_合同无效_否'] = not contract_invalid
        
        # 11. 继续履行或解除
        continue_or_dissolve = case_data.get('claim_continue_or_dissolve', '')
        checkbox_map['T3_继续履行_继续履行'] = (continue_or_dissolve == '继续履行')
        checkbox_map['T3_继续履行_解除合同'] = (continue_or_dissolve == '解除合同')
        
        # 12. 债权实现费用
        creditor_expenses = case_data.get('claim_creditor_expenses', False)
        checkbox_map['T3_债权费用_是'] = creditor_expenses
        checkbox_map['T3_债权费用_否'] = not creditor_expenses
        
        # 13. 诉讼费用
        litigation_fee = case_data.get('claim_litigation_fee', False)
        checkbox_map['T3_诉讼费用_是'] = litigation_fee
        checkbox_map['T3_诉讼费用_否'] = not litigation_fee
        
        # === T3: 约定管辖和诉前保全勾选 ===
        # 1. 有无管辖约定
        has_jurisdiction = case_data.get('has_jurisdiction_agreement', False)
        checkbox_map['T3_管辖约定_有'] = has_jurisdiction
        checkbox_map['T3_管辖约定_无'] = not has_jurisdiction
        
        # 2. 是否诉前保全
        has_preservation = case_data.get('has_preservation', False)
        checkbox_map['T3_诉前保全_是'] = has_preservation
        checkbox_map['T3_诉前保全_否'] = not has_preservation
        
        # 3. 是否申请鉴定
        has_appraisal = case_data.get('has_appraisal', False)
        checkbox_map['T3_鉴定申请_是'] = has_appraisal
        checkbox_map['T3_鉴定申请_否'] = not has_appraisal
        
        # === T4: 合同详情勾选 ===
        # 4. 工程款计价方式
        price_type = case_data.get('project_price_type', '')
        checkbox_map['T4_计价方式_综合单价'] = (price_type == '综合单价')
        checkbox_map['T4_计价方式_固定单价'] = (price_type == '固定单价')
        checkbox_map['T4_计价方式_其他'] = (price_type == '其他')
        
        # 7. 违约金/保证金约定
        checkbox_map['T4_违约金_有'] = case_data.get('has_contract_penalty', False)
        checkbox_map['T4_保证金_有'] = case_data.get('has_contract_deposit', False)
        checkbox_map['T4_迟延违约金_有'] = case_data.get('has_delay_penalty', False)
        
        # 9. 工程质量情况
        quality_qualified = case_data.get('quality_qualified', True)
        checkbox_map['T4_质量合格_是'] = quality_qualified
        checkbox_map['T4_质量合格_否'] = not quality_qualified
        
        # 10. 交付情况
        delivery_delayed = case_data.get('delivery_delayed', False)
        checkbox_map['T4_迟延交付_是'] = delivery_delayed
        checkbox_map['T4_迟延交付_否'] = not delivery_delayed
        
        # 11. 停窝工情况
        has_suspension = case_data.get('has_suspension', False)
        checkbox_map['T4_停窝工_是'] = has_suspension
        checkbox_map['T4_停窝工_否'] = not has_suspension
        
        # 12. 优先受偿权主张
        priority_claimed = case_data.get('priority_claimed', False)
        checkbox_map['T4_优先受偿权_是'] = priority_claimed
        checkbox_map['T4_优先受偿权_否'] = not priority_claimed
        
        # === T5: 调解意愿 ===
        understand_mediation = case_data.get('understand_mediation', True)
        checkbox_map['T5_了解调解_了解'] = understand_mediation
        checkbox_map['T5_了解调解_不了解'] = not understand_mediation
        
        understand_benefits = case_data.get('understand_mediation_benefits', True)
        checkbox_map['T5_了解先行调解_了解'] = understand_benefits
        checkbox_map['T5_了解先行调解_不了解'] = not understand_benefits
        
        consider_mediation = case_data.get('consider_mediation', '是')
        checkbox_map['T5_先行调解_是'] = (consider_mediation == '是')
        checkbox_map['T5_先行调解_否'] = (consider_mediation == '否')
        checkbox_map['T5_先行调解_暂不确定'] = (consider_mediation == '暂不确定')
        
        return checkbox_map
    
    def get_evidence_rules(self) -> list:
        """返回证据规则列表"""
        return [
            {"name": "建设工程施工合同", "required": True, "description": "证明双方存在建设工程施工合同关系"},
            {"name": "工程竣工验收材料", "required": False, "description": "证明工程竣工验收情况"},
            {"name": "工程款支付凭证", "required": False, "description": "证明工程款支付情况"},
            {"name": "结算文件", "required": False, "description": "工程结算相关文件"},
            {"name": "催款函件", "required": False, "description": "证明曾催告支付工程款"},
            {"name": "资质证书", "required": False, "description": "建筑业企业资质证书"},
            {"name": "招投标文件", "required": False, "description": "如经过招投标程序"},
        ]
