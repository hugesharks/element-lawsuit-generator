# -*- coding: utf-8 -*-
"""
交通事故案件证据规则适配器

从 /tmp/traffic-accident-workflow-git/scripts/evidence_list_generator.py 迁移证据规则库，
适配新的 EvidenceRule 格式。

证据规则涵盖 10 个类别、27 项证据规则：
1. 通用证据（5项）：所有案件都需要的基础证据
2. 医疗类证据（5项）：门诊/住院病历、诊断证明、费用清单等
3. 护理费证据（2项）：护理人证明、护理证明
4. 误工费证据（3项）：误工证明、劳动合同、工资流水
5. 交通费证据（1项）：交通费票据
6. 伤残鉴定证据（2项）：鉴定意见书、鉴定费发票
7. 死亡案件证据（5项）：死亡证明、殡葬证、继承关系证明等
8. 财产损失证据（2项）：定损单、维修发票
9. 被扶养人证据（3项）：身份证明、关系证明、无劳动能力证明
10. 刑事附带民事证据（1项）：刑事判决书

每条证据规则包含：
- name: 证据名称
- source: 证据来源
- purpose: 证明目的
- category: 证据类别
- condition: 触发条件（基于案件数据的字段判断）
- copies: 份数（默认1）
"""

from typing import Dict, List, Any


class EvidenceRule:
    """
    证据规则数据类

    定义单条证据规则的结构化信息，用于证据清单的自动生成。
    """

    def __init__(self, name: str, source: str, purpose: str,
                 category: str = 'common', condition_field: str = None,
                 condition_value: Any = None, condition_type: str = 'gt',
                 copies: int = 1):
        """
        初始化证据规则

        Args:
            name: 证据名称
            source: 证据来源（如"原告"、"就诊医院"、"公安交管部门"）
            purpose: 证明目的
            category: 证据类别
            condition_field: 触发条件依赖的案件数据字段名
            condition_value: 触发条件的目标值
            condition_type: 条件类型（'gt'=大于, 'eq'=等于, 'len_gt'=长度大于, 'always'=始终触发）
            copies: 份数
        """
        self.name = name
        self.source = source
        self.purpose = purpose
        self.category = category
        self.condition_field = condition_field
        self.condition_value = condition_value
        self.condition_type = condition_type
        self.copies = copies

    def is_applicable(self, case_data: dict) -> bool:
        """
        判断该证据规则是否适用于当前案件

        根据案件数据中的字段值判断是否需要此证据。

        Args:
            case_data: 案件数据字典

        Returns:
            bool: 是否适用
        """
        # 始终触发的证据
        if self.condition_type == 'always':
            return True

        # 无条件字段时默认适用
        if self.condition_field is None:
            return True

        field_value = case_data.get(self.condition_field)

        if self.condition_type == 'gt':
            # 数值大于目标值
            if field_value is None:
                return False
            return field_value > self.condition_value

        elif self.condition_type == 'eq':
            # 等于目标值
            return field_value == self.condition_value

        elif self.condition_type == 'len_gt':
            # 数组长度大于目标值
            if field_value is None:
                return False
            return len(field_value) > self.condition_value

        elif self.condition_type == 'not_none':
            # 字段不为None
            return field_value is not None

        return True

    def to_dict(self) -> dict:
        """
        转换为字典格式

        Returns:
            dict: 证据规则字典
        """
        return {
            'name': self.name,
            'source': self.source,
            'purpose': self.purpose,
            'category': self.category,
            'copies': self.copies,
        }


class TrafficAccidentEvidenceAdapter:
    """
    交通事故证据规则适配器

    管理27项证据规则，支持：
    - 获取全部规则
    - 根据案件数据推导适用的证据清单
    - 按类别分组查询
    """

    def __init__(self):
        """初始化证据规则库"""
        self._rules = self._build_rules()

    def _build_rules(self) -> List[EvidenceRule]:
        """
        构建完整的证据规则库（27项）

        从原有 evidence_list_generator.py 的 EVIDENCE_RULES 迁移，
        将 lambda 条件函数转换为结构化的条件字段。

        Returns:
            list[EvidenceRule]: 证据规则列表
        """
        rules = []

        # ============================
        # 通用证据（所有案件都需要，5项）
        # ============================
        rules.append(EvidenceRule(
            name="原告身份证复印件",
            source="原告",
            purpose="证明原告主体资格",
            category="common",
            condition_type="always",
        ))
        rules.append(EvidenceRule(
            name="道路交通事故认定书",
            source="公安交管部门",
            purpose="证明事故发生经过及责任划分",
            category="common",
            condition_type="always",
        ))
        rules.append(EvidenceRule(
            name="被告驾驶证、行驶证复印件",
            source="公安交管部门",
            purpose="证明被告身份及驾驶资格",
            category="common",
            condition_type="always",
        ))
        rules.append(EvidenceRule(
            name="机动车交强险保单",
            source="保险公司",
            purpose="证明肇事车辆投保交强险的事实",
            category="common",
            condition_type="always",
        ))
        rules.append(EvidenceRule(
            name="机动车商业三者险保单",
            source="保险公司",
            purpose="证明肇事车辆投保商业三者险的事实",
            category="common",
            condition_type="always",
        ))

        # ============================
        # 医疗类证据（5项）
        # ============================
        rules.append(EvidenceRule(
            name="门诊病历",
            source="就诊医院",
            purpose="证明伤者门诊就诊及伤情",
            category="medical",
            condition_field="medical_fee",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="住院病历（含入院记录、出院记录、手术记录、长期医嘱、临时医嘱）",
            source="就诊医院",
            purpose="证明伤者住院治疗经过及伤情",
            category="medical",
            condition_field="hospital_days",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="诊断证明书",
            source="就诊医院",
            purpose="证明伤者伤情诊断及治疗建议",
            category="medical",
            condition_field="medical_fee",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="医疗费发票及费用清单",
            source="就诊医院",
            purpose="证明医疗费支出",
            category="medical",
            condition_field="medical_fee",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="住院费用明细清单",
            source="就诊医院",
            purpose="证明住院期间用药及治疗费用的合理性",
            category="medical",
            condition_field="hospital_days",
            condition_value=0,
            condition_type="gt",
        ))

        # ============================
        # 护理费证据（2项）
        # ============================
        rules.append(EvidenceRule(
            name="护理人身份证复印件及收入证明",
            source="护理人工作单位",
            purpose="证明护理人身份及护理费计算依据",
            category="nursing",
            condition_field="nursing_custom_daily",
            condition_value=None,
            condition_type="not_none",
        ))
        rules.append(EvidenceRule(
            name="护理证明（医院出具需陪护证明）",
            source="就诊医院",
            purpose="证明住院期间需要护理的事实",
            category="nursing",
            condition_field="nursing_days",
            condition_value=0,
            condition_type="gt",
        ))

        # ============================
        # 误工费证据（3项）
        # ============================
        rules.append(EvidenceRule(
            name="误工证明（含停发工资证明）",
            source="原告工作单位",
            purpose="证明原告因伤误工及收入减少的事实",
            category="lost_wage",
            condition_field="lost_work_days",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="劳动合同、社保缴纳记录",
            source="原告工作单位/社保机构",
            purpose="证明原告劳动关系及收入情况",
            category="lost_wage",
            condition_field="lost_work_daily",
            condition_value=None,
            condition_type="not_none",
        ))
        rules.append(EvidenceRule(
            name="事故前12个月工资银行流水",
            source="银行",
            purpose="证明原告实际收入水平，作为误工费计算依据",
            category="lost_wage",
            condition_field="lost_work_daily",
            condition_value=None,
            condition_type="not_none",
        ))

        # ============================
        # 交通费证据（1项）
        # ============================
        rules.append(EvidenceRule(
            name="交通费票据",
            source="交通运输部门",
            purpose="证明原告因就医、转院等产生的交通费用",
            category="traffic",
            condition_field="traffic_fee",
            condition_value=0,
            condition_type="gt",
        ))

        # ============================
        # 伤残鉴定证据（2项）
        # ============================
        rules.append(EvidenceRule(
            name="司法鉴定意见书",
            source="司法鉴定机构",
            purpose="证明伤残等级、三期（误工期、护理期、营养期）评定",
            category="disability",
            condition_field="injury_type",
            condition_value="disability",
            condition_type="eq",
        ))
        rules.append(EvidenceRule(
            name="鉴定费发票",
            source="司法鉴定机构",
            purpose="证明鉴定费支出",
            category="disability",
            condition_field="injury_type",
            condition_value="disability",
            condition_type="eq",
        ))

        # ============================
        # 死亡案件证据（5项）
        # ============================
        rules.append(EvidenceRule(
            name="居民死亡医学证明（推断）书",
            source="医院/公安机关",
            purpose="证明受害人死亡的事实及原因",
            category="death",
            condition_field="injury_type",
            condition_value="death",
            condition_type="eq",
        ))
        rules.append(EvidenceRule(
            name="死亡殡葬证",
            source="民政部门",
            purpose="证明受害人已办理死亡殡葬手续",
            category="death",
            condition_field="injury_type",
            condition_value="death",
            condition_type="eq",
        ))
        rules.append(EvidenceRule(
            name="户口本（全户）复印件",
            source="原告",
            purpose="证明原告与死者的亲属关系，作为赔偿权利人主体资格的依据",
            category="death",
            condition_field="injury_type",
            condition_value="death",
            condition_type="eq",
        ))
        rules.append(EvidenceRule(
            name="法定继承人关系证明",
            source="户籍所在地派出所/村委会",
            purpose="证明原告为死者第一顺序法定继承人",
            category="death",
            condition_field="injury_type",
            condition_value="death",
            condition_type="eq",
        ))
        rules.append(EvidenceRule(
            name="丧葬费票据",
            source="殡仪馆",
            purpose="证明丧葬费实际支出",
            category="death",
            condition_field="injury_type",
            condition_value="death",
            condition_type="eq",
        ))

        # ============================
        # 财产损失证据（2项）
        # ============================
        rules.append(EvidenceRule(
            name="车辆损失定损单",
            source="保险公司/评估机构",
            purpose="证明车辆损失金额",
            category="property",
            condition_field="property_damage",
            condition_value=0,
            condition_type="gt",
        ))
        rules.append(EvidenceRule(
            name="车辆维修发票及维修清单",
            source="维修厂",
            purpose="证明车辆维修费用",
            category="property",
            condition_field="property_damage",
            condition_value=0,
            condition_type="gt",
        ))

        # ============================
        # 被扶养人证据（3项）
        # ============================
        rules.append(EvidenceRule(
            name="被扶养人身份证复印件",
            source="原告",
            purpose="证明被扶养人身份",
            category="dependent",
            condition_field="dependents",
            condition_value=0,
            condition_type="len_gt",
        ))
        rules.append(EvidenceRule(
            name="被扶养人与受害人关系证明",
            source="户籍所在地派出所/村委会",
            purpose="证明被扶养人与受害人的亲属关系",
            category="dependent",
            condition_field="dependents",
            condition_value=0,
            condition_type="len_gt",
        ))
        rules.append(EvidenceRule(
            name="被扶养人无劳动能力证明/在校证明",
            source="社保机构/学校",
            purpose="证明被扶养人无劳动能力或尚在求学期间",
            category="dependent",
            condition_field="dependents",
            condition_value=0,
            condition_type="len_gt",
        ))

        # ============================
        # 刑事附带民事额外证据（1项）
        # ============================
        rules.append(EvidenceRule(
            name="刑事判决书",
            source="人民法院",
            purpose="证明被告已被追究刑事责任，本案系刑事附带民事",
            category="criminal",
            condition_field="case_type",
            condition_value="criminal_attached",
            condition_type="eq",
        ))

        return rules

    def get_all_rules(self) -> list:
        """
        获取全部证据规则（27项）

        Returns:
            list: 证据规则字典列表
        """
        return [rule.to_dict() for rule in self._rules]

    def get_rules_by_category(self, category: str) -> list:
        """
        按类别获取证据规则

        Args:
            category: 证据类别（common/medical/nursing/lost_wage/traffic/
                      disability/death/property/dependent/criminal）

        Returns:
            list: 该类别下的证据规则列表
        """
        return [rule.to_dict() for rule in self._rules
                if rule.category == category]

    def collect_evidence(self, case_data: dict,
                         custom_evidence: list = None) -> list:
        """
        根据案件数据自动推导所需证据清单

        遍历所有规则，根据案件数据判断每条规则是否适用，
        然后追加用户自定义证据。

        Args:
            case_data: 案件数据字典
            custom_evidence: 用户自定义证据列表（可选）

        Returns:
            list: 适用的证据列表，每项包含 name, source, purpose, copies
        """
        evidence = []
        seen = set()

        # 按类别顺序收集
        category_order = [
            'common', 'medical', 'nursing', 'lost_wage', 'traffic',
            'disability', 'death', 'property', 'dependent', 'criminal',
        ]

        for category in category_order:
            for rule in self._rules:
                if rule.category != category:
                    continue
                if rule.name in seen:
                    continue
                if rule.is_applicable(case_data):
                    evidence.append(rule.to_dict())
                    seen.add(rule.name)

        # 追加用户自定义证据
        if custom_evidence:
            for custom in custom_evidence:
                if isinstance(custom, str):
                    evidence.append({
                        'name': custom,
                        'source': '原告',
                        'purpose': '（待补充）',
                        'category': 'custom',
                        'copies': 1,
                    })
                elif isinstance(custom, dict):
                    evidence.append({
                        'name': custom.get('name', ''),
                        'source': custom.get('source', '原告'),
                        'purpose': custom.get('purpose', '（待补充）'),
                        'category': 'custom',
                        'copies': custom.get('copies', 1),
                    })

        return evidence

    def generate_evidence_text(self, case_data: dict,
                               custom_evidence: list = None) -> str:
        """
        生成简版证据清单文本（用于嵌入起诉状）

        格式: "1. xxx；2. xxx；3. xxx。"

        Args:
            case_data: 案件数据字典
            custom_evidence: 用户自定义证据列表

        Returns:
            str: 简版证据清单文本
        """
        evidence = self.collect_evidence(case_data, custom_evidence)
        items = [ev['name'] for ev in evidence]
        return '；'.join([f'{i+1}. {item}' for i, item in enumerate(items)]) + '。'

    def get_category_summary(self) -> dict:
        """
        获取证据类别汇总

        Returns:
            dict: {类别名: 证据数量}
        """
        summary = {}
        for rule in self._rules:
            cat = rule.category
            summary[cat] = summary.get(cat, 0) + 1
        return summary
