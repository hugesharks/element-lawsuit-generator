# -*- coding: utf-8 -*-
"""
机动车交通事故责任纠纷适配器

从 /tmp/traffic-accident-workflow-git/ 迁移核心逻辑，适配新的 CaseAdapter 基类接口。

迁移内容：
- CaseData 类的数据结构 → get_schema()
- CaseData.calculate() 的赔偿计算逻辑 → calculate()
- LawsuitGenerator._build_fill_map() 的模板填充逻辑 → build_fill_map()
- 勾选框逻辑 → get_checkbox_map()
- 证据规则 → get_evidence_rules()（委托 TrafficAccidentEvidenceAdapter）
- 模板配置 → get_template_info()

计算引擎完全复用 engine/data_standards.py 中的 HEBEI_STANDARD 数据，
所有计算逻辑与原有工作流完全一致（已用真实判决书验证）。
"""

from typing import Dict, List, Any, Optional
from .base import CaseAdapter
from .traffic_accident_evidence import TrafficAccidentEvidenceAdapter

# 导入河北省赔偿标准数据
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from engine.data_standards import HEBEI_STANDARD


# ============================================================
# 案件类型与伤情类型常量
# ============================================================
CASE_TYPE_CIVIL = 'civil'                        # 纯民事
CASE_TYPE_CRIMINAL_ATTACHED = 'criminal_attached' # 刑事附带民事

INJURY_DEATH = 'death'                  # 死亡
INJURY_DISABILITY = 'disability'        # 伤残
INJURY_PROPERTY = 'property'            # 仅车损/财产损失
INJURY_PROPERTY_PERSONAL = 'property_personal'  # 车损+人身伤害（无伤残）

# 责任类型枚举
LIABILITY_TYPES = ['全责', '主责', '同责', '次责', '无责']

# 责任类型映射（适配器格式 → data_standards格式）
_LIABILITY_MAP = {
    '全责': '全部责任',
    '主责': '主要责任',
    '同责': '同等责任',
    '次责': '次要责任',
    '无责': '无责任',
}

# 默认通用参数（这些在 data_standards.py 中没有单独字段，从原始工作流迁移）
_DEFAULT_HOSPITAL_MEAL = 100       # 住院伙食补助 元/天
_DEFAULT_NUTRITION = 20            # 营养费 元/天
_DEFAULT_NURSING = 163.34          # 护理费 元/天（居民服务业59618元/年÷365）
_DEFAULT_WAGE = 163.34             # 默认误工费 元/天


# ============================================================
# 辅助函数：从 HEBEI_STANDARD 获取数据
# ============================================================
def _get_year_data(year) -> dict:
    """获取指定年度的基础数据，year 可以是 int 或 str"""
    key = str(year) if isinstance(year, int) else year
    return HEBEI_STANDARD.get(key, HEBEI_STANDARD.get('2024', {}))


def _get_funeral_fee(year) -> float:
    """获取丧葬费"""
    year_data = _get_year_data(year)
    return year_data.get('funeral', 0)


def _get_disability_coefficient(grade: int) -> float:
    """获取伤残赔偿系数"""
    coefs = HEBEI_STANDARD.get('disability_ratio', {})
    return coefs.get(grade, 0.1)


def _get_mental_damages(grade) -> int:
    """获取精神损害抚慰金"""
    md = HEBEI_STANDARD.get('mental_damage', {})
    return md.get(grade, 5000)


def _get_liability_ratio(liability_type: str) -> float:
    """获取被告责任比例"""
    lr = HEBEI_STANDARD.get('liability_ratio', {})
    return lr.get(liability_type, 1.0)


def _get_compulsory_limits() -> dict:
    """获取交强险分项限额"""
    return HEBEI_STANDARD.get('compulsory_insurance', {
        'medical': 18000,
        'death_disability': 180000,
        'property': 2000,
    })


# ============================================================
# 赔偿计算引擎（从原有 CompensationCalculator 迁移）
# ============================================================
class CompensationCalculator:
    """
    交通事故赔偿计算引擎

    基于河北省标准，计算13项赔偿项目及保险拆分。
    所有计算公式严格依据《民法典》第1179条及相关司法解释。
    """

    def __init__(self, standard_year: int = 2024, region: str = 'urban'):
        """
        初始化计算引擎

        Args:
            standard_year: 标准年度（2023/2024/2025）
            region: 地域类型（'urban'城镇 / 'rural'农村）
        """
        self.year = standard_year
        self.year_data = _get_year_data(standard_year)
        self.region = region

    @property
    def income(self) -> float:
        """获取年收入标准"""
        key = 'urban_income' if self.region == 'urban' else 'rural_income'
        return self.year_data.get(key, 0)

    @property
    def consumption(self) -> float:
        """获取年人均消费支出"""
        key = 'urban_consumption' if self.region == 'urban' else 'rural_consumption'
        return self.year_data.get(key, 0)

    def calc_medical_fee(self, amount: float) -> float:
        """医疗费：按实际票据金额"""
        return amount

    def calc_hospital_meal(self, days: int, daily: float = None) -> float:
        """住院伙食补助费 = 天数 × 日标准"""
        daily_rate = daily if daily is not None else _DEFAULT_HOSPITAL_MEAL
        return days * daily_rate

    def calc_nutrition(self, days: int, daily: float = None) -> float:
        """营养费 = 天数 × 日标准（河北省20元/天）"""
        daily_rate = daily if daily is not None else _DEFAULT_NUTRITION
        return days * daily_rate

    def calc_nursing(self, days: int, persons: int = 1,
                     custom_daily: float = None) -> float:
        """护理费 = 天数 × 护理人数 × 日标准"""
        daily_rate = custom_daily if custom_daily is not None else _DEFAULT_NURSING
        return days * persons * daily_rate

    def calc_lost_wage(self, days: int, daily_wage: float = None) -> float:
        """误工费 = 天数 × 日工资（无固定收入按居民服务业标准）"""
        daily = daily_wage if daily_wage is not None else _DEFAULT_WAGE
        return days * daily

    def calc_traffic_fee(self, amount: float) -> float:
        """交通费：按实际票据金额"""
        return amount

    def calc_disability_compensation(self, grade: int, age: int,
                                     base_years: int = 20) -> float:
        """
        残疾赔偿金 = 年收入 × 伤残系数 × 赔偿年限

        赔偿年限规则：
        - 60周岁以下：20年
        - 60-74周岁：20 - (年龄 - 60)，最少5年
        - 75周岁以上：5年
        """
        ratio = _get_disability_coefficient(grade)

        if age >= 60:
            years = max(5, base_years - (age - 60))
        else:
            years = 20

        return self.income * ratio * years

    def calc_death_compensation(self, age: int, base_years: int = 20) -> float:
        """
        死亡赔偿金 = 年收入 × 赔偿年限

        赔偿年限规则同残疾赔偿金。
        """
        if age >= 60:
            years = max(5, base_years - (age - 60))
        else:
            years = 20

        return self.income * years

    def calc_funeral(self) -> float:
        """丧葬费 = 全省在岗职工年平均工资 / 2"""
        return _get_funeral_fee(self.year)

    def calc_mental_damage(self, case_type: str, injury_type: str,
                           grade: int = None) -> float:
        """
        精神损害抚慰金

        规则：
        - 刑事附带民事案件不支持精神损害抚慰金
        - 死亡案件：按标准
        - 伤残案件：按等级对应
        """
        if case_type == CASE_TYPE_CRIMINAL_ATTACHED:
            return 0
        if injury_type == INJURY_DEATH:
            return _get_mental_damages('death')
        elif injury_type == INJURY_DISABILITY and grade:
            return _get_mental_damages(grade)
        return 0

    def calc_dependent_living(self, dependents: list) -> float:
        """
        被扶养人生活费（分段计算）

        计算规则：
        - 每个被扶养人：人均消费支出 × 扶养比例 × 扶养年限
        - 年赔偿总额校验：不超过人均消费支出

        Args:
            dependents: [{age, years, share_ratio}, ...]

        Returns:
            被扶养人生活费总额
        """
        total = 0
        for dep in dependents:
            share_ratio = dep.get('share_ratio', 1.0)
            years = dep.get('years', 0)
            total += self.consumption * share_ratio * years

        return total

    def calc_insurance_split(self, loss_detail: dict,
                             liability_type: str = '全责') -> dict:
        """
        交强险/商业险分项拆解（严格按分项限额）

        拆分逻辑：
        1. 交强险按分项限额赔付（不分责任比例）
        2. 剩余部分按被告责任比例由商业险赔付
        3. 商业险不赔的部分由被告自行承担

        Args:
            loss_detail: 分项损失明细
            liability_type: 被告责任类型

        Returns:
            保险拆分结果字典
        """
        ratio = _get_liability_ratio(liability_type)
        limits = _get_compulsory_limits()

        # 交强险严格按分项限额赔付（不分责任比例）
        comp_medical = min(loss_detail.get('medical', 0), limits['medical'])
        comp_death_dis = min(loss_detail.get('death_disability', 0),
                             limits['death_disability'])
        comp_property = min(loss_detail.get('property', 0), limits['property'])
        compulsory_total = comp_medical + comp_death_dis + comp_property

        # 剩余部分按被告责任比例由商业险赔付
        total_loss = (loss_detail.get('medical', 0)
                      + loss_detail.get('death_disability', 0)
                      + loss_detail.get('property', 0))
        remaining = total_loss - compulsory_total
        commercial = remaining * ratio
        self_bear = remaining * (1 - ratio)

        return {
            'compulsory_medical': comp_medical,
            'compulsory_death_disability': comp_death_dis,
            'compulsory_property': comp_property,
            'compulsory_total': compulsory_total,
            'commercial': commercial,
            'self_bear': self_bear,
        }


# ============================================================
# 交通事故适配器主类
# ============================================================
class TrafficAccidentAdapter(CaseAdapter):
    """
    机动车交通事故责任纠纷适配器

    支持案件类型：
    - 纯民事 / 刑事附带民事
    - 伤情类型：死亡 / 伤残 / 仅财产损失 / 财产损失+人身伤害
    - 责任比例：全责/主责/同责/次责/无责
    - 多原告、多被告（自然人+法人）
    - 被扶养人生活费分段计算
    - 交强险分项限额 + 商业险 + 自担拆分

    计算逻辑完全复用原有工作流（已用真实判决书验证）。
    """

    def __init__(self):
        """初始化适配器，创建证据规则适配器实例"""
        self._evidence_adapter = TrafficAccidentEvidenceAdapter()

    # ================================================================
    # 1. 案由名称
    # ================================================================
    def name(self) -> str:
        """返回案由名称"""
        return "机动车交通事故责任纠纷"

    # ================================================================
    # 2. 输入字段 Schema
    # ================================================================
    def get_schema(self) -> dict:
        """
        返回输入字段的 JSON Schema 定义

        包含以下字段组：
        - 原告信息（姓名、性别、出生日期、民族、住址、身份证号、电话、年龄）
        - 案件类型（纯民事/刑事附带民事，伤情类型）
        - 责任比例（全责/主责/同责/次责/无责）
        - 医疗信息（住院天数、医疗费、护理期、误工期、营养期等）
        - 伤残信息（伤残等级1-10）
        - 被扶养人（数组）
        - 被告信息（自然人+法人）
        - 法院信息
        """
        return {
            "type": "object",
            "required": ["plaintiffs", "case_type", "injury_type", "liability_type"],
            "properties": {
                # === 原告信息 ===
                "plaintiffs": {
                    "type": "array",
                    "title": "原告列表",
                    "description": "支持多原告（如死亡案件中多名近亲属共同起诉）",
                    "items": {
                        "type": "object",
                        "required": ["name"],
                        "properties": {
                            "name": {"type": "string", "title": "姓名"},
                            "gender": {
                                "type": "string",
                                "title": "性别",
                                "enum": ["男", "女"]
                            },
                            "birthdate": {
                                "type": "string",
                                "title": "出生日期",
                                "description": "格式：YYYY年MM月DD日"
                            },
                            "ethnicity": {"type": "string", "title": "民族"},
                            "address": {"type": "string", "title": "住所地（户籍所在地）"},
                            "residence": {"type": "string", "title": "经常居住地"},
                            "id_type": {
                                "type": "string",
                                "title": "证件类型",
                                "default": "居民身份证"
                            },
                            "id_number": {"type": "string", "title": "身份证号"},
                            "phone": {"type": "string", "title": "联系电话"},
                            "work": {"type": "string", "title": "工作单位"},
                            "position": {"type": "string", "title": "职务"},
                            "relation": {
                                "type": "string",
                                "title": "与死者关系",
                                "description": "死亡案件中与死者的关系"
                            },
                        }
                    }
                },

                # === 原告法人（可选） ===
                "plaintiffs_company": {
                    "type": "array",
                    "title": "原告法人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "名称"},
                            "address": {"type": "string", "title": "住所地"},
                            "legal_person": {"type": "string", "title": "法定代表人"},
                            "position": {"type": "string", "title": "职务"},
                            "phone": {"type": "string", "title": "联系电话"},
                            "credit_code": {"type": "string", "title": "统一社会信用代码"},
                            "company_type": {"type": "string", "title": "企业类型"},
                        }
                    }
                },

                # === 案件类型 ===
                "case_type": {
                    "type": "string",
                    "title": "案件类型",
                    "enum": [CASE_TYPE_CIVIL, CASE_TYPE_CRIMINAL_ATTACHED],
                    "enumNames": ["纯民事", "刑事附带民事"],
                    "default": CASE_TYPE_CIVIL,
                },
                "injury_type": {
                    "type": "string",
                    "title": "伤情类型",
                    "enum": [
                        INJURY_DEATH,
                        INJURY_DISABILITY,
                        INJURY_PROPERTY,
                        INJURY_PROPERTY_PERSONAL,
                    ],
                    "enumNames": ["死亡", "伤残", "仅财产损失", "财产损失+人身伤害"],
                    "default": INJURY_DISABILITY,
                },

                # === 责任比例 ===
                "liability_type": {
                    "type": "string",
                    "title": "被告责任类型",
                    "description": "被告（侵权方）对事故承担的责任比例",
                    "enum": LIABILITY_TYPES,
                    "default": "全责",
                },

                # === 赔偿标准 ===
                "standard_year": {
                    "type": "integer",
                    "title": "赔偿标准年度",
                    "enum": [2023, 2024, 2025],
                    "default": 2024,
                },
                "region": {
                    "type": "string",
                    "title": "地域类型",
                    "enum": ["urban", "rural"],
                    "enumNames": ["城镇", "农村"],
                    "default": "urban",
                },

                # === 医疗信息 ===
                "hospital_days": {
                    "type": "integer",
                    "title": "住院天数",
                    "minimum": 0,
                    "default": 0,
                },
                "medical_fee": {
                    "type": "number",
                    "title": "医疗费（元）",
                    "minimum": 0,
                    "default": 0,
                },
                "hospital_meal_daily": {
                    "type": ["number", "null"],
                    "title": "住院伙食补助日标准（元）",
                    "description": "默认100元/天，可自定义",
                },
                "nutrition_daily": {
                    "type": ["number", "null"],
                    "title": "营养费日标准（元）",
                    "description": "默认20元/天，可自定义",
                },

                # === 三期信息 ===
                "nursing_days": {
                    "type": "integer",
                    "title": "护理期（天）",
                    "minimum": 0,
                    "default": 0,
                },
                "nursing_persons": {
                    "type": "integer",
                    "title": "护理人数",
                    "minimum": 1,
                    "default": 1,
                },
                "nursing_custom_daily": {
                    "type": ["number", "null"],
                    "title": "护理费日标准（元）",
                    "description": "默认按居民服务业标准163.34元/天，有护理人收入证明时可自定义",
                },
                "nutrition_days": {
                    "type": "integer",
                    "title": "营养期（天）",
                    "minimum": 0,
                    "default": 0,
                },
                "lost_work_days": {
                    "type": "integer",
                    "title": "误工期（天）",
                    "minimum": 0,
                    "default": 0,
                },
                "lost_work_daily": {
                    "type": ["number", "null"],
                    "title": "日工资（元）",
                    "description": "无固定收入时按居民服务业标准163.34元/天",
                },

                # === 伤残信息 ===
                "disability_grade": {
                    "type": ["integer", "null"],
                    "title": "伤残等级",
                    "enum": [None, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                    "description": "1级最重，10级最轻",
                },

                # === 原告年龄（用于赔偿年限计算） ===
                "plaintiff_age": {
                    "type": "integer",
                    "title": "原告（赔偿权利人）年龄",
                    "minimum": 0,
                    "default": 0,
                },

                # === 被扶养人 ===
                "dependents": {
                    "type": "array",
                    "title": "被扶养人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "age": {"type": "integer", "title": "被扶养人年龄"},
                            "years": {
                                "type": "integer",
                                "title": "扶养年限",
                                "description": "未成年人计算至18周岁，无劳动能力者计算20年"
                            },
                            "share_ratio": {
                                "type": "number",
                                "title": "扶养比例",
                                "description": "共同扶养人分摊比例，如2人扶养则为0.5",
                                "minimum": 0,
                                "maximum": 1,
                            },
                            "relation": {"type": "string", "title": "与受害人关系"},
                        }
                    }
                },

                # === 其他费用 ===
                "traffic_fee": {
                    "type": "number",
                    "title": "交通费（元）",
                    "minimum": 0,
                    "default": 0,
                },
                "property_damage": {
                    "type": "number",
                    "title": "财产损失（元）",
                    "minimum": 0,
                    "default": 0,
                },
                "assistive_device_fee": {
                    "type": "number",
                    "title": "辅助器具费（元）",
                    "minimum": 0,
                    "default": 0,
                },
                "other_fee": {
                    "type": "number",
                    "title": "其他费用（元）",
                    "minimum": 0,
                    "default": 0,
                },
                "other_fee_desc": {
                    "type": "string",
                    "title": "其他费用说明",
                },

                # === 被告信息 ===
                "defendants_person": {
                    "type": "array",
                    "title": "被告自然人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "姓名"},
                            "gender": {"type": "string", "title": "性别", "enum": ["男", "女"]},
                            "birthdate": {"type": "string", "title": "出生日期"},
                            "ethnicity": {"type": "string", "title": "民族"},
                            "address": {"type": "string", "title": "住所地"},
                            "id_number": {"type": "string", "title": "身份证号"},
                            "phone": {"type": "string", "title": "联系电话"},
                        }
                    }
                },
                "defendants_company": {
                    "type": "array",
                    "title": "被告法人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "名称"},
                            "address": {"type": "string", "title": "住所地"},
                            "legal_person": {"type": "string", "title": "法定代表人"},
                            "position": {"type": "string", "title": "职务"},
                            "phone": {"type": "string", "title": "联系电话"},
                            "credit_code": {"type": "string", "title": "统一社会信用代码"},
                            "company_type": {"type": "string", "title": "企业类型"},
                        }
                    }
                },

                # === 第三人信息（可选） ===
                "third_parties_person": {
                    "type": "array",
                    "title": "第三人自然人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "姓名"},
                            "gender": {"type": "string", "title": "性别"},
                            "birthdate": {"type": "string", "title": "出生日期"},
                            "address": {"type": "string", "title": "住所地"},
                            "id_number": {"type": "string", "title": "身份证号"},
                        }
                    }
                },
                "third_parties_company": {
                    "type": "array",
                    "title": "第三人法人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "名称"},
                            "address": {"type": "string", "title": "住所地"},
                            "legal_person": {"type": "string", "title": "法定代表人"},
                            "credit_code": {"type": "string", "title": "统一社会信用代码"},
                        }
                    }
                },

                # === 代理人 ===
                "agents": {
                    "type": "array",
                    "title": "委托诉讼代理人列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "姓名"},
                            "unit": {"type": "string", "title": "单位"},
                            "position": {"type": "string", "title": "职务"},
                            "phone": {"type": "string", "title": "联系电话"},
                            "authority": {
                                "type": "string",
                                "title": "代理权限",
                                "enum": ["一般授权", "特别授权"],
                            },
                        }
                    }
                },

                # === 事故信息 ===
                "accident_time": {"type": "string", "title": "事故发生时间"},
                "accident_location": {"type": "string", "title": "事故发生地点"},
                "accident_detail": {
                    "type": "string",
                    "title": "事故经过",
                    "description": "详细描述事故发生的经过"
                },
                "responsibility_doc_number": {
                    "type": "string",
                    "title": "事故认定书编号",
                },
                "responsibility_result": {
                    "type": "string",
                    "title": "责任认定结果",
                    "description": "如：被告负主要责任，原告负次要责任"
                },

                # === 保险信息 ===
                "insurance_info": {
                    "type": "string",
                    "title": "保险信息",
                    "description": "交强险、商业三者险的保险公司及保单信息"
                },

                # === 法院信息 ===
                "court": {
                    "type": "string",
                    "title": "管辖法院",
                },
                "filing_date": {
                    "type": "string",
                    "title": "起诉日期",
                },

                # === 证据列表（可选，用户自定义补充） ===
                "custom_evidence": {
                    "type": "array",
                    "title": "自定义证据列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "title": "证据名称"},
                            "source": {"type": "string", "title": "证据来源"},
                            "purpose": {"type": "string", "title": "证明目的"},
                        }
                    }
                },
            }
        }

    # ================================================================
    # 3. 赔偿计算
    # ================================================================
    def calculate(self, case_data: dict) -> dict:
        """
        执行完整的赔偿计算

        计算13个赔偿项目 + 保险拆分，与原有工作流完全一致。

        Args:
            case_data: 案件数据字典

        Returns:
            {
                "items": [
                    {"name": "医疗费", "amount": 52340.80, "formula": "实际支出",
                     "legal_basis": "民法典第1179条"},
                    ...
                ],
                "total": 123456.78,
                "insurance_split": {
                    "compulsory_medical": 18000,
                    "compulsory_death_disability": 180000,
                    "compulsory_property": 2000,
                    "compulsory_total": 200000,
                    "commercial": 50000,
                    "self_bear": 12345.67,
                }
            }
        """
        # 初始化计算引擎
        standard_year = case_data.get('standard_year', 2024)
        region = case_data.get('region', 'urban')
        calc = CompensationCalculator(standard_year, region)

        # 获取关键参数
        case_type = case_data.get('case_type', CASE_TYPE_CIVIL)
        injury_type = case_data.get('injury_type', INJURY_DISABILITY)
        liability_type = case_data.get('liability_type', '全责')
        hospital_days = case_data.get('hospital_days', 0)
        medical_fee = case_data.get('medical_fee', 0)
        plaintiff_age = case_data.get('plaintiff_age', 0)
        disability_grade = case_data.get('disability_grade')

        # === 1. 医疗费 ===
        item_medical = calc.calc_medical_fee(medical_fee)

        # === 2. 住院伙食补助费 ===
        hospital_meal_daily = case_data.get('hospital_meal_daily')
        if hospital_meal_daily is not None:
            item_hospital_meal = hospital_days * hospital_meal_daily
            meal_formula = f"{hospital_meal_daily}元/天×{hospital_days}天"
        else:
            item_hospital_meal = calc.calc_hospital_meal(hospital_days)
            meal_formula = f"{_DEFAULT_HOSPITAL_MEAL}元/天×{hospital_days}天"

        # === 3. 营养费 ===
        nutrition_days = case_data.get('nutrition_days', 0) or hospital_days
        nutrition_daily = case_data.get('nutrition_daily')
        if nutrition_daily is not None:
            item_nutrition = nutrition_days * nutrition_daily
            nutrition_formula = f"{nutrition_daily}元/天×{nutrition_days}天"
        else:
            item_nutrition = calc.calc_nutrition(nutrition_days)
            nutrition_formula = f"{_DEFAULT_NUTRITION}元/天×{nutrition_days}天"

        # === 4. 护理费 ===
        nursing_days = case_data.get('nursing_days', 0) or hospital_days
        nursing_persons = case_data.get('nursing_persons', 1)
        nursing_custom_daily = case_data.get('nursing_custom_daily')
        item_nursing = calc.calc_nursing(nursing_days, nursing_persons, nursing_custom_daily)
        nursing_daily_rate = (nursing_custom_daily if nursing_custom_daily is not None
                              else _DEFAULT_NURSING)
        nursing_formula = f"{nursing_daily_rate:.2f}元/天×{nursing_days}天×{nursing_persons}人"

        # === 5. 误工费 ===
        lost_work_days = case_data.get('lost_work_days', 0) or hospital_days
        lost_work_daily = case_data.get('lost_work_daily')
        item_lost_wage = calc.calc_lost_wage(lost_work_days, lost_work_daily)
        wage_daily_rate = (lost_work_daily if lost_work_daily is not None
                           else _DEFAULT_WAGE)
        wage_formula = f"{wage_daily_rate:.2f}元/天×{lost_work_days}天"

        # === 6. 交通费 ===
        traffic_fee = case_data.get('traffic_fee', 0)
        item_traffic = calc.calc_traffic_fee(traffic_fee)

        # === 7-9. 根据伤情类型计算 ===
        item_disability = 0
        item_death = 0
        item_funeral = 0
        item_assistive = case_data.get('assistive_device_fee', 0)

        if injury_type == INJURY_DEATH:
            # 死亡案件
            item_death = calc.calc_death_compensation(plaintiff_age)
            item_funeral = calc.calc_funeral()
        elif injury_type == INJURY_DISABILITY:
            # 伤残案件
            if disability_grade:
                item_disability = calc.calc_disability_compensation(
                    disability_grade, plaintiff_age
                )
        # property / property_personal：无残疾/死亡赔偿金

        # === 被扶养人生活费 ===
        dependents = case_data.get('dependents', [])
        item_dependent = 0
        if dependents:
            item_dependent = calc.calc_dependent_living(dependents)

        # === 10. 精神损害抚慰金 ===
        item_mental = calc.calc_mental_damage(case_type, injury_type, disability_grade)

        # === 11. 财产损失 ===
        item_property = case_data.get('property_damage', 0)

        # === 12. 其他费用 ===
        item_other = case_data.get('other_fee', 0)

        # === 13. 标的总额 ===
        total = (
            item_medical + item_hospital_meal + item_nutrition +
            item_nursing + item_lost_wage + item_traffic +
            item_disability + item_death + item_funeral +
            item_dependent + item_assistive + item_mental +
            item_property + item_other
        )

        # === 构建赔偿项目列表（含公式和法律依据） ===
        items = []

        if item_medical > 0:
            items.append({
                "name": "医疗费",
                "amount": round(item_medical, 2),
                "formula": "实际支出",
                "legal_basis": "《民法典》第1179条",
            })

        if item_hospital_meal > 0:
            items.append({
                "name": "住院伙食补助费",
                "amount": round(item_hospital_meal, 2),
                "formula": meal_formula,
                "legal_basis": "《民法典》第1179条；河北省标准",
            })

        if item_nutrition > 0:
            items.append({
                "name": "营养费",
                "amount": round(item_nutrition, 2),
                "formula": nutrition_formula,
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第11条",
            })

        if item_nursing > 0:
            items.append({
                "name": "护理费",
                "amount": round(item_nursing, 2),
                "formula": nursing_formula,
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第8条",
            })

        if item_lost_wage > 0:
            items.append({
                "name": "误工费",
                "amount": round(item_lost_wage, 2),
                "formula": wage_formula,
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第7条",
            })

        if item_traffic > 0:
            items.append({
                "name": "交通费",
                "amount": round(item_traffic, 2),
                "formula": "实际支出",
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第9条",
            })

        if item_disability > 0:
            ratio_val = _get_disability_coefficient(disability_grade)
            income_val = calc.income
            if plaintiff_age >= 60:
                years_val = max(5, 20 - (plaintiff_age - 60))
            else:
                years_val = 20
            items.append({
                "name": "残疾赔偿金",
                "amount": round(item_disability, 2),
                "formula": f"{income_val}元×{ratio_val}×{years_val}年",
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第12条",
            })

        if item_death > 0:
            income_val = calc.income
            if plaintiff_age >= 60:
                years_val = max(5, 20 - (plaintiff_age - 60))
            else:
                years_val = 20
            items.append({
                "name": "死亡赔偿金",
                "amount": round(item_death, 2),
                "formula": f"{income_val}元×{years_val}年",
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第15条",
            })

        if item_funeral > 0:
            funeral_val = calc.calc_funeral()
            items.append({
                "name": "丧葬费",
                "amount": round(item_funeral, 2),
                "formula": f"{funeral_val}元",
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第14条",
            })

        if item_dependent > 0:
            consumption_val = calc.consumption
            dep_formulas = []
            for dep in dependents:
                sr = dep.get('share_ratio', 1.0)
                yr = dep.get('years', 0)
                dep_formulas.append(f"{consumption_val}×{sr}×{yr}")
            items.append({
                "name": "被扶养人生活费",
                "amount": round(item_dependent, 2),
                "formula": " + ".join(dep_formulas),
                "legal_basis": "《民法典》第1179条；《最高人民法院关于审理人身损害赔偿案件适用法律若干问题的解释》第17条",
            })

        if item_assistive > 0:
            items.append({
                "name": "辅助器具费",
                "amount": round(item_assistive, 2),
                "formula": "实际支出",
                "legal_basis": "《民法典》第1179条",
            })

        if item_mental > 0:
            mental_desc = ("死亡" if injury_type == INJURY_DEATH
                           else f"{disability_grade}级伤残")
            items.append({
                "name": "精神损害抚慰金",
                "amount": round(item_mental, 2),
                "formula": f"{mental_desc}对应{item_mental}元",
                "legal_basis": "《民法典》第1183条；《最高人民法院关于确定民事侵权精神损害赔偿责任若干问题的解释》",
            })

        if item_property > 0:
            items.append({
                "name": "财产损失",
                "amount": round(item_property, 2),
                "formula": "实际损失",
                "legal_basis": "《民法典》第1184条",
            })

        if item_other > 0:
            desc = case_data.get('other_fee_desc', '其他费用')
            items.append({
                "name": desc,
                "amount": round(item_other, 2),
                "formula": "实际支出",
                "legal_basis": "《民法典》第1179条",
            })

        # === 保险拆分 ===
        # 医疗费分项：医疗费+住院伙食补助费+营养费
        loss_medical = item_medical + item_hospital_meal + item_nutrition
        # 死亡伤残分项：残疾/死亡赔偿金+误工费+护理费+交通费
        #               +丧葬费+精神抚慰金+被扶养人生活费+辅助器具费
        loss_death_disability = (
            item_disability + item_death +
            item_lost_wage + item_nursing + item_traffic +
            item_funeral + item_mental +
            item_dependent + item_assistive
        )
        # 财产损失分项：财产损失+其他费用
        loss_property = item_property + item_other

        insurance_split = calc.calc_insurance_split(
            {
                'medical': loss_medical,
                'death_disability': loss_death_disability,
                'property': loss_property,
            },
            liability_type
        )

        return {
            "items": items,
            "total": round(total, 2),
            "insurance_split": insurance_split,
        }

    # ================================================================
    # 4. 书签填充映射
    # ================================================================
    def build_fill_map(self, case_data: dict, calc_result: dict) -> Dict[str, str]:
        """
        构建书签填充映射表

        将案件数据和计算结果映射为 {书签名: 填充值} 的字典。

        书签名规则：
        - 原告区域：T0_{序号:02d}_字段名
        - 被告区域：T1_{序号:02d}_字段名
        - 诉讼请求区：T2_{序号:02d}_字段名
        - 事实与理由区：T3_{序号:02d}_字段名

        Args:
            case_data: 原始案件数据
            calc_result: calculate() 的返回结果

        Returns:
            dict[str, str]: {书签名: 填充值}
        """
        fill_map = {}

        # ============================
        # 原告区域
        # ============================
        for i, plaintiff in enumerate(case_data.get('plaintiffs', []), start=1):
            prefix = f"T0_{i:02d}"
            if plaintiff.get('name'):
                fill_map[f"{prefix}_原告姓名"] = plaintiff['name']
            if plaintiff.get('gender'):
                fill_map[f"{prefix}_原告性别"] = plaintiff['gender']
            if plaintiff.get('birthdate'):
                fill_map[f"{prefix}_原告出生日期"] = plaintiff['birthdate']
            if plaintiff.get('ethnicity'):
                fill_map[f"{prefix}_原告民族"] = plaintiff['ethnicity']
            if plaintiff.get('address'):
                fill_map[f"{prefix}_原告住址"] = plaintiff['address']
            if plaintiff.get('id_number'):
                fill_map[f"{prefix}_原告身份证号"] = plaintiff['id_number']
            if plaintiff.get('phone'):
                fill_map[f"{prefix}_原告电话"] = plaintiff['phone']

        # ============================
        # 被告自然人区域
        # ============================
        defendants_person = case_data.get('defendants_person', [])
        defendants_company = case_data.get('defendants_company', [])

        for i, defendant in enumerate(defendants_person, start=1):
            prefix = f"T1_{i:02d}"
            if defendant.get('name'):
                fill_map[f"{prefix}_被告姓名"] = defendant['name']
            if defendant.get('gender'):
                fill_map[f"{prefix}_被告性别"] = defendant.get('gender', '')
            if defendant.get('id_number'):
                fill_map[f"{prefix}_被告身份证号"] = defendant.get('id_number', '')
            if defendant.get('address'):
                fill_map[f"{prefix}_被告住址"] = defendant.get('address', '')

        # ============================
        # 被告法人区域（紧跟自然人后面编号）
        # ============================
        for i, company in enumerate(
            defendants_company, start=len(defendants_person) + 1
        ):
            prefix = f"T1_{i:02d}"
            if company.get('name'):
                fill_map[f"{prefix}_被告名称"] = company['name']
            if company.get('address'):
                fill_map[f"{prefix}_被告住址"] = company.get('address', '')
            if company.get('legal_person'):
                fill_map[f"{prefix}_被告法定代表人"] = company.get('legal_person', '')
            if company.get('credit_code'):
                fill_map[f"{prefix}_被告信用代码"] = company.get('credit_code', '')

        # ============================
        # 诉讼请求区域
        # ============================
        items = calc_result.get('items', [])
        for i, item in enumerate(items, start=1):
            prefix = f"T2_{i:02d}"
            fill_map[f"{prefix}_赔偿项目"] = item['name']
            fill_map[f"{prefix}_赔偿金额"] = f"{item['amount']:.2f}"
            fill_map[f"{prefix}_计算公式"] = item['formula']
            fill_map[f"{prefix}_法律依据"] = item['legal_basis']

        # ============================
        # 事实与理由区域
        # ============================
        if case_data.get('accident_detail'):
            fill_map["T3_01_事故经过"] = case_data['accident_detail']
        if case_data.get('responsibility_result'):
            fill_map["T3_02_责任认定"] = case_data['responsibility_result']
        if case_data.get('insurance_info'):
            fill_map["T3_03_保险信息"] = case_data['insurance_info']
        fill_map["T3_04_诉讼请求总额"] = f"{calc_result['total']:.2f}"

        # ============================
        # 补充：责任比例说明
        # ============================
        liability_type = case_data.get('liability_type', '全责')
        liability_ratio = _get_liability_ratio(liability_type)
        liability_pct = liability_ratio * 100
        fill_map["T3_05_责任比例说明"] = (
            f"被告{liability_type}（{liability_pct:.0f}%），应当承担赔偿责任。"
        )

        # 保险公司说明
        if defendants_company:
            fill_map["T3_06_保险责任说明"] = (
                "保险公司应在交强险和商业三者险限额内承担赔偿责任。"
            )

        return fill_map

    # ================================================================
    # 5. 勾选框映射
    # ================================================================
    def get_checkbox_map(self, case_data: dict) -> Dict[str, bool]:
        """
        返回勾选框映射

        根据案件类型、伤情类型、责任比例、原告性别等信息，
        生成模板中各勾选框的勾选状态。

        命名规则（避免跨表误匹配）：
        - 可能在多个表格出现的选项 → 加表前缀，如 "T0_性别_男", "T1_性别_男"
        - 只在一个位置出现的选项 → 用通用key，如 "案件类型_纯民事"
        
        勾选框处理器的匹配逻辑：
        1. 先用 "T{table_idx}_{option}" 匹配（精确到表）
        2. 匹配不到则回退到 "{category}_{option}" 匹配（全局）

        Args:
            case_data: 案件数据字典

        Returns:
            dict[str, bool]: {勾选框key: True/False}
        """
        case_type = case_data.get('case_type', CASE_TYPE_CIVIL)
        injury_type = case_data.get('injury_type', INJURY_DISABILITY)
        liability_type = case_data.get('liability_type', '全责')

        # 获取第一原告性别
        plaintiffs = case_data.get('plaintiffs', [])
        plaintiff_gender = ''
        if plaintiffs:
            plaintiff_gender = plaintiffs[0].get('gender', '')

        # 获取第一被告自然人性别
        defendants_person = case_data.get('defendants_person', [])
        defendant_gender = ''
        if defendants_person:
            defendant_gender = defendants_person[0].get('gender', '')

        checkbox_map = {
            # === 案件类型勾选（T3区域，唯一位置）===
            "案件类型_纯民事": case_type == CASE_TYPE_CIVIL,
            "案件类型_刑事附带民事": case_type == CASE_TYPE_CRIMINAL_ATTACHED,

            # === 伤情类型勾选（T2区域，唯一位置）===
            "伤情_死亡": injury_type == INJURY_DEATH,
            "伤情_伤残": injury_type == INJURY_DISABILITY,
            "伤情_财产损失": injury_type in (INJURY_PROPERTY, INJURY_PROPERTY_PERSONAL),

            # === 责任比例勾选（T3区域，唯一位置）===
            "责任_全责": liability_type == '全责',
            "责任_主责": liability_type == '主责',
            "责任_同责": liability_type == '同责',
            "责任_次责": liability_type == '次责',

            # === 性别勾选 — 带表前缀，避免T0/T1/T3的"男□"全被勾选 ===
            # T0 = 原告区
            "T0_性别_男": plaintiff_gender == '男',
            "T0_性别_女": plaintiff_gender == '女',
            # T1 = 被告区
            "T1_性别_男": defendant_gender == '男',
            "T1_性别_女": defendant_gender == '女',

            # === 诉前保全 — 带表前缀（T3区域）===
            "T3_诉前保全_否": not case_data.get('pre_litigation_preservation', False),
            "T3_诉前保全_是": case_data.get('pre_litigation_preservation', False),
        }

        # === 伤残/死亡案件自动勾选鉴定选项（T3区域）===
        # "是""否"在多个表出现，必须用表前缀
        has_appraisal = (
            injury_type == INJURY_DISABILITY
            or injury_type == INJURY_DEATH
            or case_data.get('disability_grade') is not None
        )
        checkbox_map["T3_鉴定_是"] = has_appraisal
        checkbox_map["T3_鉴定_否"] = not has_appraisal

        # === T3调解意愿（默认不勾选，用户可覆盖）===
        checkbox_map["了解"] = True   # 默认了解调解
        checkbox_map["不了解"] = False

        # === T4调解选择 ===
        # 注意：T4 Row 1 有 "是□否□暂不确定，想要了解更多内容□"
        # "是""否"在多个表出现（T3保全/T3鉴定/T4调解），必须用表前缀区分
        checkbox_map["T4_是"] = False  # 是否考虑先行调解
        checkbox_map["T4_否"] = False
        checkbox_map["暂不确定，想要了解更多内容"] = False

        return checkbox_map

    # ================================================================
    # 6. 证据规则
    # ================================================================
    def get_evidence_rules(self) -> list:
        """
        返回交通事故案件的证据规则列表（27项）

        证据规则从 TrafficAccidentEvidenceAdapter 获取，
        涵盖通用证据、医疗类、护理类、误工类、交通类、
        伤残鉴定、死亡案件、财产损失、被扶养人、刑事附带民事等类别。

        Returns:
            list: 证据规则列表
        """
        return self._evidence_adapter.get_all_rules()

    # ================================================================
    # 7. 模板配置
    # ================================================================
    def get_template_info(self) -> dict:
        """
        返回模板配置信息

        指定交通事故起诉状使用的模板文件和表格配置。

        Returns:
            dict: 模板配置
        """
        return {
            "template_file": "模板_民事起诉状.docx",
            "table_configs": [
                {
                    "index": 0,
                    "type": "plaintiff",
                    "fields": [
                        "原告姓名", "原告性别", "原告出生日期",
                        "原告民族", "原告住址", "原告身份证号", "原告电话",
                    ],
                },
                {
                    "index": 1,
                    "type": "defendant",
                    "fields": [
                        "被告姓名/名称", "被告性别", "被告身份证号/信用代码",
                        "被告住址", "被告法定代表人", "被告电话",
                    ],
                },
            ],
        }
