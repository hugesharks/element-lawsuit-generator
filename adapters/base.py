# -*- coding: utf-8 -*-
"""
案由适配器基类

定义案由适配器的抽象接口，用于：
- 标准化不同案由的输入字段
- 执行案由特定的计算逻辑
- 生成书签填充映射
- 获取证据规则
- 获取模板信息

使用方式：
    class TrafficAccidentAdapter(CaseAdapter):
        def name(self) -> str:
            return "机动车交通事故责任纠纷"
        
        def get_schema(self) -> dict:
            return {...}
        
        def calculate(self, case_data: dict) -> dict:
            return {...}
        
        def build_fill_map(self, case_data: dict, calc_result: dict) -> dict:
            return {...}
        
        def get_checkbox_map(self, case_data: dict) -> dict:
            return {...}
        
        def get_evidence_rules(self) -> list:
            return [...]
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Any, Optional


class CaseAdapter(ABC):
    """
    案由适配器基类
    
    所有具体案由的适配器都应继承此类并实现抽象方法。
    适配器负责：
    1. 定义输入字段schema
    2. 执行案由特定的计算
    3. 生成书签填充映射
    4. 生成勾选框映射
    5. 提供证据规则
    
    设计原则：
    - 每个案由对应一个适配器
    - 适配器是无状态的（除了配置参数）
    - 计算逻辑封装在calculate方法中
    - 模板映射逻辑封装在build_fill_map方法中
    """
    
    @abstractmethod
    def name(self) -> str:
        """
        案由名称
        
        Returns:
            案由名称，如 '机动车交通事故责任纠纷'
        """
        pass
    
    @abstractmethod
    def get_schema(self) -> dict:
        """
        返回输入字段schema
        
        Returns:
            字段定义字典，格式：
            {
                "fields": [
                    {
                        "name": "plaintiff_name",       # 字段名（英文）
                        "label": "原告姓名",              # 显示标签
                        "type": "str",                   # 数据类型：str/int/float/bool/date
                        "required": True,                # 是否必填
                        "default": None,                 # 默认值
                        "options": None,                 # 可选值列表（如 [1,2,3,4,5,6,7,8,9,10]）
                        "description": "原告真实姓名",    # 字段说明
                    },
                    ...
                ],
                "party_fields": {
                    "plaintiff": [  # 自然人原告
                        {"name": "姓名", "type": "str", "required": True},
                        {"name": "性别", "type": "str", "required": True, "options": ["男", "女"]},
                        {"name": "出生日期", "type": "date", "required": True},
                        {"name": "民族", "type": "str", "required": True},
                        {"name": "住址", "type": "str", "required": True},
                        {"name": "身份证号", "type": "str", "required": True},
                        {"name": "电话", "type": "str", "required": False},
                    ],
                    "defendant_person": [  # 自然人被告
                        {"name": "姓名", "type": "str", "required": True},
                        {"name": "性别", "type": "str", "required": False},
                        {"name": "身份证号", "type": "str", "required": False},
                        {"name": "住址", "type": "str", "required": True},
                    ],
                    "defendant_company": [  # 法人被告
                        {"name": "名称", "type": "str", "required": True},
                        {"name": "住址", "type": "str", "required": True},
                        {"name": "法定代表人", "type": "str", "required": True},
                        {"name": "统一社会信用代码", "type": "str", "required": True},
                    ],
                }
            }
        """
        pass
    
    @abstractmethod
    def calculate(self, case_data: dict) -> dict:
        """
        执行计算（赔偿/费用等）
        
        Args:
            case_data: 案件数据字典，包含所有输入字段的值
            
        Returns:
            计算结果字典，格式：
            {
                "items": [
                    {
                        "name": "医疗费",           # 项目名称
                        "amount": 52340.80,         # 金额
                        "formula": "发票金额之和",   # 计算公式
                        "legal_basis": "《民法典》第1179条",  # 法律依据
                        "detail": "...",            # 详细说明（可选）
                    },
                    ...
                ],
                "total": 123456.78,  # 赔偿总额
                "insurance_split": {  # 保险拆分（可选）
                    "compulsory": {
                        "death_disability_limit": 180000,
                        "death_disability_amount": 0,
                        "medical_limit": 18000,
                        "medical_amount": 18000,
                        "property_limit": 2000,
                        "property_amount": 2000,
                    },
                    "commercial": {
                        "limit": 100000,
                        "amount": 50000,
                    },
                    "self_pay": 53456.78,
                },
            }
        """
        pass
    
    @abstractmethod
    def build_fill_map(self, case_data: dict, calc_result: dict) -> dict[str, str]:
        """
        生成书签填充映射表
        
        将案件数据和计算结果映射为模板书签对应的值
        
        Args:
            case_data: 案件数据字典
            calc_result: calculate方法的返回结果
            
        Returns:
            书签填充映射，格式：
            {
                "T0_01_原告姓名": "张某某",
                "T0_02_原告性别": "男",
                "T0_03_原告出生日期": "1990年1月1日",
                "T2_01_诉讼请求金额": "123456.78",
                "T2_02_医疗费": "52340.80",
                ...
            }
        """
        pass
    
    @abstractmethod
    def get_checkbox_map(self, case_data: dict) -> dict[str, bool]:
        """
        生成勾选框映射表
        
        Args:
            case_data: 案件数据字典
            
        Returns:
            勾选框映射，格式：
            {
                "性别_男": True,
                "性别_女": False,
                "是否住院_是": True,
                "是否住院_否": False,
                "户籍性质_城镇": True,
                "户籍性质_农村": False,
                ...
            }
        """
        pass
    
    @abstractmethod
    def get_evidence_rules(self) -> list:
        """
        返回证据规则列表
        
        Returns:
            EvidenceRule对象列表或字典列表
        """
        pass
    
    def get_template_info(self) -> dict:
        """
        返回模板信息
        
        Returns:
            模板配置字典，格式：
            {
                "template_file": "模板_民事起诉状.docx",  # 主模板文件名
                "table_configs": [  # 表格配置列表
                    {
                        "table_index": 3,           # 表格索引
                        "template_row_index": 0,    # 模板行索引
                        "target_count": 5,          # 目标行数（可为动态值）
                        "bookmark_prefix": "T3",    # 书签前缀
                        "field_names": ["姓名", "性别", ...],  # 字段名列表
                    },
                    ...
                ],
            }
        """
        return {}
    
    def get_party_types(self) -> List[str]:
        """
        获取支持的当事人类型列表
        
        Returns:
            当事人类型列表，如 ["plaintiff", "defendant_person", "defendant_company"]
        """
        schema = self.get_schema()
        return list(schema.get("party_fields", {}).keys())
    
    def get_required_fields(self) -> List[str]:
        """
        获取所有必填字段列表
        
        Returns:
            必填字段name列表
        """
        schema = self.get_schema()
        return [f["name"] for f in schema.get("fields", []) if f.get("required", False)]
    
    def validate_case_data(self, case_data: dict) -> List[str]:
        """
        验证案件数据完整性
        
        Args:
            case_data: 案件数据字典
            
        Returns:
            缺失或无效的字段错误信息列表，空列表表示验证通过
        """
        errors = []
        schema = self.get_schema()
        
        for field in schema.get("fields", []):
            name = field["name"]
            label = field.get("label", name)
            required = field.get("required", False)
            field_type = field.get("type", "str")
            options = field.get("options")
            
            value = case_data.get(name)
            
            # 检查必填
            if required and (value is None or value == ""):
                errors.append(f"缺少必填字段：{label}")
                continue
            
            # 检查类型
            if value is not None and value != "":
                if field_type == "int" and not isinstance(value, int):
                    try:
                        int(value)
                    except (ValueError, TypeError):
                        errors.append(f"字段 {label} 应为整数")
                elif field_type == "float" and not isinstance(value, (int, float)):
                    try:
                        float(value)
                    except (ValueError, TypeError):
                        errors.append(f"字段 {label} 应为数字")
                
                # 检查可选值
                if options and value not in options:
                    errors.append(f"字段 {label} 的值不在可选范围内：{options}")
        
        return errors
