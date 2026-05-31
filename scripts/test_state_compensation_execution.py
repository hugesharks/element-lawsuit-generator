#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
国家赔偿（错误执行）适配器测试脚本
"""

import sys
import os

# 添加项目根目录到路径
SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SCRIPT_DIR)

from adapters.state_compensation_execution import StateCompensationExecutionAdapter


def test_natural_person():
    """测试1: 自然人赔偿请求人 + 完整赔偿项目"""
    print("=" * 60)
    print("测试1: 自然人赔偿请求人 + 完整赔偿项目")
    print("=" * 60)
    
    data = {
        # 自然人信息
        "requester_type": "natural",
        "requester_name": "张三",
        "requester_gender": "男",
        "requester_birthdate": "1980年5月10日",
        "requester_ethnicity": "汉族",
        "requester_work_unit": "某科技公司",
        "requester_position": "工程师",
        "requester_phone": "13800138000",
        "requester_residence": "北京市朝阳区建国路88号",
        "requester_living_place": "北京市海淀区中关村大街100号",
        
        # 委托代理人
        "has_agent": True,
        "agent_type": ["律师"],
        "agent_name": "李律师",
        "agent_unit": "某律师事务所",
        "agent_position": "主任律师",
        "agent_phone": "13900139000",
        "agent_authority": "特别授权",
        
        # 赔偿义务机关
        "compensation_authority_name": "北京市第一中级人民法院",
        "compensation_authority_address": "北京市朝阳区建国门南大街10号",
        "compensation_authority_legal_representative": "王院长",
        "compensation_authority_position": "院长",
        
        # 自赔决定
        "has_self_decision": True,
        "decision_number": "(2024)京01赔决字第001号",
        "decision_result": "决定赔偿人民币50000元",
        
        # 执行标的
        "execution_subject_matter": "依据(2022)京01刑初字第123号生效判决，对申请人名下房产进行强制执行，执行标的为人民币200万元。",
        
        # 财产权赔偿
        "has_property_compensation": True,
        "property_compensation_type": "赔偿损失",
        "property_compensation_amount": "150000",
        
        # 停产停业损失
        "has_business_loss": False,
        
        # 人身自由赔偿
        "has_personal_freedom_compensation": False,
        
        # 其他赔偿
        "other_compensation_requests": "无",
        
        # 事实与理由
        "legal_basis_and_reasons": """一、法律依据
根据《中华人民共和国国家赔偿法》第三十六条规定，侵犯公民、法人和其他组织的财产权造成损害的，按照下列规定处理：
(一) 处罚款、罚金、追缴、没收财产的，返还财产；
(二) 查封、扣押、冻结财产的，解除对财产的查封、扣押、冻结；
(三) 应当返还的财产损坏的，能够恢复原状的恢复原状，不能恢复原状的，按照损害程度给付相应的赔偿金；
(四) 应当返还的财产灭失的，给付相应的赔偿金；
(五) 财产已经拍卖或者变卖的，给付拍卖或者变卖所得的价款；变卖的价款明显低于财产价值的，应当支付相应的赔偿金。

二、事实与理由
1. 申请人因(2022)京01刑初字第123号案件被采取强制执行措施；
2. 该刑事判决后经再审程序被撤销，申请人被宣告无罪；
3. 原执行依据已丧失法律效力，继续执行属于错误执行；
4. 错误执行导致申请人财产损失共计人民币150000元；
5. 赔偿义务机关至今未对错误执行行为予以纠正和赔偿。""",
        
        "other_explanations": "申请人家庭生活困难，上述损失对申请人造成重大影响。",
        
        "has_similar_cases": True,
        "similar_case_number": "(2020)最高法委赔监字第45号指导性案例",
        
        "evidence_list": """1. (2022)京01刑初字第123号刑事判决书
2. (2023)京01刑再字第1号再审刑事判决书（宣告无罪）
3. 执行裁定书及执行标的清单
4. 财产损失鉴定报告
5. 其他相关证据材料""",
    }
    
    adapter = StateCompensationExecutionAdapter(data)
    output_path = adapter.generate()
    
    print(f"\n✓ 生成文件: {output_path}")
    return output_path


def test_legal_person():
    """测试2: 法人赔偿请求人 + 停产停业损失 + 人身自由赔偿"""
    print("\n" + "=" * 60)
    print("测试2: 法人赔偿请求人 + 停产停业损失 + 人身自由赔偿")
    print("=" * 60)
    
    data = {
        # 法人信息
        "requester_type": "legal",
        "requester_company_name": "北京某某科技有限公司",
        "requester_company_address": "北京市海淀区中关村软件园8号楼",
        "requester_company_registered_place": "北京市海淀区市场监督管理局",
        "requester_legal_representative": "赵某某",
        "requester_representative_position": "执行董事兼总经理",
        "requester_phone": "010-88888888",
        "requester_unified_social_credit_code": "91110108MA00XXXXXX",
        "requester_company_type": "有限责任公司",
        
        # 无委托代理人
        "has_agent": False,
        
        # 赔偿义务机关
        "compensation_authority_name": "北京市第二中级人民法院",
        "compensation_authority_address": "北京市朝阳区东大桥路甲8号",
        "compensation_authority_legal_representative": "刘院长",
        "compensation_authority_position": "院长",
        
        # 无自赔决定
        "has_self_decision": False,
        
        # 执行标的
        "execution_subject_matter": "对申请人名下银行存款及办公设备进行强制执行，执行标的为人民币500万元。",
        
        # 财产权赔偿 - 返还原物
        "has_property_compensation": True,
        "property_compensation_type": "返还原物",
        "property_compensation_amount": "",
        
        # 停产停业损失
        "has_business_loss": True,
        "business_loss_type": "留守职工工资",
        "employee_wages_amount": "300000",
        "employee_wages_calculation": "公司原有职工50人，留守人员10人，月平均工资8000元，自2023年1月至2024年6月共18个月",
        "tax_social_fee_amount": "",
        "tax_social_fee_calculation": "",
        "utilities_rent_amount": "",
        "utilities_rent_calculation": "",
        
        # 人身自由赔偿
        "has_personal_freedom_compensation": True,
        "illegal_detention_days": "30",
        "detention_period": "2022年6月1日至2022年6月30日",
        "personal_freedom_compensation_amount": "10593.6",
        
        # 其他赔偿
        "other_compensation_requests": "1. 律师费损失：人民币50000元\n2. 差旅费损失：人民币8000元\n3. 精神损害抚慰金：人民币50000元",
        
        # 事实与理由
        "legal_basis_and_reasons": """根据《国家赔偿法》第三十九条规定，赔偿请求人请求国家赔偿的时效为两年，自其知道或者应当知道国家机关及其工作人员行使职权时的行为侵犯其人身权、财产权之日起计算。

本案中，申请人法定代表人于2024年1月15日收到再审无罪判决，此时方知其人身权、财产权受到侵犯，故本案赔偿请求未超过时效。""",
        
        "other_explanations": "申请人在案件审理期间积极配合调查，现案件已终结，请求尽快处理赔偿事宜。",
        
        "has_similar_cases": False,
        
        "evidence_list": """1. 营业执照复印件
2. 法定代表人身份证明
3. 再审无罪判决书
4. 执行文书
5. 财产损失明细表
6. 职工工资发放记录""",
    }
    
    adapter = StateCompensationExecutionAdapter(data)
    output_path = adapter.generate()
    
    print(f"\n✓ 生成文件: {output_path}")
    return output_path


def test_minimal():
    """测试3: 最小测试 - 仅有必填字段"""
    print("\n" + "=" * 60)
    print("测试3: 最小测试 - 仅有必填字段")
    print("=" * 60)
    
    data = {
        # 自然人信息
        "requester_type": "natural",
        "requester_name": "李四",
        "requester_gender": "女",
        
        # 无代理人
        "has_agent": False,
        
        # 赔偿义务机关
        "compensation_authority_name": "某市中级人民法院",
        
        # 无自赔决定
        "has_self_decision": False,
        
        # 执行标的
        "execution_subject_matter": "执行标的为车辆一辆。",
        
        # 无其他赔偿
        "has_property_compensation": False,
        "has_business_loss": False,
        "has_personal_freedom_compensation": False,
        "other_compensation_requests": "",
        
        "legal_basis_and_reasons": "依据《国家赔偿法》相关规定提出赔偿申请。",
        "other_explanations": "",
        "has_similar_cases": False,
        "evidence_list": "1. 身份证复印件\n2. 判决书\n3. 执行文书",
    }
    
    adapter = StateCompensationExecutionAdapter(data)
    output_path = adapter.generate()
    
    print(f"\n✓ 生成文件: {output_path}")
    return output_path


def main():
    print("国家赔偿（错误执行）适配器 - 端到端测试")
    print("=" * 60)
    
    try:
        output1 = test_natural_person()
        output2 = test_legal_person()
        output3 = test_minimal()
        
        print("\n" + "=" * 60)
        print("✓ 所有测试通过!")
        print("=" * 60)
        print(f"\n生成的文件:")
        print(f"  1. {output1}")
        print(f"  2. {output2}")
        print(f"  3. {output3}")
        return 0
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
