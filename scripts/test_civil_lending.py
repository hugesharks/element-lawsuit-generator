# -*- coding: utf-8 -*-
"""
民间借贷纠纷适配器端到端测试

测试案例：出借人张三借给李四10万元，年利率12%，已逾期3个月未还

测试流程：
1. 加载模板和适配器
2. 准备测试数据
3. 为模板添加书签
4. 填充书签和勾选框
5. 保存并验证输出
"""

import os
import sys
import zipfile
import xml.etree.ElementTree as ET

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from adapters.civil_lending import CivilLendingAdapter
from engine.template_engine import BookmarkEngine

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(PROJECT_ROOT, "template_cache")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
TEMPLATE_FILE = "民事起诉状-民间借贷纠纷.docx"

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def create_test_case_data():
    """创建测试案件数据"""
    return {
        # === 原告（出借人）信息 ===
        "plaintiff_type": "自然人",
        "plaintiff_name": "张三",
        "plaintiff_gender": "男",
        "plaintiff_birthdate": "1985年6月15日",
        "plaintiff_work_unit": "",
        "plaintiff_position": "",
        "plaintiff_address": "北京市朝阳区建国路88号",
        "plaintiff_residence": "北京市朝阳区建国路88号",
        "plaintiff_id_type": "居民身份证",
        "plaintiff_id_number": "110101198506151234",
        
        # === 被告（借款人）信息 ===
        "defendant_type": "自然人",
        "defendant_name": "李四",
        "defendant_gender": "男",
        "defendant_birthdate": "1980年3月20日",
        "defendant_work_unit": "",
        "defendant_position": "",
        "defendant_address": "北京市海淀区中关村大街1号",
        "defendant_residence": "北京市海淀区中关村大街1号",
        "defendant_id_type": "居民身份证",
        "defendant_id_number": "110108198003201234",
        
        # === 第三人 ===
        "has_third_party": False,
        
        # === 委托代理人 ===
        "has_agent": False,
        
        # === 合同信息 ===
        "contract_name": "借款合同",
        "contract_number": "JK-2024-001",
        "contract_sign_date": "2023年10月1日",
        "contract_sign_place": "北京市",
        
        # === 借款信息 ===
        "lender_name": "张三",
        "borrower_name": "李四",
        "loan_amount_agreed": 100000.0,
        "loan_amount_actual": 100000.0,
        "loan_provision_method": "转账",
        "loan_provision_method_other": "",
        
        # === 借款期限 ===
        "loan_start_date": "2023年10月1日",
        "loan_end_date": "2024年10月1日",
        "is_matured": True,  # 已到期
        
        # === 利率 ===
        "interest_rate": 12.0,  # 年利率12%
        "interest_rate_type": "年",
        "contract_article_number": "第三条",
        
        # === 实际提供时间 ===
        "actual_provision_date": "2023年10月1日",
        "actual_provision_amount": 100000.0,
        
        # === 还款方式 ===
        "repayment_method": "到期一次性还本付息",
        "repayment_method_other": "",
        
        # === 还款情况 ===
        "repaid_principal": 0.0,  # 未还本金
        "repaid_interest": 0.0,   # 未还利息
        "interest_paid_to_date": "",  # 未支付利息
        
        # === 逾期 ===
        "is_overdue": True,  # 已逾期
        "overdue_start_date": "2024年10月2日",
        "overdue_months": 3,
        
        # === 物的担保 ===
        "has_property_guarantee": False,
        "guarantor_name": "",
        "guarantee_property": "",
        "has_maximum_guarantee": False,
        
        # === 抵押质押登记 ===
        "has_registration": False,
        
        # === 保证担保 ===
        "has_surety": False,
        
        # === 其他担保 ===
        "has_other_guarantee": False,
        
        # === 其他说明 ===
        "other_matters": "",
        
        # === 诉讼请求 ===
        "principal_calc_date": "2025年1月15日",
        "principal_amount": 100000.0,  # 本金10万元
        "interest_calc_date": "2025年1月15日",
        "interest_amount": 15320.0,  # 利息15320元（100000 * 12% / 365 * 466天 ≈ 15320）
        "request_interest_to_actual": True,  # 请求支付至实际清偿日
        
        "request_acceleration": False,  # 不要求加速到期
        "request_rescission": False,    # 不要求解除合同
        "request_guarantee_rights": False,
        "request_creditor_expenses": False,
        "request_litigation_fee": True,  # 要求被告承担诉讼费
        "other_requests": "",
        
        # === 标的总额 ===
        "total_amount": 115320.0,  # 本金 + 利息（不含后续利息）
        
        # === 约定管辖 ===
        "has_jurisdiction_agreement": False,
        
        # === 诉前保全 ===
        "has_preservation": False,
        
        # === 请求依据 ===
        "contract_basis": "借款合同第三条约定",
        "legal_basis": "《民法典》第667条、第675条，《最高人民法院关于审理民间借贷案件适用法律若干问题的规定》(2020修订)第25条、第29条",
        
        # === 调解意愿 ===
        "understand_mediation": True,
        "understand_mediation_benefits": True,
        "consider_mediation": "是",
        
        # === 证据清单 ===
        "evidence_list": "1. 借款合同；2. 转账凭证",
    }


def add_bookmarks_to_template(engine: BookmarkEngine, case_data: dict):
    """
    为模板添加书签
    
    书签设计基于模板段落结构：
    - T0: 原告信息（自然人 Row 2，法人 Row 3）
    - T1: 委托代理人 Row 0、被告自然人 Row 1、被告法人 Row 2、第三人 Row 3
    - T2: 诉讼请求（Row 3-12）
    - T3: 事实与理由 Part 1（Row 2-13）
    - T4: 事实与理由 Part 2（Row 0-5）
    - T5: 调解意愿（Row 0）
    """
    
    # === T0: 原告信息 ===
    plaintiff_type = case_data.get('plaintiff_type', '自然人')
    
    if plaintiff_type == '自然人':
        # T0 Row 2 原告自然人 - 使用 add_bookmarks_to_cell_paragraphs
        # 段落结构: Para0-姓名, Para1-性别, Para2-出生日期+民族, Para3-工作单位+职务+电话, Para4-住所地, Para5-经常居住地, Para6-证件类型, Para7-证件号码
        engine.add_bookmarks_to_cell_paragraphs(
            table_index=0, 
            row_index=2, 
            cell_index=1,
            bookmark_prefix="T0_01",
            field_names=["姓名", "性别", "出生日期", "工作单位职务电话", "住所地", "经常居住地", "证件类型", "证件号码"],
            paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7]
        )
    else:
        # T0 Row 3 原告法人
        engine.add_bookmarks_to_cell_paragraphs(
            table_index=0, 
            row_index=3, 
            cell_index=1,
            bookmark_prefix="T0_01",
            field_names=["名称", "住所地", "注册地", "法定代表人职务电话", "统一社会信用代码", "类型", "类型续", "所有制性质"],
            paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7, 8]
        )
    
    # === T1: 委托代理人 + 被告 + 第三人 ===
    has_agent = case_data.get('has_agent', False)
    if has_agent:
        # T1 Row 0 委托代理人
        engine.add_bookmarks_to_cell_paragraphs(
            table_index=1, 
            row_index=0, 
            cell_index=1,
            bookmark_prefix="T1_01",
            field_names=["姓名", "单位职务电话", "代理权限"],
            paragraph_indices=[1, 2, 3]
        )
    
    defendant_type = case_data.get('defendant_type', '自然人')
    has_third_party = case_data.get('has_third_party', False)
    
    if defendant_type == '自然人':
        # T1 Row 1 被告自然人
        engine.add_bookmarks_to_cell_paragraphs(
            table_index=1, 
            row_index=1, 
            cell_index=1,
            bookmark_prefix="T1_05",
            field_names=["姓名", "性别", "出生日期", "工作单位职务电话", "住所地", "经常居住地", "证件类型", "证件号码"],
            paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7]
        )
    else:
        # T1 Row 2 被告法人
        engine.add_bookmarks_to_cell_paragraphs(
            table_index=1, 
            row_index=2, 
            cell_index=1,
            bookmark_prefix="T1_05",
            field_names=["名称", "住所地", "注册地", "法定代表人职务电话", "统一社会信用代码", "类型", "类型续", "所有制性质"],
            paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7, 8]
        )
    
    if has_third_party:
        tp_type = case_data.get('third_party_type', '自然人')
        if tp_type == '自然人':
            # T1 Row 3 第三人自然人
            engine.add_bookmarks_to_cell_paragraphs(
                table_index=1, 
                row_index=3, 
                cell_index=1,
                bookmark_prefix="T1_14",
                field_names=["姓名", "性别", "出生日期", "工作单位职务电话", "住所地", "经常居住地", "证件类型", "证件号码"],
                paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7]
            )
        else:
            # T1 Row 4 第三人法人（实际在T2 Row 0）
            engine.add_bookmarks_to_cell_paragraphs(
                table_index=2, 
                row_index=0, 
                cell_index=1,
                bookmark_prefix="T1_17",
                field_names=["名称", "住所地", "注册地", "法定代表人职务电话", "统一社会信用代码", "类型", "类型续", "所有制性质"],
                paragraph_indices=[0, 1, 2, 3, 4, 5, 6, 7, 8]
            )
    
    # === T2: 诉讼请求 ===
    # T2 Row 3 本金 - 使用 cell_index=1 让书签插入 Cell 1
    engine.add_bookmarks_to_row(2, 3, "T2_01", ["本金截止日期"], cell_index=1)
    engine.add_bookmarks_to_row(2, 3, "T2_02", ["本金金额"], cell_index=1)
    
    # T2 Row 4 利息
    engine.add_bookmarks_to_row(2, 4, "T2_03", ["利息截止日期"], cell_index=1)
    engine.add_bookmarks_to_row(2, 4, "T2_04", ["利息金额"], cell_index=1)
    
    # T2 Row 5 提前还款/解除合同
    engine.add_bookmarks_to_row(2, 5, "T2_05", ["提前还款解除合同"], cell_index=1)
    
    # T2 Row 6 担保权利
    engine.add_bookmarks_to_row(2, 6, "T2_06", ["担保权利内容"], cell_index=1)
    
    # T2 Row 7 债权费用
    engine.add_bookmarks_to_row(2, 7, "T2_07", ["债权费用明细"], cell_index=1)
    
    # T2 Row 9 其他请求
    engine.add_bookmarks_to_row(2, 9, "T2_08", ["其他请求"], cell_index=1)
    
    # T2 Row 10 标的总额
    engine.add_bookmarks_to_row(2, 10, "T2_09", ["标的总额"], cell_index=1)
    
    # T2 Row 12 约定管辖
    engine.add_bookmarks_to_row(2, 12, "T2_10", ["管辖条款"], cell_index=1)
    
    # === T3: 事实与理由 Part 1 ===
    # T3 Row 2 合同签订情况
    engine.add_bookmarks_to_row(3, 2, "T3_01", ["合同名称", "合同编号", "合同签订日期", "合同签订地点"], cell_index=1)
    
    # T3 Row 3 签订主体
    engine.add_bookmarks_to_row(3, 3, "T3_05", ["出借人", "借款人"], cell_index=1)
    
    # T3 Row 4 借款金额
    engine.add_bookmarks_to_row(3, 4, "T3_07", ["约定金额", "实际提供金额", "提供方式现金", "提供方式转账", "提供方式其他"], cell_index=1)
    
    # T3 Row 5 借款期限
    engine.add_bookmarks_to_row(3, 5, "T3_12", ["是否到期", "到期日期"], cell_index=1)
    
    # T3 Row 6 借款利率
    engine.add_bookmarks_to_row(3, 6, "T3_14", ["利率", "合同条款"], cell_index=1)
    
    # T3 Row 7 借款提供时间
    engine.add_bookmarks_to_row(3, 7, "T3_16", ["提供日期", "提供金额"], cell_index=1)
    
    # T3 Row 8 还款方式
    engine.add_bookmarks_to_row(3, 8, "T3_18", ["到期一次性", "按月计息", "按季计息", "按年计息", "其他还款方式"], cell_index=1)
    
    # T3 Row 9 还款情况
    engine.add_bookmarks_to_row(3, 9, "T3_23", ["已还本金", "已还利息", "还息截止日期"], cell_index=1)
    
    # T3 Row 10 逾期
    engine.add_bookmarks_to_row(3, 10, "T3_26", ["是否逾期", "逾期起始日期", "逾期月数"], cell_index=1)
    
    # T3 Row 11 物的担保
    engine.add_bookmarks_to_row(3, 11, "T3_29", ["是否物的担保", "签订时间"], cell_index=1)
    
    # T3 Row 12 担保人担保物
    engine.add_bookmarks_to_row(3, 12, "T3_31", ["担保人", "担保物"], cell_index=1)
    
    # T3 Row 13 最高额担保
    engine.add_bookmarks_to_row(3, 13, "T3_33", ["是否最高额", "确定时间", "担保额度"], cell_index=1)
    
    # === T4: 事实与理由 Part 2 ===
    # T4 Row 0 抵押质押登记
    engine.add_bookmarks_to_row(4, 0, "T4_01", ["是否登记", "登记类型正式", "登记类型预告"], cell_index=1)
    
    # T4 Row 1 保证合同
    engine.add_bookmarks_to_row(4, 1, "T4_04", ["是否保证", "保证签订时间", "保证人", "主要内容", "一般保证", "连带责任"], cell_index=1)
    
    # T4 Row 2 其他担保
    engine.add_bookmarks_to_row(4, 2, "T4_10", ["是否其他担保", "形式", "签订时间"], cell_index=1)
    
    # T4 Row 3 其他说明
    engine.add_bookmarks_to_row(4, 3, "T4_13", ["其他说明"], cell_index=1)
    
    # T4 Row 4 请求依据
    engine.add_bookmarks_to_row(4, 4, "T4_14", ["合同约定", "法律规定"], cell_index=1)
    
    # T4 Row 5 证据清单
    engine.add_bookmarks_to_row(4, 5, "T4_16", ["证据清单"], cell_index=1)
    
    print("书签添加完成！")
    print(f"所有书签: {engine.list_bookmarks()}")


def main():
    """主测试函数"""
    print("=" * 60)
    print("民间借贷纠纷适配器端到端测试")
    print("=" * 60)
    
    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # 模板路径
    template_path = os.path.join(TEMPLATE_DIR, TEMPLATE_FILE)
    if not os.path.exists(template_path):
        print(f"错误: 模板文件不存在: {template_path}")
        sys.exit(1)
    
    # 创建测试数据
    case_data = create_test_case_data()
    print(f"\n测试案件数据:")
    print(f"  出借人: {case_data['lender_name']}")
    print(f"  借款人: {case_data['borrower_name']}")
    print(f"  借款金额: {case_data['loan_amount_actual']}元")
    print(f"  年利率: {case_data['interest_rate']}%")
    print(f"  是否逾期: {case_data['is_overdue']}")
    print(f"  本金: {case_data['principal_amount']}元")
    print(f"  利息: {case_data['interest_amount']}元")
    
    # 加载适配器
    adapter = CivilLendingAdapter()
    print(f"\n适配器名称: {adapter.name()}")
    print(f"模板文件: {adapter.get_template_name()}")
    
    # 执行计算
    calc_result = adapter.calculate(case_data)
    print(f"\n计算结果:")
    print(f"  标的总额: {calc_result['total']}元")
    for item in calc_result['items']:
        print(f"  - {item['name']}: {item['amount']}元 ({item['formula']})")
    
    # 生成书签填充映射
    fill_map = adapter.build_fill_map(case_data, calc_result)
    print(f"\n书签填充映射数量: {len(fill_map)}")
    
    # 生成勾选框映射
    checkbox_map = adapter.get_checkbox_map(case_data)
    print(f"勾选框映射数量: {len(checkbox_map)}")
    
    # 创建引擎并添加书签
    print("\n" + "-" * 40)
    print("步骤1: 为模板添加书签")
    print("-" * 40)
    
    engine = BookmarkEngine(template_path)
    add_bookmarks_to_template(engine, case_data)
    
    # 填充书签
    print("\n" + "-" * 40)
    print("步骤2: 填充书签内容")
    print("-" * 40)
    
    # 根据原告类型调整填充映射
    plaintiff_type = case_data.get('plaintiff_type', '自然人')
    
    if plaintiff_type == '自然人':
        # 原告自然人字段
        fill_map['T0_01_姓名'] = case_data.get('plaintiff_name', '')
        fill_map['T0_01_性别'] = ''
        fill_map['T0_01_出生日期'] = case_data.get('plaintiff_birthdate', '')
        fill_map['T0_01_工作单位职务电话'] = ''
        fill_map['T0_01_住所地'] = case_data.get('plaintiff_address', '')
        fill_map['T0_01_经常居住地'] = case_data.get('plaintiff_residence', '')
        fill_map['T0_01_证件类型'] = case_data.get('plaintiff_id_type', '')
        fill_map['T0_01_证件号码'] = case_data.get('plaintiff_id_number', '')
    else:
        # 原告法人字段
        fill_map['T0_01_名称'] = case_data.get('plaintiff_name', '')
        fill_map['T0_01_住所地'] = case_data.get('plaintiff_address', '')
        fill_map['T0_01_注册地'] = case_data.get('plaintiff_registration_address', '')
        fill_map['T0_01_法定代表人职务电话'] = ''
        fill_map['T0_01_统一社会信用代码'] = case_data.get('plaintiff_id_number', '')
        fill_map['T0_01_类型'] = ''
        fill_map['T0_01_类型续'] = ''
        fill_map['T0_01_所有制性质'] = ''
    
    # 被告字段
    defendant_type = case_data.get('defendant_type', '自然人')
    if defendant_type == '自然人':
        fill_map['T1_05_姓名'] = case_data.get('defendant_name', '')
        fill_map['T1_05_性别'] = ''
        fill_map['T1_05_出生日期'] = case_data.get('defendant_birthdate', '')
        fill_map['T1_05_工作单位职务电话'] = ''
        fill_map['T1_05_住所地'] = case_data.get('defendant_address', '')
        fill_map['T1_05_经常居住地'] = case_data.get('defendant_residence', '')
        fill_map['T1_05_证件类型'] = case_data.get('defendant_id_type', '')
        fill_map['T1_05_证件号码'] = case_data.get('defendant_id_number', '')
    else:
        fill_map['T1_05_名称'] = case_data.get('defendant_name', '')
        fill_map['T1_05_住所地'] = case_data.get('defendant_address', '')
        fill_map['T1_05_注册地'] = case_data.get('defendant_registration_address', '')
        fill_map['T1_05_法定代表人职务电话'] = ''
        fill_map['T1_05_统一社会信用代码'] = case_data.get('defendant_id_number', '')
        fill_map['T1_05_类型'] = ''
        fill_map['T1_05_类型续'] = ''
        fill_map['T1_05_所有制性质'] = ''
    
    # 填充书签
    engine.fill(fill_map)
    print(f"已填充 {len(fill_map)} 个书签")
    
    # 填充勾选框
    print("\n" + "-" * 40)
    print("步骤3: 填充勾选框")
    print("-" * 40)
    
    replaced = engine.fill_checkboxes_squares(checkbox_map)
    print(f"已替换 {replaced} 个勾选框")
    
    # 诉前保全直接处理
    if case_data.get('has_preservation'):
        print("\n" + "-" * 40)
        print("步骤4: 处理诉前保全")
        print("-" * 40)
        engine.fill_preservation_direct(
            table_index=2,
            row_index=13,  # T2 Row 13
            preservation_court=case_data.get('preservation_court', ''),
            preservation_time=case_data.get('preservation_time', '')
        )
        print("诉前保全信息已填写")
    
    # 保存文件
    print("\n" + "-" * 40)
    print("步骤5: 保存文件")
    print("-" * 40)
    
    output_path = os.path.join(OUTPUT_DIR, "民间借贷纠纷_测试_张三诉李四.docx")
    engine.save(output_path)
    print(f"输出文件: {output_path}")
    
    # 验证
    print("\n" + "-" * 40)
    print("步骤6: 验证输出")
    print("-" * 40)
    
    verify_output(output_path, case_data)
    
    engine.cleanup()
    
    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)
    
    return output_path


def verify_output(output_path: str, case_data: dict):
    """验证输出文件"""
    print(f"\n验证文件: {output_path}")
    
    # 检查文件是否存在
    if not os.path.exists(output_path):
        print("  ✗ 文件不存在")
        return False
    
    # 检查文件大小
    file_size = os.path.getsize(output_path)
    print(f"  文件大小: {file_size} bytes")
    
    # 解压并检查内容
    with zipfile.ZipFile(output_path, 'r') as z:
        doc_xml = z.read('word/document.xml')
    
    root = ET.fromstring(doc_xml)
    
    # 统计书签数量
    bookmarks = []
    for bm in root.iter(f'{W}bookmarkStart'):
        name = bm.get('name')
        if name and not name.startswith('_'):
            bookmarks.append(name)
    
    print(f"  书签数量: {len(bookmarks)}")
    
    # 检查关键内容
    full_text = ''.join(t.text for t in root.iter(f'{W}t') if t.text)
    
    checks = [
        ("张三", "原告姓名"),
        ("李四", "被告姓名"),
        ("100000", "借款金额"),
        ("12", "年利率"),
    ]
    
    all_passed = True
    for keyword, desc in checks:
        if keyword in full_text:
            print(f"  ✓ {desc}存在")
        else:
            print(f"  ✗ {desc}不存在")
            all_passed = False
    
    # 检查勾选框
    if '☑' in full_text or '■' in full_text:
        checkbox_count = full_text.count('☑') + full_text.count('■')
        print(f"  ✓ 已勾选勾选框数量: {checkbox_count}")
    else:
        print("  ✗ 未找到已勾选的勾选框")
    
    return all_passed


if __name__ == "__main__":
    main()
