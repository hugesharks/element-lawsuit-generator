# -*- coding: utf-8 -*-
"""
离婚纠纷适配器端到端测试脚本

测试流程：
1. 加载适配器 MarriageFamilyAdapter
2. 准备测试案例数据
3. 调用 build_fill_map 和 get_checkbox_map 生成书签映射
4. 在模板中插入书签
5. 填充书签和勾选框
6. 保存输出文件并验证
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from adapters.marriage_family import MarriageFamilyAdapter
from engine.template_engine import BookmarkEngine


# ============================================================
# 测试案例数据
# ============================================================
TEST_CASE = {
    # === 原告信息 ===
    'plaintiff_name': '张丽',
    'plaintiff_gender': '女',
    'plaintiff_birth_date': '1988年3月15日',
    'plaintiff_ethnicity': '汉族',
    'plaintiff_work_unit': '',
    'plaintiff_position': '',
    'plaintiff_phone': '',
    'plaintiff_address': '河北省邢台市襄都区',
    'plaintiff_residence': '河北省邢台市襄都区',
    'plaintiff_id_type': '居民身份证',
    'plaintiff_id_number': '130502198803151234',
    
    # === 被告信息 ===
    'defendant_name': '王强',
    'defendant_gender': '男',
    'defendant_birth_date': '1986年7月20日',
    'defendant_ethnicity': '汉族',
    'defendant_work_unit': '',
    'defendant_position': '',
    'defendant_phone': '',
    'defendant_address': '河北省邢台市信都区',
    'defendant_residence': '河北省邢台市信都区',
    'defendant_id_type': '居民身份证',
    'defendant_id_number': '130503198607201234',
    
    # === 委托代理人 ===
    'has_agent': True,
    'agent_name': '赵律师',
    'agent_unit': '河北某律师事务所',
    'agent_position': '律师',
    'agent_phone': '13900139002',
    'agent_authorization': '一般授权',
    
    # === 送达地址 ===
    'delivery_address': '',
    
    # === 诉讼请求 ===
    'divorce_request': True,
    'divorce_specific': '因被告婚后多次出轨，严重伤害夫妻感情，双方已分居两年，无法继续共同生活，请求依法解除婚姻关系。',
    
    # 财产分割
    'has_property': True,
    'property_detail': '夫妻共同财产：1. 邢台市房产一处，价值80万元，归原告所有；2. 大众汽车一辆，价值15万元，归被告所有。',
    
    # 债务分担
    'has_debt': True,
    'debt_detail': '房屋贷款30万元，由被告承担。',
    
    # 子女抚养
    'has_children': True,
    'children': [
        {'name': '王小明', 'age': 8, 'custody': '原告', 'monthly_support': 2000},
    ],
    
    # 抚养费
    'has_child_support': True,
    'child_support_payer': '被告',
    'child_support_detail': '每月2000元，于每月15日前支付，至子女年满18周岁止。',
    
    # 探望权
    'has_visitation': True,
    'visitation_holder': '被告',
    'visitation_detail': '每月第二、四周周末，法定节假日轮流抚养。',
    
    # 损害赔偿/补偿/帮助
    'has_damage_compensation': False,
    'damage_type': '',
    'damage_amount': 0,
    
    # 诉讼费用
    'claim_litigation_fee': True,
    
    # 其他请求
    'other_requests': '',
    
    # === 诉前保全 ===
    'has_preservation': False,
    'preservation_court': '',
    'preservation_time': '',
    'preservation_case_number': '',
    
    # === 事实与理由 ===
    'marriage_start_date': '2015年10月1日',
    'children_status': '育有一子王小明，2016年5月出生',
    'living_status': '婚后初期感情尚可，2019年起被告经常夜不归宿，双方矛盾加剧',
    'divorce_reason': '被告婚后多次出轨，严重伤害夫妻感情，双方已分居两年',
    'previous_divorce_suit': '无',
    
    'property_facts': '婚后购买邢台市房产一处，首付30万元，贷款30万元；购买大众汽车一辆，价税合计15万元。',
    'debt_facts': '购房贷款30万元，尚欠银行约25万元。',
    'custody_facts': '子女一直随原告生活，由原告抚养照顾，被告未尽抚养义务。',
    'child_support_facts': '被告有稳定工作收入，月薪约8000元，具备支付抚养费能力。',
    'visitation_facts': '被告作为不直接抚养方，依法享有探望权。',
    'damage_facts': '',
    
    'other_facts': '',
    
    # === 调解意愿 ===
    'understand_mediation': True,
    'understand_mediation_benefits': True,
    'consider_mediation': '是',
}


def setup_test_bookmarks(engine: BookmarkEngine):
    """
    在模板中插入测试书签
    
    根据模板结构，在各个表格的指定位置插入书签。
    书签命名遵循：{表前缀}_{序号}_{字段名}
    """
    print("\n=== 步骤1: 插入书签 ===")
    
    # ============================================================
    # T0: 当事人信息
    # ============================================================
    
    # T0 Row 2: 原告信息 - Cell 1
    # 段落结构：
    # Para 0: 姓名
    # Para 1: 性别勾选
    # Para 2: 出生日期 + 民族
    # Para 3: 工作单位 + 职务 + 联系电话
    # Para 4: 住所地
    # Para 5: 经常居住地
    # Para 6: 证件类型
    # Para 7: 证件号码
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=0, row_index=2, cell_index=1,
        bookmark_prefix='T0_原告',
        field_names=['姓名', '出生日期', '民族', '工作单位', '职务', '联系电话', '住所地', '经常居住地', '证件类型', '证件号码'],
        paragraph_indices=[0, 2, 2, 3, 3, 3, 4, 5, 6, 7]
    )
    print("  ✓ T0 Row 2 原告信息书签已插入")
    
    # T0 Row 3: 委托代理人 - Cell 1
    # 段落结构：
    # Para 0: 有□勾选
    # Para 1: 姓名
    # Para 2: 单位 + 职务 + 联系电话
    # Para 3: 代理权限勾选
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=0, row_index=3, cell_index=1,
        bookmark_prefix='T0_代理人',
        field_names=['姓名', '单位', '职务', '联系电话'],
        paragraph_indices=[1, 2, 2, 2]
    )
    print("  ✓ T0 Row 3 委托代理人书签已插入")
    
    # ============================================================
    # T1: 被告信息 + 诉讼请求
    # ============================================================
    
    # T1 Row 0: 被告信息 - Cell 1
    # 段落结构：同原告
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=1, row_index=0, cell_index=1,
        bookmark_prefix='T1_被告',
        field_names=['姓名', '出生日期', '民族', '工作单位', '职务', '联系电话', '住所地', '经常居住地', '证件类型', '证件号码'],
        paragraph_indices=[0, 2, 2, 3, 3, 3, 4, 5, 6, 7]
    )
    print("  ✓ T1 Row 0 被告信息书签已插入")
    
    # T1 Row 3: 离婚请求 - Cell 1 (具体主张)
    engine.add_bookmarks_to_row(1, 3, 'T1', ['离婚主张'], cell_index=1)
    print("  ✓ T1 Row 3 离婚请求书签已插入")
    
    # T1 Row 4: 夫妻共同财产 - Cell 1
    # 无财产□ 有财产□ 在 Para 0
    # 财产明细在 Para 4
    engine.add_bookmarks_to_row(1, 4, 'T1', ['财产明细'], cell_index=1)
    print("  ✓ T1 Row 4 夫妻共同财产书签已插入")
    
    # T1 Row 5: 夫妻共同债务 - Cell 1
    engine.add_bookmarks_to_row(1, 5, 'T1', ['债务明细'], cell_index=1)
    print("  ✓ T1 Row 5 夫妻共同债务书签已插入")
    
    # T1 Row 6: 子女直接抚养 - Cell 1
    # Para 0: 无此问题□ 有此问题□
    # Para 1: 子女1：归属
    # Para 2: 子女2
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=1, row_index=6, cell_index=1,
        bookmark_prefix='T1',
        field_names=['子女1姓名', '子女1归属', '子女2姓名', '子女2归属'],
        paragraph_indices=[1, 1, 2, 2]
    )
    print("  ✓ T1 Row 6 子女直接抚养书签已插入")
    
    # T1 Row 7: 子女抚养费 - Cell 1
    # Para 0: 无此问题□ 有此问题□ + 承担主体
    # Para 1: 金额及明细
    # Para 2: 支付方式
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=1, row_index=7, cell_index=1,
        bookmark_prefix='T1',
        field_names=['抚养费明细', '支付方式'],
        paragraph_indices=[1, 2]
    )
    print("  ✓ T1 Row 7 子女抚养费书签已插入")
    
    # ============================================================
    # T2: 事实与理由
    # ============================================================
    
    # T2 Row 0: 探望权 - Cell 1
    # Para 0: 无此问题□ 有此问题□
    # Para 1: 有此问题□（备用）
    # Para 2: 探望权行使主体 + 行使方式
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=2, row_index=0, cell_index=1,
        bookmark_prefix='T2',
        field_names=['探望权行使方式'],
        paragraph_indices=[2]
    )
    print("  ✓ T2 Row 0 探望权书签已插入")
    
    # T2 Row 1: 损害赔偿/补偿/帮助 - Cell 1
    # Para 0: 无此问题□
    # Para 1: 离婚损害赔偿□ 金额：
    # Para 2: 离婚经济补偿□ 金额：
    # Para 3: 离婚经济帮助□ 金额：
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=2, row_index=1, cell_index=1,
        bookmark_prefix='T2',
        field_names=['损害赔偿金额', '经济补偿金额', '经济帮助金额'],
        paragraph_indices=[1, 2, 3]
    )
    print("  ✓ T2 Row 1 损害赔偿书签已插入")
    
    # T2 Row 8: 婚姻关系基本情况 - Cell 1
    # Para 0: 结婚时间：
    # Para 1: 生育子女情况：
    # Para 2: 双方生活情况：
    # Para 3: 离婚事由：
    # Para 4: 之前有无提起过离婚诉讼：
    engine.add_bookmarks_to_cell_paragraphs(
        table_index=2, row_index=8, cell_index=1,
        bookmark_prefix='T2',
        field_names=['结婚时间', '生育子女情况', '双方生活情况', '离婚事由', '之前离婚诉讼'],
        paragraph_indices=[0, 1, 2, 3, 4]
    )
    print("  ✓ T2 Row 8 婚姻关系基本情况书签已插入")
    
    # T2 Row 9: 夫妻共同财产情况 - Cell 1
    engine.add_bookmarks_to_row(2, 9, 'T2', ['财产事实'], cell_index=1)
    print("  ✓ T2 Row 9 夫妻共同财产事实书签已插入")
    
    # T2 Row 10: 夫妻共同债务情况 - Cell 1
    engine.add_bookmarks_to_row(2, 10, 'T2', ['债务事实'], cell_index=1)
    print("  ✓ T2 Row 10 夫妻共同债务事实书签已插入")
    
    # T2 Row 11: 子女直接抚养情况 - Cell 1
    engine.add_bookmarks_to_row(2, 11, 'T2', ['抚养事实'], cell_index=1)
    print("  ✓ T2 Row 11 子女直接抚养事实书签已插入")
    
    # T2 Row 12: 子女抚养费情况 - Cell 1
    engine.add_bookmarks_to_row(2, 12, 'T2', ['抚养费事实'], cell_index=1)
    print("  ✓ T2 Row 12 子女抚养费事实书签已插入")
    
    # T2 Row 13: 探望权情况 - Cell 1
    engine.add_bookmarks_to_row(2, 13, 'T2', ['探望权事实'], cell_index=1)
    print("  ✓ T2 Row 13 探望权事实书签已插入")
    
    # T2 Row 14: 损害赔偿/补偿/帮助情况 - Cell 1
    engine.add_bookmarks_to_row(2, 14, 'T2', ['损害事实'], cell_index=1)
    print("  ✓ T2 Row 14 损害事实书签已插入")
    
    # T2 Row 15: 其他 - Cell 1
    engine.add_bookmarks_to_row(2, 15, 'T2', ['其他事实'], cell_index=1)
    print("  ✓ T2 Row 15 其他事实书签已插入")
    
    print("\n所有书签插入完成！")


def fill_template(engine: BookmarkEngine, fill_map: dict, checkbox_map: dict, case_data: dict):
    """填充模板"""
    print("\n=== 步骤2: 填充书签 ===")
    
    # 填充书签
    engine.fill(fill_map)
    filled_count = len(fill_map)
    print(f"  ✓ 已填充 {filled_count} 个书签")
    
    # 填充勾选框
    checkbox_count = engine.fill_checkboxes_squares(checkbox_map)
    print(f"  ✓ 已替换 {checkbox_count} 个勾选框")
    
    # 诉前保全：直接填充（不走书签）
    if case_data.get('has_preservation'):
        print("\n=== 步骤3: 诉前保全直接填充 ===")
        engine.fill_preservation_direct(
            table_index=2,
            row_index=5,
            preservation_court=case_data.get('preservation_court', ''),
            preservation_time=case_data.get('preservation_time', '')
        )
        print("  ✓ 诉前保全信息已填充")


def verify_output(output_path: str, engine: BookmarkEngine):
    """验证输出文件"""
    print("\n=== 步骤4: 验证输出 ===")
    
    # 列出所有书签
    bookmarks = engine.list_bookmarks()
    print(f"\n模板中现有书签 ({len(bookmarks)} 个):")
    for bm in sorted(bookmarks)[:30]:
        print(f"  - {bm}")
    if len(bookmarks) > 30:
        print(f"  ... 还有 {len(bookmarks) - 30} 个书签")
    
    # 检查文件是否存在
    if os.path.exists(output_path):
        file_size = os.path.getsize(output_path)
        print(f"\n✓ 输出文件已生成: {output_path}")
        print(f"  文件大小: {file_size / 1024:.1f} KB")
    else:
        print(f"\n✗ 输出文件未生成: {output_path}")
    
    return len(bookmarks)


def run_test():
    """运行测试"""
    print("=" * 80)
    print("离婚纠纷适配器 - 端到端测试")
    print("=" * 80)
    
    # 路径配置
    template_path = project_root / "template_cache" / "民事起诉状-离婚纠纷.docx"
    output_dir = project_root / "output"
    output_path = output_dir / "test_marriage_family_output.docx"
    
    # 确保输出目录存在
    output_dir.mkdir(exist_ok=True)
    
    # 检查模板是否存在
    if not template_path.exists():
        print(f"\n✗ 错误: 模板文件不存在 - {template_path}")
        return False
    
    print(f"\n模板文件: {template_path}")
    print(f"输出文件: {output_path}")
    
    # 1. 加载适配器
    print("\n=== 初始化 ===")
    adapter = MarriageFamilyAdapter()
    print(f"适配器名称: {adapter.name()}")
    print(f"模板名称: {adapter.get_template_name()}")
    
    # 2. 生成计算结果
    print("\n=== 计算 ===")
    calc_result = adapter.calculate(TEST_CASE)
    print(f"计算项目: {len(calc_result.get('items', []))} 项")
    print(f"标的总额: {calc_result.get('total', 0)} 元")
    if calc_result.get('items'):
        for item in calc_result['items']:
            print(f"  - {item['name']}: {item['amount']} 元")
    
    # 3. 生成书签填充映射
    print("\n=== 生成填充映射 ===")
    fill_map = adapter.build_fill_map(TEST_CASE, calc_result)
    print(f"书签数量: {len(fill_map)} 个")
    
    # 4. 生成勾选框映射
    checkbox_map = adapter.get_checkbox_map(TEST_CASE)
    print(f"勾选框数量: {len(checkbox_map)} 个")
    
    # 5. 加载模板引擎
    print("\n=== 加载模板引擎 ===")
    engine = BookmarkEngine(str(template_path))
    
    # 6. 插入书签
    setup_test_bookmarks(engine)
    
    # 7. 填充模板
    fill_template(engine, fill_map, checkbox_map, TEST_CASE)
    
    # 8. 保存输出
    print("\n=== 保存文件 ===")
    saved_path = engine.save(str(output_path))
    print(f"✓ 文件已保存: {saved_path}")
    
    # 9. 验证
    bookmark_count = verify_output(str(output_path), engine)
    
    # 10. 清理
    engine.cleanup()
    
    print("\n" + "=" * 80)
    print("测试完成!")
    print("=" * 80)
    
    return True


if __name__ == '__main__':
    try:
        success = run_test()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
