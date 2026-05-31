# -*- coding: utf-8 -*-
"""
端到端测试：案数据 → 映射表 → 模板调整 → 打书签 → 填充 → 输出.docx

测试流程：
1. 创建测试案件数据
2. TrafficAccidentAdapter 计算赔偿 + 生成映射表
3. BookmarkEngine 加载空白模板
4. 动态调整表格行数（原告/被告/赔偿项目）
5. 给调整后的行打书签
6. 填充书签内容 + 勾选框
7. 保存输出.docx

运行：cd element-lawsuit-generator && python3 scripts/end_to_end_test.py
"""

import sys
import os
import json

# 确保可以导入项目模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from adapters.traffic_accident import TrafficAccidentAdapter
from engine.template_engine import BookmarkEngine


# ============================================================
# 1. 测试案件数据
# ============================================================
TEST_CASE = {
    # === 原告信息（2人，模拟死亡案件多名近亲属起诉） ===
    "plaintiffs": [
        {
            "name": "张大明",
            "gender": "男",
            "birthdate": "1985年03月15日",
            "ethnicity": "汉族",
            "address": "河北省邢台市襄都区新华南路123号",
            "id_number": "130502198503151234",
            "phone": "13800138001",
        },
        {
            "name": "李小红",
            "gender": "女",
            "birthdate": "1987年08月22日",
            "ethnicity": "汉族",
            "address": "河北省邢台市襄都区新华南路123号",
            "id_number": "130502198708221234",
            "phone": "13800138002",
        },
    ],

    # === 案件类型 ===
    "case_type": "civil",           # 纯民事
    "injury_type": "disability",    # 伤残
    "liability_type": "主责",       # 被告主要责任

    # === 赔偿标准 ===
    "standard_year": 2025,
    "region": "urban",

    # === 医疗信息 ===
    "hospital_days": 30,
    "medical_fee": 35000,
    "nursing_days": 60,
    "nursing_persons": 1,
    "nutrition_days": 90,
    "lost_work_days": 120,

    # === 伤残信息 ===
    "disability_grade": 10,
    "plaintiff_age": 40,

    # === 被扶养人 ===
    "dependents": [
        {"age": 8, "years": 10, "share_ratio": 0.5, "relation": "女儿"},
    ],

    # === 其他费用 ===
    "traffic_fee": 2000,
    "property_damage": 5000,
    "other_fee": 8000,
    "other_fee_desc": "鉴定费",

    # === 被告信息（1自然人 + 1法人/保险公司） ===
    "defendants_person": [
        {
            "name": "王大力",
            "gender": "男",
            "id_number": "130502199001011234",
            "address": "河北省邢台市信都区钢铁北路456号",
        },
    ],
    "defendants_company": [
        {
            "name": "中国平安财产保险股份有限公司邢台中心支公司",
            "address": "河北省邢台市襄都区中兴东大街88号",
            "legal_person": "赵经理",
            "credit_code": "91130500MA0XXXXX",
        },
    ],

    # === 事故信息 ===
    "accident_time": "2025年10月15日14时30分",
    "accident_location": "邢台市襄都区新华南路与团结路交叉口",
    "accident_detail": (
        "2025年10月15日14时30分许，被告王大力驾驶冀E12345号小型轿车，"
        "沿新华南路由南向北行驶至团结路交叉口时，闯红灯与正常过马路的"
        "原告张大明发生碰撞，造成张大明受伤、车辆损坏的交通事故。"
        "经邢台市公安交通警察支队认定，王大力负事故主要责任。"
    ),
    "responsibility_doc_number": "邢公交认字[2025]第1015号",
    "responsibility_result": "被告王大力负事故主要责任，原告张大明负次要责任",

    # === 保险信息 ===
    "insurance_info": (
        "肇事车辆冀E12345号在中国平安财产保险股份有限公司投保了"
        "交强险（保单号：PDAA202513050000001234）和商业三者险100万元"
        "（保单号：PDBT202513050000005678），事故发生在保险期间内。"
    ),

    # === 诉前保全 ===
    "pre_litigation_preservation": True,
    "preservation_court": "邢台市襄都区人民法院",
    "preservation_time": "2025年11月5日",

    # === 法院信息 ===
    "court": "邢台市襄都区人民法院",
    "filing_date": "2026年1月10日",
}


# ============================================================
# 2. 主测试流程
# ============================================================
def main():
    print("=" * 60)
    print("要素式文书生成 端到端测试")
    print("=" * 60)

    # --- Step 1: 初始化适配器 ---
    print("\n[Step 1] 初始化交通事故适配器...")
    adapter = TrafficAccidentAdapter()
    print(f"  案由: {adapter.name()}")

    # --- Step 2: 计算赔偿 ---
    print("\n[Step 2] 计算赔偿...")
    calc_result = adapter.calculate(TEST_CASE)
    print(f"  赔偿项目数: {len(calc_result['items'])}")
    print(f"  总金额: ¥{calc_result['total']:,.2f}")
    print(f"  交强险合计: ¥{calc_result['insurance_split']['compulsory_total']:,.2f}")
    print(f"  商业险: ¥{calc_result['insurance_split']['commercial']:,.2f}")
    print(f"  自担: ¥{calc_result['insurance_split']['self_bear']:,.2f}")

    # 打印赔偿明细
    print("\n  赔偿明细:")
    for item in calc_result['items']:
        print(f"    {item['name']}: ¥{item['amount']:,.2f}  ({item['formula']})")

    # --- Step 3: 生成映射表 ---
    print("\n[Step 3] 生成书签映射表...")
    fill_map = adapter.build_fill_map(TEST_CASE, calc_result)
    print(f"  映射条数: {len(fill_map)}")
    for key, val in sorted(fill_map.items()):
        print(f"    {key} = {val[:40]}{'...' if len(val) > 40 else ''}")

    # --- Step 4: 生成勾选框映射 ---
    print("\n[Step 4] 生成勾选框映射...")
    checkbox_map = adapter.get_checkbox_map(TEST_CASE)
    print(f"  勾选框数: {len(checkbox_map)}")
    for key, val in checkbox_map.items():
        mark = "☑" if val else "☐"
        print(f"    {mark} {key}")

    # --- Step 5: 加载空白模板 ---
    print("\n[Step 5] 加载空白模板...")
    template_path = os.path.join(
        os.path.dirname(__file__), '..', 'template_cache',
        '民事起诉状-机动车交通事故责任纠纷.docx'
    )
    template_path = os.path.abspath(template_path)
    print(f"  模板路径: {template_path}")

    engine = BookmarkEngine(template_path)

    # 列出模板中已有书签
    existing_bookmarks = engine.list_bookmarks()
    print(f"  模板已有书签: {len(existing_bookmarks)}")
    for bm in existing_bookmarks:
        print(f"    - {bm}")

    # --- Step 6: 动态调整表格行数 ---
    print("\n[Step 6] 动态调整表格行数...")

    # 原告人数
    num_plaintiffs = len(TEST_CASE.get('plaintiffs', []))
    # 被告人数（自然人 + 法人）
    num_defendants = (
        len(TEST_CASE.get('defendants_person', []))
        + len(TEST_CASE.get('defendants_company', []))
    )
    # 赔偿项目数
    num_claim_items = len(calc_result['items'])

    print(f"  原告: {num_plaintiffs}人, 被告: {num_defendants}人, 赔偿项: {num_claim_items}项")

    # T0 表格：原告区
    # 模板 T0 有4行：行0=说明, 行1=标题, 行2=原告自然人, 行3=原告法人
    # 需要：说明行(1) + 标题行(1) + 原告自然人行(N) + 原告法人行(1) = N+3
    t0_target = 3 + num_plaintiffs
    if t0_target > 4:
        t0_prefixes = engine.adjust_table_rows(
            table_index=0,
            target_count=t0_target,
            template_row_index=2,  # 以原告自然人行作为模板
        )
        print(f"  T0（原告表）新增行: {len(t0_prefixes)}，总行数→{t0_target}")
    else:
        t0_prefixes = []
        print(f"  T0（原告表）无需调整")

    # T1 表格：被告区
    # 模板 T1 有4行：行0=代理人, 行1=被告自然人, 行2=被告法人, 行3=第三人
    # 被告自然人从行1开始
    num_person_def = len(TEST_CASE.get('defendants_person', []))
    num_company_def = len(TEST_CASE.get('defendants_company', []))
    # 需要的行数 = 代理人行 + 被告自然人行 + 被告法人行 + 第三人行
    # 模板已有1行被告自然人+1行被告法人，如果需要更多
    t1_target = 1 + max(num_person_def, 1) + max(num_company_def, 1) + 1
    t1_current = 4  # 模板有4行
    if t1_target > t1_current:
        t1_prefixes = engine.adjust_table_rows(
            table_index=1,
            target_count=t1_target,
            template_row_index=1,  # 以被告自然人行作为模板
        )
        print(f"  T1（被告表）新增行: {len(t1_prefixes)}")
    else:
        t1_prefixes = []
        print(f"  T1（被告表）无需调整")

    # T2 表格：诉讼请求区
    # 模板 T2 有13行：行0=第三人法人, 行1=标题, 行2=说明, 行3-12=赔偿项1-10
    # 赔偿项目从行3开始，模板有10个位置（行3-12）
    # 如果赔偿项超过10个，需要扩展
    t2_claim_start = 3
    t2_available = 13 - t2_claim_start  # 10个位置
    if num_claim_items > t2_available:
        t2_target = t2_claim_start + num_claim_items
        t2_prefixes = engine.adjust_table_rows(
            table_index=2,
            target_count=t2_target,
            template_row_index=3,  # 以第一个赔偿项行作为模板
        )
        print(f"  T2（诉讼请求表）新增行: {len(t2_prefixes)}")
    else:
        t2_prefixes = []
        print(f"  T2（诉讼请求表）无需调整（{num_claim_items}项 ≤ {t2_available}个位置）")

    # --- Step 7: 给调整后的行打书签 ---
    print("\n[Step 7] 给表格行打书签...")

    # T0 原告区：给每个原告行打书签
    plaintiff_fields = ["原告姓名", "原告性别", "原告出生日期", "原告民族",
                        "原告住址", "原告身份证号", "原告电话"]
    for i in range(num_plaintiffs):
        row_idx = 2 + i  # 从行2开始
        prefix = f"T0_{i+1:02d}"
        engine.add_bookmarks_to_row(0, row_idx, prefix, plaintiff_fields)
        print(f"  T0 行{row_idx}: 书签前缀 {prefix}")

    # T1 被告区：给被告行打书签
    defendant_person_fields = ["被告姓名", "被告性别", "被告身份证号", "被告住址"]
    for i in range(num_person_def):
        row_idx = 1 + i  # 从行1开始
        prefix = f"T1_{i+1:02d}"
        engine.add_bookmarks_to_row(1, row_idx, prefix, defendant_person_fields)
        print(f"  T1 行{row_idx}: 书签前缀 {prefix}（自然人）")

    defendant_company_fields = ["被告名称", "被告住址", "被告法定代表人", "被告信用代码"]
    for i in range(num_company_def):
        row_idx = 1 + num_person_def + i  # 紧跟自然人后面
        prefix = f"T1_{num_person_def + i + 1:02d}"
        engine.add_bookmarks_to_row(1, row_idx, prefix, defendant_company_fields)
        print(f"  T1 行{row_idx}: 书签前缀 {prefix}（法人）")

    # T2 诉讼请求区：给赔偿项行打书签
    # 先清理Cell 0中的模板标签（如"1. 医疗费"），避免与书签填充内容叠加
    claim_fields = ["赔偿项目", "赔偿金额", "计算公式", "法律依据"]
    for i in range(num_claim_items):
        row_idx = 3 + i  # 从行3开始
        engine.clear_cell_extra_paragraphs(2, row_idx, 0)
        prefix = f"T2_{i+1:02d}"
        engine.add_bookmarks_to_row(2, row_idx, prefix, claim_fields)
        print(f"  T2 行{row_idx}: 书签前缀 {prefix}")

    # T3 事实与理由区：给关键行打书签
    # 行8=事故经过, 行9=责任认定, 行10=投保情况, 行11=请求依据, 行12=证据清单
    t3_fields_map = {
        8: ("T3_01", ["事故经过"]),
        9: ("T3_02", ["责任认定"]),
        10: ("T3_03", ["保险信息"]),
        11: ("T3_04", ["请求依据"]),
        12: ("T3_05", ["证据清单"]),
    }
    for row_idx, (prefix, fields) in t3_fields_map.items():
        engine.add_bookmarks_to_row(3, row_idx, prefix, fields)
        print(f"  T3 行{row_idx}: 书签前缀 {prefix}")

    # T3 行4 诉前保全：不打书签！
    # 因为"是□/否□"勾选框和"保全法院/保全时间"字段在同一个段落里，
    # 打书签会把勾选框也清掉。改为在勾选框处理阶段直接文本替换。

    # --- Step 8: 填充书签内容 ---
    print("\n[Step 8] 填充书签内容...")
    engine.fill(fill_map)
    print(f"  已填充 {len(fill_map)} 个书签")

    # --- Step 8.5: T3 Row 4 诉前保全直接替换（不走书签）---
    if TEST_CASE.get('pre_litigation_preservation'):
        engine.fill_preservation_direct(
            table_index=3,
            row_index=4,
            preservation_court=TEST_CASE.get('preservation_court', ''),
            preservation_time=TEST_CASE.get('preservation_time', '')
        )
        print("  T3 行4: 诉前保全直接替换完成")

    # --- Step 9: 填充事实与理由区的特殊字段 ---
    print("\n[Step 9] 填充事实与理由区...")
    t3_fill = {
        "T3_01_事故经过": TEST_CASE.get('accident_detail', ''),
        "T3_02_责任认定": TEST_CASE.get('responsibility_result', ''),
        "T3_03_保险信息": TEST_CASE.get('insurance_info', ''),
        "T3_04_请求依据": (
            f"根据《民法典》第1179条、第1208条、第1213条，"
            f"《道路交通安全法》第76条，"
            f"被告应赔偿原告各项损失共计{calc_result['total']:.2f}元。"
        ),
        "T3_05_证据清单": "（详见附件证据目录）",
    }
    # 诉前保全信息
    if TEST_CASE.get('pre_litigation_preservation'):
        t3_fill["T3_06_保全法院"] = TEST_CASE.get('preservation_court', '')
        t3_fill["T3_06_保全时间"] = TEST_CASE.get('preservation_time', '')
    engine.fill(t3_fill)
    print(f"  已填充 {len(t3_fill)} 个事实理由字段")

    # --- Step 10: 填充勾选框 ---
    print("\n[Step 10] 填充勾选框...")
    # 使用引擎层的三级匹配方法处理勾选框
    replaced_count = engine.fill_checkboxes_squares(checkbox_map)
    print(f"  已处理 {len(checkbox_map)} 个勾选框，实际替换 {replaced_count} 个■")

    # --- Step 11: 保存输出 ---
    print("\n[Step 11] 保存输出文件...")
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'output')
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, '测试输出_交通事故起诉状.docx')
    output_path = os.path.abspath(output_path)

    engine.save(output_path)
    engine.cleanup()

    file_size = os.path.getsize(output_path)
    print(f"  输出路径: {output_path}")
    print(f"  文件大小: {file_size:,} 字节")

    # --- Step 12: 验证输出 ---
    print("\n[Step 12] 验证输出文件...")
    verify_output(output_path, fill_map, calc_result)

    print("\n" + "=" * 60)
    print("✅ 端到端测试完成！")
    print("=" * 60)

    return output_path


# ============================================================
def verify_output(output_path: str, fill_map: dict, calc_result: dict):
    """验证输出文件的完整性"""
    import zipfile
    import xml.etree.ElementTree as ET

    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    with zipfile.ZipFile(output_path) as z:
        xml_content = z.read('word/document.xml')

    root = ET.fromstring(xml_content)

    # 1. 检查是否还有未填充的书签占位符
    unfilled = []
    for t in root.iter(f'{W}t'):
        if t.text and '{T' in t.text and '}' in t.text:
            unfilled.append(t.text)

    if unfilled:
        print(f"  ⚠️  发现 {len(unfilled)} 个未填充的书签占位符:")
        for u in unfilled[:10]:
            print(f"    - {u}")
    else:
        print(f"  ✅ 所有书签占位符已填充")

    # 2. 检查关键字段是否出现
    all_text = []
    for t in root.iter(f'{W}t'):
        if t.text:
            all_text.append(t.text)
    full_text = ''.join(all_text)

    checks = [
        ("原告姓名", "张大明"),
        ("第二原告", "李小红"),
        ("被告姓名", "王大力"),
        ("保险公司", "平安"),
        ("残疾赔偿金", "95088"),  # 47544×0.1×20
        ("总金额", str(int(calc_result['total']))),
    ]

    for label, expected in checks:
        if expected in full_text:
            print(f"  ✅ {label}: 已出现（{expected}）")
        else:
            print(f"  ❌ {label}: 未找到（期望含'{expected}'）")

    # 3. 检查勾选框
    checked_count = full_text.count('■')
    unchecked_count = full_text.count('□')
    print(f"  📊 勾选框统计: 已勾选■={checked_count}, 未勾选□={unchecked_count}")


if __name__ == '__main__':
    output = main()
    print(f"\n输出文件: {output}")
