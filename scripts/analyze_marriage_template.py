# -*- coding: utf-8 -*-
"""
分析离婚纠纷模板的详细结构
提取每个表格每个单元格的段落结构
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# OOXML命名空间
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def analyze_template(docx_path):
    """分析模板结构"""
    print(f"=" * 80)
    print(f"分析模板: {docx_path}")
    print(f"=" * 80)
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        xml_content = zf.read('word/document.xml')
    
    root = ET.fromstring(xml_content)
    
    # 获取所有表格
    tables = list(root.iter(f'{W}tbl'))
    print(f"\n共找到 {len(tables)} 个表格\n")
    
    for table_idx, table in enumerate(tables):
        print(f"{'='*80}")
        print(f"表格 {table_idx}: T{table_idx}")
        print(f"{'='*80}")
        
        rows = list(table.iter(f'{W}tr'))
        for row_idx, row in enumerate(rows):
            cells = list(row.iter(f'{W}tc'))
            print(f"\n--- Row {row_idx} (共 {len(cells)} 个单元格) ---")
            
            for cell_idx, cell in enumerate(cells):
                paragraphs = list(cell.iter(f'{W}p'))
                print(f"\n  Cell {cell_idx}:")
                
                for para_idx, para in enumerate(paragraphs):
                    runs = list(para.iter(f'{W}r'))
                    para_text_parts = []
                    for run in runs:
                        for t_elem in run.iter(f'{W}t'):
                            if t_elem.text:
                                para_text_parts.append(repr(t_elem.text))
                    
                    # 检查是否有书签
                    bookmarks = []
                    for bm in para.iter(f'{W}bookmarkStart'):
                        bm_name = bm.get('name')
                        if bm_name and not bm_name.startswith('_'):
                            bookmarks.append(bm_name)
                    
                    text_display = ''.join([t.strip('"') for t in para_text_parts])[:60]
                    bm_display = f" | 书签名: {bookmarks}" if bookmarks else ""
                    print(f"    Para {para_idx}: [{', '.join(para_text_parts)[:80]}...]{bm_display}")

def analyze_template_detailed(docx_path):
    """更详细的分析：提取每个表格每个单元格每个段落的完整文本"""
    print(f"\n{'#'*80}")
    print(f"# 详细段落结构分析")
    print(f"{'#'*80}")
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        xml_content = zf.read('word/document.xml')
    
    root = ET.fromstring(xml_content)
    tables = list(root.iter(f'{W}tbl'))
    
    # 分析T0表格（当事人信息）
    print("\n\n### 表格0 (T0) - 当事人信息表 ###")
    if len(tables) > 0:
        table = tables[0]
        rows = list(table.iter(f'{W}tr'))
        for row_idx, row in enumerate(rows):
            cells = list(row.iter(f'{W}tc'))
            cell_texts = []
            for cell in cells:
                cell_full_text = []
                for para in cell.iter(f'{W}p'):
                    para_text = []
                    for t in para.iter(f'{W}t'):
                        if t.text:
                            para_text.append(t.text)
                    if para_text:
                        cell_full_text.append(''.join(para_text))
                cell_texts.append(' | '.join(cell_full_text))
            print(f"  Row {row_idx}: {cell_texts}")
    
    # 分析T1表格（被告+诉讼请求）
    print("\n\n### 表格1 (T1) - 被告信息 + 诉讼请求 ###")
    if len(tables) > 1:
        table = tables[1]
        rows = list(table.iter(f'{W}tr'))
        for row_idx, row in enumerate(rows):
            cells = list(row.iter(f'{W}tc'))
            cell_texts = []
            for cell in cells:
                cell_full_text = []
                for para in cell.iter(f'{W}p'):
                    para_text = []
                    for t in para.iter(f'{W}t'):
                        if t.text:
                            para_text.append(t.text)
                    if para_text:
                        cell_full_text.append('§'.join(para_text))
                cell_texts.append(' || '.join(cell_full_text))
            print(f"  Row {row_idx}: {cell_texts}")
    
    # 分析T2表格（事实与理由）
    print("\n\n### 表格2 (T2) - 事实与理由 ###")
    if len(tables) > 2:
        table = tables[2]
        rows = list(table.iter(f'{W}tr'))
        for row_idx, row in enumerate(rows):
            cells = list(row.iter(f'{W}tc'))
            cell_texts = []
            for cell in cells:
                cell_full_text = []
                for para in cell.iter(f'{W}p'):
                    para_text = []
                    for t in para.iter(f'{W}t'):
                        if t.text:
                            para_text.append(t.text)
                    if para_text:
                        cell_full_text.append('§'.join(para_text))
                cell_texts.append(' || '.join(cell_full_text))
            print(f"  Row {row_idx}: {cell_texts}")
    
    # 分析T3表格（请求依据+证据清单+调解意愿）
    print("\n\n### 表格3 (T3) - 请求依据 + 证据清单 + 调解意愿 ###")
    if len(tables) > 3:
        table = tables[3]
        rows = list(table.iter(f'{W}tr'))
        for row_idx, row in enumerate(rows):
            cells = list(row.iter(f'{W}tc'))
            cell_texts = []
            for cell in cells:
                cell_full_text = []
                for para in cell.iter(f'{W}p'):
                    para_text = []
                    for t in para.iter(f'{W}t'):
                        if t.text:
                            para_text.append(t.text)
                    if para_text:
                        cell_full_text.append('§'.join(para_text))
                cell_texts.append(' || '.join(cell_full_text))
            print(f"  Row {row_idx}: {cell_texts}")


def analyze_checkboxes(docx_path):
    """分析模板中的勾选框"""
    print(f"\n{'#'*80}")
    print(f"# 勾选框分析")
    print(f"{'#'*80}")
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        xml_content = zf.read('word/document.xml')
    
    root = ET.fromstring(xml_content)
    
    # 找所有包含□的段落
    checkboxes = []
    for para in root.iter(f'{W}p'):
        para_text = []
        for t in para.iter(f'{W}t'):
            if t.text:
                para_text.append(t.text)
        full_text = ''.join(para_text)
        if '□' in full_text or '☐' in full_text:
            # 找到段落所在表格
            tables = list(root.iter(f'{W}tbl'))
            table_idx = None
            row_idx = None
            for ti, table in enumerate(tables):
                if para in list(table.iter()):
                    rows = list(table.iter(f'{W}tr'))
                    for ri, row in enumerate(rows):
                        if para in list(row.iter()):
                            table_idx = ti
                            row_idx = ri
                            break
                    break
            checkboxes.append({
                'table': table_idx,
                'row': row_idx,
                'text': full_text[:100]
            })
    
    print(f"\n找到 {len(checkboxes)} 处勾选框:")
    for i, cb in enumerate(checkboxes):
        print(f"  {i+1}. T{cb['table']} Row{cb['row']}: {cb['text']}")


if __name__ == '__main__':
    template_path = '/app/data/所有对话/主对话/element-lawsuit-generator/template_cache/民事起诉状-离婚纠纷.docx'
    analyze_template(template_path)
    analyze_template_detailed(template_path)
    analyze_checkboxes(template_path)
