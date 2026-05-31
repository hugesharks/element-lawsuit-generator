# -*- coding: utf-8 -*-
"""
OOXML书签定位填充引擎

核心功能：
- 解包docx到临时目录，直接操作XML
- 定位书签(bookmarkStart/bookmarkEnd)并填充内容
- 处理勾选框(☐→☑)的前序文本匹配替换
- 动态调整表格行数并为新行添加书签
- 重新打包为标准docx格式

技术要点：
- XML命名空间必须预先注册，避免序列化时前缀变异
- 书签属性直接用get('name')，不带命名空间前缀
- 行clone使用deepcopy，保持样式完整
"""

import os
import shutil
import zipfile
import tempfile
import copy
from pathlib import Path
from typing import Optional
import xml.etree.ElementTree as ET

# OOXML命名空间定义
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

# 为每个命名空间注册前缀，确保序列化时保持一致
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

# W命名空间前缀，用于XML标签
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class BookmarkEngine:
    """
    OOXML书签定位填充引擎
    
    工作流程：
    1. 解压docx到临时目录
    2. 解析word/document.xml
    3. 定位书签并填充内容
    4. 重新打包为docx
    
    使用示例：
        engine = BookmarkEngine("模板.docx")
        engine.fill({"原告姓名": "张三", "案号": "(2024)冀01民初123号"})
        engine.save("输出.docx")
        engine.cleanup()
    """
    
    def __init__(self, template_path: str):
        """
        加载模板文件，解包到临时目录
        
        Args:
            template_path: 模板docx文件路径
            
        Raises:
            FileNotFoundError: 模板文件不存在
            zipfile.BadZipFile: 文件不是有效的docx格式
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
        
        self.template_path = template_path
        self._temp_dir = tempfile.mkdtemp(prefix="bookmark_engine_")
        self._document_xml_path = os.path.join(self._temp_dir, "word", "document.xml")
        self._tree: Optional[ET.ElementTree] = None
        self._root: Optional[ET.Element] = None
        
        # 解压docx
        self._extract_docx()
        # 加载document.xml
        self._load_document()
    
    def _extract_docx(self):
        """解压docx文件到临时目录"""
        with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
            zip_ref.extractall(self._temp_dir)
    
    def _load_document(self):
        """加载并解析document.xml"""
        self._tree = ET.parse(self._document_xml_path)
        self._root = self._tree.getroot()
    
    def list_bookmarks(self) -> list[str]:
        """
        列出模板中所有书签名称
        
        Returns:
            书签名称列表，按出现顺序排列
        """
        bookmarks = []
        for bm_start in self._root.iter(f'{W}bookmarkStart'):
            name = bm_start.get('name')
            if name and not name.startswith('_'):
                # 过滤掉Word自动生成的书签（通常以下划线开头）
                bookmarks.append(name)
        return bookmarks
    
    def fill(self, fill_map: dict[str, str]) -> None:
        """
        按书签名填充内容
        
        Args:
            fill_map: 书签名称到内容的映射
                示例: {"T0_01_原告姓名": "张某某", "T0_02_案号": "(2024)冀01民初123号"}
        """
        for bookmark_name, content in fill_map.items():
            self._fill_bookmark(bookmark_name, str(content))
    
    def _fill_bookmark(self, bookmark_name: str, content: str):
        """
        填充单个书签的内容
        
        查找书签的bookmarkStart和bookmarkEnd，
        删除之间的所有内容，在bookmarkStart后插入新内容
        
        Args:
            bookmark_name: 书签名称
            content: 要填充的内容
        """
        # 查找书签范围
        result = self._find_bookmark_range(bookmark_name)
        if result is None:
            return  # 书签不存在，跳过
        
        bm_start, bm_end, parent = result
        
        # 收集bookmarkStart和bookmarkEnd之间的所有元素
        children = list(parent)
        start_idx = children.index(bm_start)
        end_idx = children.index(bm_end)
        
        # 删除书签范围内的所有内容节点（保留bookmarkStart和bookmarkEnd）
        elements_to_remove = []
        for i in range(start_idx + 1, end_idx):
            if children[i] != bm_start and children[i] != bm_end:
                elements_to_remove.append(children[i])
        
        for elem in elements_to_remove:
            parent.remove(elem)
        
        # 创建新的run元素填充内容
        new_run = self._create_run_with_text(content)
        # 在bookmarkEnd之前插入
        insert_idx = list(parent).index(bm_end)
        parent.insert(insert_idx, new_run)
    
    def _find_bookmark_range(self, bookmark_name: str):
        """
        找到书签的bookmarkStart和bookmarkEnd
        
        Args:
            bookmark_name: 书签名称
            
        Returns:
            (bookmarkStart元素, bookmarkEnd元素, 父元素) 或 None
        """
        for bm_start in self._root.iter(f'{W}bookmarkStart'):
            # 关键：XML属性不继承默认命名空间，直接用get('name')
            if bm_start.get('name') == bookmark_name:
                bm_id = bm_start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                parent = None
                
                # 找到父元素
                for p in self._root.iter():
                    if bm_start in list(p):
                        parent = p
                        break
                
                if parent is None:
                    continue
                
                # 找对应的bookmarkEnd
                for bm_end in parent.iter(f'{W}bookmarkEnd'):
                    if bm_end.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id') == bm_id:
                        return (bm_start, bm_end, parent)
        
        return None
    
    def _create_run_with_text(self, text: str) -> ET.Element:
        """
        创建包含文本的run元素
        
        Args:
            text: 文本内容
            
        Returns:
            w:r元素，包含w:t子元素
        """
        run = ET.SubElement(ET.Element('dummy'), f'{W}r')
        run = ET.Element(f'{W}r')
        text_elem = ET.SubElement(run, f'{W}t')
        text_elem.text = text
        # 保留空格
        text_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return run
    
    def fill_checkbox(self, checkbox_map: dict[str, bool]) -> None:
        """
        填充勾选框
        
        Args:
            checkbox_map: 勾选框映射
                key为前序文本（如"性别_男"），value为是否勾选
                示例: {"性别_男": True, "性别_女": False, "是否上诉_是": True}
        """
        # 遍历所有段落
        for paragraph in self._root.iter(f'{W}p'):
            self._fill_checkbox_in_paragraph(paragraph, checkbox_map)
    
    def _fill_checkbox_in_paragraph(self, paragraph: ET.Element, checkbox_map: dict[str, bool]):
        """
        在单个段落中处理☐→☑替换
        
        逻辑：
        1. 收集段落内所有<w:t>文本节点
        2. 拼接完整文本，检查是否含☐
        3. 对每个☐节点，收集其前面的所有文本作为前序文本
        4. 根据前序文本匹配checkbox_map的key
        5. 匹配到且值为True → 替换为☑
        
        Args:
            paragraph: w:p段落元素
            checkbox_map: 勾选框映射
        """
        # 收集段落内所有文本节点及其run元素
        text_nodes = []
        for run in paragraph.iter(f'{W}r'):
            for t_elem in run.iter(f'{W}t'):
                if t_elem.text:
                    text_nodes.append((t_elem, t_elem.text))
        
        if not text_nodes:
            return
        
        # 检查是否有勾选框
        full_text = ''.join(text for _, text in text_nodes)
        if '☐' not in full_text:
            return
        
        # 对每个包含☐的节点处理
        for t_elem, text in text_nodes:
            if '☐' not in text:
                continue
            
            # 收集该节点之前的所有文本
            preceding_text = ""
            for prev_elem, prev_text in text_nodes:
                if prev_elem is t_elem:
                    break
                preceding_text += prev_text
            
            # 处理当前节点中的☐
            new_text = text
            for key, should_check in checkbox_map.items():
                # 检查前序文本是否匹配key的前缀
                # key格式如 "性别_男"，前序文本可能包含"性别："
                key_prefix = key.split('_')[0] if '_' in key else key
                key_suffix = key.split('_')[1] if '_' in key else ""
                
                if key_prefix in preceding_text or key in preceding_text:
                    if key_suffix and key_suffix in text and '☐' in text:
                        # 找到匹配的勾选框
                        if should_check:
                            new_text = new_text.replace('☐', '☑', 1)
                        break
            
            t_elem.text = new_text
    
    def adjust_table_rows(self, table_index: int, target_count: int,
                          template_row_index: int = 0) -> list[str]:
        """
        调整表格行数
        
        Args:
            table_index: 第几个表格（从0开始）
            target_count: 目标行数
            template_row_index: 用哪一行作为模板行（默认第0行）
            
        Returns:
            新增行的书签前缀列表，供后续打书签用
            示例: ["T3_01", "T3_02", "T3_03"]
            
        Raises:
            IndexError: 表格索引超出范围
        """
        tables = list(self._root.iter(f'{W}tbl'))
        if table_index >= len(tables):
            raise IndexError(f"表格索引超出范围: {table_index}，共有{len(tables)}个表格")
        
        table = tables[table_index]
        rows = list(table.iter(f'{W}tr'))
        current_count = len(rows)
        
        new_prefixes = []
        
        if current_count == target_count:
            # 行数相同，不需要调整
            return new_prefixes
        
        elif current_count > target_count:
            # 删除多余的行
            rows_to_remove = rows[target_count:]
            for row in rows_to_remove:
                table.remove(row)
        
        else:
            # 需要添加行
            template_row = rows[template_row_index]
            
            for i in range(current_count, target_count):
                # 深拷贝模板行
                new_row = self._clone_table_row(template_row)
                # 生成书签前缀
                prefix = f"T{table_index}_{i+1:02d}"
                new_prefixes.append(prefix)
                # 插入到表格末尾
                table.append(new_row)
        
        return new_prefixes
    
    def _clone_table_row(self, template_row: ET.Element) -> ET.Element:
        """
        深拷贝表格行
        
        使用copy.deepcopy确保样式属性完整保留，
        清空文本内容但保持结构
        
        Args:
            template_row: 模板行元素
            
        Returns:
            新的行元素
        """
        new_row = copy.deepcopy(template_row)
        
        # 清空文本内容，保持结构
        for t_elem in new_row.iter(f'{W}t'):
            t_elem.text = ''
        
        # 移除原有的书签（如果有）
        for bm_start in list(new_row.iter(f'{W}bookmarkStart')):
            parent = self._find_parent(new_row, bm_start)
            if parent is not None:
                parent.remove(bm_start)
        
        for bm_end in list(new_row.iter(f'{W}bookmarkEnd')):
            parent = self._find_parent(new_row, bm_end)
            if parent is not None:
                parent.remove(bm_end)
        
        return new_row
    
    def _find_parent(self, root: ET.Element, child: ET.Element) -> Optional[ET.Element]:
        """在指定根元素下查找子元素的父元素"""
        for parent in root.iter():
            if child in list(parent):
                return parent
        return None
    
    def add_bookmarks_to_row(self, table_index: int, row_index: int,
                              bookmark_prefix: str, field_names: list[str]) -> None:
        """
        给指定行的各个单元格打书签
        
        Args:
            table_index: 表格索引
            row_index: 行索引
            bookmark_prefix: 书签前缀，如 "T3_01"
            field_names: 字段名列表，如 ["原告姓名", "原告性别", "原告出生日期"]
            
        示例:
            engine.add_bookmarks_to_row(3, 0, "T3_01", ["姓名", "性别", "年龄"])
            # 生成书签：T3_01_姓名, T3_01_性别, T3_01_年龄
        """
        tables = list(self._root.iter(f'{W}tbl'))
        if table_index >= len(tables):
            raise IndexError(f"表格索引超出范围: {table_index}")
        
        table = tables[table_index]
        rows = list(table.iter(f'{W}tr'))
        if row_index >= len(rows):
            raise IndexError(f"行索引超出范围: {row_index}")
        
        row = rows[row_index]
        self._insert_bookmarks_to_cloned_row(row, bookmark_prefix, field_names)
    
    def _insert_bookmarks_to_cloned_row(self, row: ET.Element, prefix: str,
                                         field_names: list[str]):
        """
        给clone出来的行打书签
        
        关键：先清空单元格中已有的文本内容（如模板标签"1. 医疗费"），
        避免填充后出现"医疗费1. 医疗费"的叠加问题。
        保留run元素和样式属性，只清空text内容。
        
        Args:
            row: 行元素
            prefix: 书签前缀
            field_names: 字段名列表
        """
        cells = list(row.iter(f'{W}tc'))
        
        for i, field_name in enumerate(field_names):
            if i >= len(cells):
                break
            
            cell = cells[i]
            paragraphs = list(cell.iter(f'{W}p'))
            if not paragraphs:
                # 没有段落，创建一个
                paragraph = ET.SubElement(cell, f'{W}p')
            else:
                paragraph = paragraphs[0]
            
            # 清空第一个段落所有run的文本（保留run元素和样式属性）
            # 原因：模板行可能含标签文本（如"3. 营养费"），不清空会导致
            # 填充后出现"营养费3. 营养费"叠加。
            # 注意：只清第一个段落！多段落单元格（如T0/T1当事人信息）的
            # 后续段落含其他字段，由 clear_cell_extra_paragraphs 单独处理。
            for run in paragraph.iter(f'{W}r'):
                for t_elem in run.iter(f'{W}t'):
                    t_elem.text = ''
            
            runs = list(paragraph.iter(f'{W}r'))
            if not runs:
                # 没有run元素，创建一个
                run = ET.SubElement(paragraph, f'{W}r')
                t_elem = ET.SubElement(run, f'{W}t')
                t_elem.text = ''
            else:
                run = runs[0]
                t_elem = list(run.iter(f'{W}t'))
                if not t_elem:
                    t_elem = ET.SubElement(run, f'{W}t')
                else:
                    t_elem = t_elem[0]
            
            # 书签名称
            bookmark_name = f"{prefix}_{field_name}"
            
            # 生成书签ID（需要唯一）
            bookmark_id = abs(hash(bookmark_name)) % 100000
            
            # 创建bookmarkStart
            bm_start = ET.Element(f'{W}bookmarkStart')
            bm_start.set(f'{{{NAMESPACES["w"]}}}id', str(bookmark_id))
            bm_start.set('name', bookmark_name)
            
            # 创建bookmarkEnd
            bm_end = ET.Element(f'{W}bookmarkEnd')
            bm_end.set(f'{{{NAMESPACES["w"]}}}id', str(bookmark_id))
            
            # 插入位置
            run_index = list(paragraph).index(run)
            paragraph.insert(run_index, bm_start)
            paragraph.insert(run_index + 2, bm_end)
            
            # 设置占位符文本
            t_elem.text = f'{{{bookmark_name}}}'
    
    def clear_cell_extra_paragraphs(self, table_index: int, row_index: int,
                                     cell_index: int) -> None:
        """
        清空指定单元格中除第一个段落以外的所有段落文本。
        
        用于清理模板标签重叠问题。典型场景：
        T2赔偿项表格的Cell 0中，Para 0为空（用于放书签），
        Para 1含模板标签"1. 医疗费"。填充书签后变成"医疗费1. 医疗费"。
        调用此方法后Para 1被清空，只剩书签内容"医疗费"。
        
        ⚠️ 不要用于T0/T1当事人表格，那些单元格的多段落包含不同字段。
        
        Args:
            table_index: 表格索引
            row_index: 行索引
            cell_index: 单元格索引
        """
        tables = list(self._root.iter(f'{W}tbl'))
        if table_index >= len(tables):
            return
        
        rows = list(tables[table_index].iter(f'{W}tr'))
        if row_index >= len(rows):
            return
        
        cells = list(rows[row_index].iter(f'{W}tc'))
        if cell_index >= len(cells):
            return
        
        cell = cells[cell_index]
        paragraphs = list(cell.iter(f'{W}p'))
        
        # 清空第2个及之后的段落
        for para in paragraphs[1:]:
            for run in para.iter(f'{W}r'):
                for t_elem in run.iter(f'{W}t'):
                    t_elem.text = ''

    # ================================================================
    # 勾选框填充方法（□→■ 模式）
    # ================================================================

    def fill_checkboxes_squares(self, checkbox_map: dict[str, bool]) -> int:
        """
        处理模板中的□勾选框（带表上下文 + 行上下文感知）

        模板使用□（WHITE SQUARE U+25A1）作为未勾选状态，
        勾选后替换为■（BLACK SQUARE U+25A0）。

        checkbox_map 的 key 格式约定：
        - 表前缀+选项："T0_性别_男", "T3_诉前保全_是"
          → 精确匹配该表中对应选项文本的□
        - 全局选项："案件类型_纯民事", "了解"
          → 在任何表中匹配选项文本

        匹配策略（三级）：
        1. 表精确匹配：用 "T{idx}" 前缀的key精确匹配
        2. 全局匹配：无前缀key匹配选项文本
        3. 行消歧：同一表内多个相同选项文本（如T3的两个"是□"），
           通过同行Cell 0的关键词消歧

        Args:
            checkbox_map: {勾选框key: True/False}

        Returns:
            实际替换的勾选框数量
        """
        # 预处理：分离表前缀key和全局key
        table_option_map = {}   # {"0": {"男": {"T0_性别_男": True}}, "3": {"是": {...}}}
        global_option_map = {}  # {"纯民事": True, "了解": True, ...}

        for key, should_check in checkbox_map.items():
            # key格式: "T{digit}_{category}_{option}" 如 "T0_性别_男"
            # 或 "T{digit}_{option}" 如 "T4_是"
            if key.startswith('T') and len(key) > 2 and key[1].isdigit() and '_' in key[2:]:
                # 表前缀key：提取 table_idx 和 option_text
                table_idx = key[1]  # digit at index 1
                rest = key[3:]      # skip "T{digit}_"
                # 取最后一段作为选项文本
                sub_parts = rest.split('_')
                option_text = sub_parts[-1]
                
                if table_idx not in table_option_map:
                    table_option_map[table_idx] = {}
                if option_text not in table_option_map[table_idx]:
                    table_option_map[table_idx][option_text] = {}
                table_option_map[table_idx][option_text][key] = should_check
            else:
                # 全局key：取最后一段作为选项文本
                parts = key.split('_')
                option_text = parts[-1]
                # True 优先
                if option_text in global_option_map:
                    global_option_map[option_text] = global_option_map[option_text] or should_check
                else:
                    global_option_map[option_text] = should_check

        # 预先构建表格列表
        tables = list(self._root.iter(f'{W}tbl'))

        # 构建行上下文缓存：{(table_idx, row_idx): cell0_text}
        row_context_cache = {}
        for ti, table in enumerate(tables):
            rows = list(table.iter(f'{W}tr'))
            for ri, row in enumerate(rows):
                cells = list(row.iter(f'{W}tc'))
                if cells:
                    cell0_texts = []
                    for t in cells[0].iter(f'{W}t'):
                        if t.text:
                            cell0_texts.append(t.text)
                    row_context_cache[(ti, ri)] = ''.join(cell0_texts)

        # 遍历所有段落
        replaced_count = 0
        for paragraph in self._root.iter(f'{W}p'):
            text_nodes = []
            for run in paragraph.iter(f'{W}r'):
                for t_elem in run.iter(f'{W}t'):
                    if t_elem.text:
                        text_nodes.append(t_elem)

            if not text_nodes:
                continue

            full_text = ''.join(t.text for t in text_nodes)
            if '□' not in full_text:
                continue

            # 确定表格索引和行索引
            table_idx, row_idx = self._find_table_and_row(
                paragraph, tables, row_context_cache
            )

            # 获取行上下文关键词（来自同Cell 0的文本）
            row_keywords = ''
            if table_idx is not None and row_idx is not None:
                row_keywords = row_context_cache.get((table_idx, row_idx), '')

            for t_elem in text_nodes:
                text = t_elem.text
                if '□' not in text:
                    continue

                new_text = text
                positions = [i for i, c in enumerate(new_text) if c == '□']
                for pos in reversed(positions):
                    before = new_text[:pos].rstrip()
                    option_label = self._extract_option_label(before)

                    if not option_label:
                        continue

                    # 三级匹配
                    should_check = self._match_checkbox(
                        option_label, table_idx, row_keywords,
                        table_option_map, global_option_map
                    )

                    if should_check is True:
                        new_text = new_text[:pos] + '■' + new_text[pos+1:]
                        replaced_count += 1

                if new_text != text:
                    t_elem.text = new_text

        return replaced_count

    def fill_preservation_direct(self, table_index: int, row_index: int,
                                  preservation_court: str,
                                  preservation_time: str) -> None:
        """
        T3 Row 4 诉前保全：直接文本替换（不走书签机制）
        
        模板原文："是□    保全法院：              保全时间：保全案号：否□"
        处理后：  "是■    保全法院：邢台市襄都区人民法院    保全时间：2025年11月5日保全案号：否□"
        
        因为"是□/否□"勾选框和"保全法院/保全时间"字段在同一个段落里，
        打书签会把整个段落清空导致勾选框丢失，所以用直接替换。
        
        Args:
            table_index: 表格索引（T3=3）
            row_index: 行索引（Row 4=4）
            preservation_court: 保全法院名称
            preservation_time: 保全时间
        """
        tables = list(self._root.iter(f'{W}tbl'))
        if table_index >= len(tables):
            return
        
        rows = list(tables[table_index].iter(f'{W}tr'))
        if row_index >= len(rows):
            return
        
        row = rows[row_index]
        cells = list(row.iter(f'{W}tc'))
        
        if len(cells) < 2:
            return
        
        cell1 = cells[1]
        
        # 收集Cell 1所有文本节点
        text_nodes = []
        for run in cell1.iter(f'{W}r'):
            for t_elem in run.iter(f'{W}t'):
                if t_elem.text:
                    text_nodes.append(t_elem)
        
        # 替换逻辑（每个文本节点独立处理）：
        # 1. "是□" → "是■"
        # 2. "              " (纯空格节点，保全法院占位) → 保全法院名
        # 3. "保全时间：" (独立节点) → "保全时间：{时间}"
        for t_elem in text_nodes:
            text = t_elem.text
            
            # 替换"是□"为"是■"
            if '是□' in text:
                t_elem.text = text.replace('是□', '是■')
                continue
            
            # 保全法院：纯空格文本节点 → 替换为法院名
            if preservation_court and text.strip() == '' and len(text) >= 10:
                # 纯空格且长度够长（模板占位通常是14个空格）
                t_elem.text = preservation_court
                continue
            
            # 保全时间：独立文本节点 → 追加时间
            if preservation_time and text.strip() == '保全时间：':
                t_elem.text = f'保全时间：{preservation_time}'
                continue

    def _find_table_and_row(self, paragraph: ET.Element, tables: list,
                             row_context_cache: dict) -> tuple:
        """
        确定段落属于第几个表格的第几行
        
        Args:
            paragraph: 段落元素
            tables: 表格元素列表
            row_context_cache: 行上下文缓存（未使用，保留接口兼容）
            
        Returns:
            (table_idx, row_idx) 或 (None, None)
        """
        for ti, table in enumerate(tables):
            if not self._is_descendant(paragraph, table):
                continue
            # 找到了table，现在找row
            rows = list(table.iter(f'{W}tr'))
            for ri, row in enumerate(rows):
                if self._is_descendant(paragraph, row):
                    return (ti, ri)
            return (ti, None)
        return (None, None)

    def _is_descendant(self, child: ET.Element, potential_ancestor: ET.Element) -> bool:
        """检查child是否是potential_ancestor的后代元素"""
        for elem in potential_ancestor.iter():
            if elem is child:
                return True
        return False

    def _extract_option_label(self, text: str) -> str:
        """
        从□前面的文本中提取选项标签

        模式：
        - "性别：男□" → "男"
        - "是□" → "是"
        - "全责□" → "全责"
        - "暂不确定，想要了解更多内容□" → "暂不确定，想要了解更多内容"
        """
        text = text.rstrip()
        if not text:
            return ''

        # 移除尾部的冒号、空格等
        while text and text[-1] in '：: \t　':
            text = text[:-1]

        if not text:
            return ''

        # 取最后一个"词"——按常见分隔符切分
        # 先尝试取最后2个字（覆盖"全责""了解"等）
        if len(text) >= 2:
            two_char = text[-2:]
            known_two = ['全责', '主责', '同责', '次责', '无责',
                          '了解', '国有', '控股', '参股', '民营',
                          '有□', '无□']  # 有□无□不会到这里，因为□已被剥离
            if two_char in known_two:
                return two_char

        # 取最后一个字
        return text[-1:]

    def _match_checkbox(self, option_label: str, table_idx: int,
                         row_keywords: str, table_option_map: dict,
                         global_option_map: dict) -> bool:
        """
        三级匹配单个勾选框
        
        1. 表精确匹配
        2. 行消歧（同一表内多个同名选项，用行关键词区分）
        3. 全局匹配
        
        Args:
            option_label: 从□前提取的选项文本（如"男""是""否""了解"）
            table_idx: 段落所在表格索引（int或None）
            row_keywords: 同行Cell 0的文本（用于消歧）
            table_option_map: {table_idx_str: {option_text: {full_key: should_check}}}
            global_option_map: {option_text: should_check}
        
        Returns:
            True/False/None
        """
        if table_idx is not None:
            tbl_map = table_option_map.get(str(table_idx), {})
            if option_label in tbl_map:
                entries = tbl_map[option_label]
                # 如果只有一个entry，直接返回
                if len(entries) == 1:
                    return list(entries.values())[0]
                
                # 多个entry（如T3有两个"是"：保全和鉴定）
                # 用行关键词消歧
                for full_key, should_check in entries.items():
                    # full_key 如 "T3_诉前保全_是" 或 "T3_鉴定_是"
                    # 提取中间描述词
                    rest = full_key[2:]  # "3_诉前保全_是"
                    parts = rest.split('_')
                    if len(parts) >= 3:
                        desc = parts[1]  # "诉前保全" 或 "鉴定"
                        # 检查行关键词是否包含描述词或相关词
                        if self._row_matches_desc(row_keywords, desc):
                            return should_check
                
                # 消歧失败，返回None（不勾选，避免误匹配）
                return None

        # 全局匹配
        return global_option_map.get(option_label)

    def _row_matches_desc(self, row_keywords: str, desc: str) -> bool:
        """检查行关键词是否匹配描述词"""
        # 直接匹配
        if desc in row_keywords:
            return True
        # 关联词映射
        desc_map = {
            '诉前保全': ['保全', '诉前'],
            '鉴定': ['鉴定', '申请鉴定'],
            '调解': ['调解', '先行调解'],
        }
        for related in desc_map.get(desc, []):
            if related in row_keywords:
                return True
        return False

    def save(self, output_path: str) -> str:
        """
        保存为docx文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            输出文件的绝对路径
        """
        # 保存修改后的document.xml
        self._tree.write(self._document_xml_path, 
                        xml_declaration=True, 
                        encoding='UTF-8')
        
        # 重新打包为docx
        output_path = os.path.abspath(output_path)
        self._pack_docx(output_path)
        
        return output_path
    
    def _pack_docx(self, output_path: str):
        """
        将临时目录重新打包为docx
        
        Args:
            output_path: 输出文件路径
        """
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(self._temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, self._temp_dir)
                    docx.write(file_path, arcname)
    
    def cleanup(self):
        """清理临时文件"""
        if self._temp_dir and os.path.exists(self._temp_dir):
            shutil.rmtree(self._temp_dir)
            self._temp_dir = None
    
    def __del__(self):
        """析构时自动清理"""
        self.cleanup()
    
    def __enter__(self):
        """支持with语句"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """退出with语句时清理"""
        self.cleanup()
