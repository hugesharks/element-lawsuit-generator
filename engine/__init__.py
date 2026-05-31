# -*- coding: utf-8 -*-
"""
要素式法律文书引擎核心模块

本模块提供以下核心能力：
- BookmarkEngine: OOXML书签定位填充引擎
- EvidenceEngine: 证据清单生成引擎
- ReportEngine: 赔偿报告生成引擎
- 河北省交通事故赔偿标准数据
"""

from .template_engine import BookmarkEngine
from .evidence_engine import EvidenceEngine, EvidenceRule
from .report_engine import ReportEngine
from .data_standards import HEBEI_STANDARD

__all__ = [
    "BookmarkEngine",
    "EvidenceEngine", 
    "EvidenceRule",
    "ReportEngine",
    "HEBEI_STANDARD",
]
