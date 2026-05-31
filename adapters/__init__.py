# -*- coding: utf-8 -*-
"""
适配器模块

本模块提供案由适配器基类，用于扩展不同类型案件的处理能力。
"""

from .base import CaseAdapter
from .traffic_accident import TrafficAccidentAdapter
from .labor_dispute import LaborDisputeAdapter

__all__ = ["CaseAdapter", "TrafficAccidentAdapter", "LaborDisputeAdapter"]
