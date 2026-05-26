#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板管理器：下载/缓存/匹配空白模板
支持从GitHub远程下载，本地缓存
"""

import os
import json
import shutil
import urllib.request
import urllib.parse
from typing import Optional, List, Dict


# GitHub仓库配置
GITHUB_REPO = "hugesharks/element-lawsuit-templates"
GITHUB_BRANCH = "main"
GITHUB_RAW_BASE = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}"

# 11个领域分类到模板目录的映射
CATEGORY_DIRS = {
    "01-刑事自诉": "01-刑事自诉",
    "02-婚姻家事": "02-婚姻家事",
    "03-合同纠纷": "03-合同纠纷",
    "04-劳动争议": "04-劳动争议",
    "05-交通事故": "05-交通事故",
    "06-保险纠纷": "06-保险纠纷",
    "07-知识产权": "07-知识产权",
    "08-行政纠纷": "08-行政纠纷",
    "09-国家赔偿": "09-国家赔偿",
    "10-公益诉讼": "10-公益诉讼",
    "11-海商海事": "11-海商海事",
}

# 案由到分类的映射（在运行时从case_keywords.json加载）
CASE_TO_CATEGORY = None


class TemplateManager:
    """模板管理器"""

    def __init__(self, cache_dir: str = None, config_dir: str = None, local_template_dir: str = None):
        """
        Args:
            cache_dir: 远程模板缓存目录
            config_dir: 配置文件目录
            local_template_dir: 本地模板目录（优先使用，不下载）
        """
        base_dir = os.path.dirname(os.path.dirname(__file__))
        
        if cache_dir is None:
            cache_dir = os.path.join(base_dir, 'template_cache')
        if config_dir is None:
            config_dir = os.path.join(base_dir, 'configs')
        
        self.cache_dir = cache_dir
        self.config_dir = config_dir
        self.local_template_dir = local_template_dir
        
        # 确保缓存目录存在
        os.makedirs(self.cache_dir, exist_ok=True)
        
        # 加载案由分类映射
        self._load_case_category_mapping()
    
    def _load_case_category_mapping(self):
        """从配置文件加载案由到分类的映射"""
        global CASE_TO_CATEGORY
        keywords_path = os.path.join(self.config_dir, 'case_keywords.json')
        if os.path.exists(keywords_path):
            with open(keywords_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            CASE_TO_CATEGORY = {
                name: info.get('category', '')
                for name, info in config.get('case_keywords', {}).items()
            }
        else:
            CASE_TO_CATEGORY = {}

    def get_template(self, case_type: str, doc_type: str, use_local: bool = True) -> Optional[str]:
        """
        获取模板文件路径（优先本地，其次缓存，最后远程下载）
        
        Args:
            case_type: 案由（如"民间借贷纠纷"）
            doc_type: 文书类型（如"民事起诉状"）
            use_local: 是否优先使用本地模板目录
            
        Returns:
            模板文件路径，失败返回None
        """
        filename = f"{doc_type}-{case_type}.docx"
        
        # 1. 优先使用本地模板目录
        if use_local and self.local_template_dir and os.path.isdir(self.local_template_dir):
            local_path = os.path.join(self.local_template_dir, filename)
            if os.path.exists(local_path):
                return local_path
        
        # 2. 检查缓存
        cached_path = os.path.join(self.cache_dir, filename)
        if os.path.exists(cached_path):
            return cached_path
        
        # 3. 远程下载
        category = CASE_TO_CATEGORY.get(case_type, '') if CASE_TO_CATEGORY else ''
        category_dir = CATEGORY_DIRS.get(category, '')
        
        if category_dir:
            success = self._download_from_github(category_dir, filename)
            if success and os.path.exists(cached_path):
                return cached_path
        
        # 4. 尝试不带分类目录直接下载（兼容旧结构）
        # 留作备用
        
        return None

    def _download_from_github(self, category_dir: str, filename: str) -> bool:
        """
        从GitHub下载模板
        
        Args:
            category_dir: 分类目录名（如"03-合同纠纷"）
            filename: 文件名（如"民事起诉状-民间借贷纠纷.docx"）
        """
        # URL编码中文路径
        encoded_category = urllib.parse.quote(category_dir)
        encoded_filename = urllib.parse.quote(filename)
        url = f"{GITHUB_RAW_BASE}/templates/{encoded_category}/{encoded_filename}"
        
        cached_path = os.path.join(self.cache_dir, filename)
        temp_path = cached_path + '.tmp'
        
        try:
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                'Accept': 'application/octet-stream'
            })
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()
            
            # 验证是否为有效的docx文件（ZIP格式以PK开头）
            if len(data) < 4 or data[:2] != b'PK':
                return False
            
            with open(temp_path, 'wb') as f:
                f.write(data)
            
            # 原子操作：先写临时文件，再重命名
            if os.path.exists(cached_path):
                os.remove(cached_path)
            os.rename(temp_path, cached_path)
            
            return True
            
        except Exception as e:
            # 清理临时文件
            if os.path.exists(temp_path):
                os.remove(temp_path)
            print(f"下载模板失败: {filename}, 错误: {e}")
            return False

    def list_local_templates(self) -> List[str]:
        """列出本地所有可用模板"""
        templates = set()
        
        # 本地模板目录
        if self.local_template_dir and os.path.isdir(self.local_template_dir):
            for f in os.listdir(self.local_template_dir):
                if f.endswith('.docx'):
                    templates.add(f)
        
        # 缓存目录
        if os.path.isdir(self.cache_dir):
            for f in os.listdir(self.cache_dir):
                if f.endswith('.docx'):
                    templates.add(f)
        
        return sorted(templates)

    def get_template_info(self, filename: str) -> Dict:
        """解析模板文件名获取信息"""
        name = filename.replace('.docx', '')
        parts = name.split('-', 1)
        return {
            'doc_type': parts[0] if parts else '',
            'case_type': parts[1] if len(parts) > 1 else '',
            'filename': filename,
            'category': CASE_TO_CATEGORY.get(parts[1] if len(parts) > 1 else '', '') if CASE_TO_CATEGORY else ''
        }

    def clear_cache(self):
        """清空远程模板缓存"""
        if os.path.isdir(self.cache_dir):
            shutil.rmtree(self.cache_dir)
            os.makedirs(self.cache_dir, exist_ok=True)

    def pre_download_all(self, local_dir: str = None) -> Dict:
        """
        预下载所有模板到缓存
        
        Args:
            local_dir: 如果提供，从本地目录复制而不是下载
            
        Returns:
            {"success": int, "failed": int, "errors": [str]}
        """
        result = {"success": 0, "failed": 0, "errors": []}
        
        if local_dir and os.path.isdir(local_dir):
            # 从本地目录复制
            for f in os.listdir(local_dir):
                if f.endswith('.docx') and not f.endswith('-实例.docx'):
                    src = os.path.join(local_dir, f)
                    dst = os.path.join(self.cache_dir, f)
                    try:
                        shutil.copy2(src, dst)
                        result["success"] += 1
                    except Exception as e:
                        result["failed"] += 1
                        result["errors"].append(f"复制 {f} 失败: {e}")
        else:
            # 从GitHub下载所有分类下的模板
            for category_dir in CATEGORY_DIRS.values():
                try:
                    encoded_category = urllib.parse.quote(category_dir)
                    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/templates/{encoded_category}"
                    req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=15) as resp:
                        import json
                        files = json.loads(resp.read().decode('utf-8'))
                    
                    for file_info in files:
                        if file_info['name'].endswith('.docx'):
                            success = self._download_from_github(category_dir, file_info['name'])
                            if success:
                                result["success"] += 1
                            else:
                                result["failed"] += 1
                                result["errors"].append(f"下载 {file_info['name']} 失败")
                except Exception as e:
                    result["errors"].append(f"列出 {category_dir} 目录失败: {e}")
        
        return result


# 单独测试
if __name__ == '__main__':
    # 使用本地模板目录测试
    tm = TemplateManager(local_template_dir='/tmp/templates-blank')
    
    # 测试获取模板
    path = tm.get_template("民间借贷纠纷", "民事起诉状")
    print(f"模板路径: {path}")
    
    # 测试列出模板
    templates = tm.list_local_templates()
    print(f"可用模板数量: {len(templates)}")
    for t in templates[:5]:
        print(f"  {t}")
    
    # 测试预下载（从本地）
    result = tm.pre_download_all('/tmp/templates-blank')
    print(f"预下载结果: {result}")
