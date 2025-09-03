#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基本功能测试
"""

import unittest
import os
import sys

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestBasicFunctionality(unittest.TestCase):
    """基本功能测试类"""
    
    def test_import_main_module(self):
        """测试能否导入主模块"""
        try:
            import excel_image_extractor_gui
            self.assertTrue(True, "成功导入主模块")
        except ImportError as e:
            self.fail(f"导入主模块失败: {e}")
    
    def test_requirements_file_exists(self):
        """测试 requirements.txt 文件是否存在"""
        self.assertTrue(
            os.path.exists("requirements.txt"),
            "requirements.txt 文件应该存在"
        )
    
    def test_main_script_exists(self):
        """测试主脚本文件是否存在"""
        self.assertTrue(
            os.path.exists("excel_image_extractor_gui.py"),
            "主脚本文件应该存在"
        )

if __name__ == '__main__':
    unittest.main() 