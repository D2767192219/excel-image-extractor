#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel图片提取器 - GitHub Actions 构建脚本
专门为 CI/CD 环境优化的构建脚本
"""

import os
import sys
import platform
import subprocess
import shutil
from pathlib import Path

def get_platform_info():
    """获取平台信息"""
    system = platform.system().lower()
    if system == "windows":
        return "windows"
    elif system == "darwin":
        return "macos"
    else:
        return "linux"

def install_dependencies():
    """安装依赖"""
    print("正在安装依赖...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("依赖安装完成")
    except subprocess.CalledProcessError as e:
        print(f"依赖安装失败: {e}")
        sys.exit(1)

def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # 清理之前的构建
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")
    
    # PyInstaller 参数
    platform_name = get_platform_info()
    
    if platform_name == "windows":
        exe_name = "Excel图片提取器"
        spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_image_extractor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{exe_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)
'''
    else:
        exe_name = "excel-image-extractor"
        spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_image_extractor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{exe_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)
'''
    
    # 写入 spec 文件
    with open("build_workflow.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
    
    # 使用 PyInstaller 构建
    try:
        subprocess.run([sys.executable, "-m", "PyInstaller", "build_workflow.spec"], check=True)
        print(f"构建完成: {exe_name}")
    except subprocess.CalledProcessError as e:
        print(f"构建失败: {e}")
        sys.exit(1)

def create_package():
    """创建分发包"""
    print("创建分发包...")
    
    platform_name = get_platform_info()
    python_version = f"{sys.version_info.major}.{sys.version_info.minor}"
    
    if platform_name == "windows":
        package_name = f"Excel图片提取器-Windows-{python_version}.zip"
        exe_path = "dist/Excel图片提取器.exe"
        if not os.path.exists(exe_path):
            print(f"错误: 找不到可执行文件 {exe_path}")
            sys.exit(1)
    else:
        package_name = f"excel-image-extractor-{platform_name.capitalize()}-{python_version}.zip"
        exe_path = "dist/excel-image-extractor"
        if not os.path.exists(exe_path):
            print(f"错误: 找不到可执行文件 {exe_path}")
            sys.exit(1)
    
    # 创建 zip 包
    try:
        if platform_name == "windows":
            # Windows 使用 PowerShell 的 Compress-Archive
            subprocess.run([
                "powershell", "-Command", 
                f"Compress-Archive -Path '{exe_path}' -DestinationPath '{package_name}' -Force"
            ], check=True)
        else:
            # Linux/macOS 使用 zip 命令
            subprocess.run(["zip", "-r", package_name, exe_path], check=True)
        
        print(f"分发包创建完成: {package_name}")
        
        # 输出文件路径供 workflow 使用
        print(f"::set-output name=package_path::{package_name}")
        
    except subprocess.CalledProcessError as e:
        print(f"创建分发包失败: {e}")
        sys.exit(1)

def main():
    """主函数"""
    print("=== Excel图片提取器 - GitHub Actions 构建脚本 ===")
    print(f"平台: {get_platform_info()}")
    print(f"Python版本: {sys.version}")
    
    # 检查必要文件
    required_files = ["excel_image_extractor_gui.py", "requirements.txt"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"错误: 找不到必要文件 {file}")
            sys.exit(1)
    
    # 执行构建流程
    install_dependencies()
    build_executable()
    create_package()
    
    print("构建流程完成!")

if __name__ == "__main__":
    main() 