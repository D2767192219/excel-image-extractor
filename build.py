#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 将Excel图片提取器打包成完全独立的可执行文件
使用方法：python build.py
"""

import os
import sys
import platform
import subprocess
import shutil

def clean_build_dirs():
    """清理构建目录"""
    print("清理旧的构建文件...")
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已删除 {dir_name} 目录")

def build_app():
    """构建独立的可执行程序"""
    try:
        system = platform.system()
        app_name = "Excel图片提取器"
        
        # 清理旧的构建文件
        clean_build_dirs()
        
        print(f"开始构建独立的可执行程序 {app_name}...")
        
        # 基本的 PyInstaller 命令
        cmd = [
            'pyinstaller',
            '--noconfirm',
            '--clean',
            '--windowed',  # 使用 GUI 模式
            '--onefile',   # 打包成单文件
            '--noupx',     # 不使用 UPX 压缩，提高兼容性
            f'--name={app_name}',
            # 添加所有必要的依赖
            '--hidden-import=tkinter',
            '--hidden-import=tkinter.filedialog',
            '--hidden-import=PIL._tkinter_finder',
            '--hidden-import=openpyxl',
            '--collect-submodules=openpyxl',
            '--collect-submodules=PIL',
            # 添加运行时钩子
            '--runtime-hook=runtime_hook.py',
        ]
        
        # 创建运行时钩子文件
        with open('runtime_hook.py', 'w', encoding='utf-8') as f:
            f.write("""
import os
import sys

# 确保能找到打包的数据文件
if getattr(sys, 'frozen', False):
    # 如果是打包后的程序
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)  # 切换到程序所在目录
""")
        
        # 根据操作系统添加特定选项
        if system == 'Darwin':  # macOS
            cmd.extend([
                '--add-binary', f'/System/Library/Frameworks/Tk.framework/Tk{os.pathsep}tk',
                '--add-binary', f'/System/Library/Frameworks/Tcl.framework/Tcl{os.pathsep}tcl',
            ])
            if os.path.exists('app.icns'):
                cmd.extend(['--icon=app.icns'])
        elif system == 'Windows':
            if os.path.exists('app.ico'):
                cmd.extend(['--icon=app.ico'])
            # Windows 特定的打包选项
            cmd.extend([
                '--version-file=version.txt',  # 添加版本信息
                '--uac-admin',  # 请求管理员权限
            ])
            
            # 创建版本信息文件
            with open('version.txt', 'w', encoding='utf-8') as f:
                f.write("""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        u'040904B0',
        [StringStruct(u'FileDescription', u'Excel图片提取器'),
         StringStruct(u'ProductName', u'Excel图片提取器'),
         StringStruct(u'ProductVersion', u'1.0.0'),
         StringStruct(u'CompanyName', u''),
         StringStruct(u'LegalCopyright', u'Copyright (C) 2024')])
    ]),
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)""")
        
        # 添加主程序
        cmd.append('excel_image_extractor_gui.py')
        
        # 执行打包命令
        print("正在执行打包命令...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            print("打包失败！")
            print("错误信息：")
            print(result.stderr)
            sys.exit(1)
        
        # 检查是否成功生成了可执行文件
        exe_path = os.path.join('dist', app_name)
        if system == 'Windows':
            exe_path += '.exe'
        
        if os.path.exists(exe_path):
            print("\n打包成功！")
            print(f"可执行文件位于：{os.path.abspath(exe_path)}")
            print("\n注意事项：")
            print("1. 这是一个完全独立的程序，不需要安装Python或其他依赖")
            print("2. 直接双击即可运行")
            if system == 'Windows':
                print("3. 如果Windows报毒，这是误报，可以添加到白名单")
            elif system == 'Darwin':
                print("3. macOS首次运行时，请右键点击程序，选择\"打开\"")
        else:
            print("\n警告：打包可能未完全成功，未找到可执行文件")
            
        # 清理临时文件
        if os.path.exists('runtime_hook.py'):
            os.remove('runtime_hook.py')
        if os.path.exists('version.txt'):
            os.remove('version.txt')
            
    except Exception as e:
        print(f"构建过程中出现错误：{e}")
        sys.exit(1)

if __name__ == '__main__':
    build_app() 