@echo off
chcp 65001
echo 正在配置打包环境...

:: 检查Python是否已安装
python --version >nul 2>&1
if errorlevel 1 (
    echo Python未安装！
    echo 请先下载并安装Python 3.8或更高版本
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时请勾选"Add Python to PATH"选项
    pause
    exit /b 1
)

:: 安装必要的包
echo 正在安装必要的包...
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl pillow

:: 运行打包脚本
echo 开始打包...
python build.py

:: 检查是否成功
if exist "dist\Excel图片提取器.exe" (
    echo.
    echo 打包成功！
    echo 可执行文件位于：%CD%\dist\Excel图片提取器.exe
    echo 您可以直接运行这个exe文件，不需要Python环境
) else (
    echo 打包失败，请检查错误信息
)

pause 