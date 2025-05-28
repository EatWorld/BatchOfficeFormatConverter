@echo off
chcp 65001 >nul
title Office格式批量转换工具 - GUI版本

echo.
echo ==========================================
echo   Office格式批量转换工具 - GUI版本
echo ==========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.x
    echo 下载地址：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo 正在启动GUI界面...
echo.

REM 运行GUI程序
python "%~dp0run_gui.py"

REM 如果程序异常退出，显示错误信息
if errorlevel 1 (
    echo.
    echo 程序运行时发生错误，错误代码：%errorlevel%
    echo.
    echo 可能的解决方案：
    echo 1. 确保已安装pywin32库：pip install pywin32
    echo 2. 确保Microsoft Office已正确安装并激活
    echo 3. 尝试以管理员身份运行此批处理文件
    echo.
    pause
)