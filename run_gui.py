#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office格式批量转换工具 - GUI启动器

这是一个简单的启动脚本，用于运行带有图形界面的Office格式转换工具。

使用方法：
1. 确保已安装所需依赖：pip install pywin32
2. 双击运行此文件，或在命令行中执行：python run_gui.py
3. 在图形界面中选择要转换的目录和选项
4. 点击"开始转换"按钮开始处理文件

注意事项：
- 需要安装并激活Microsoft Office
- 建议以管理员权限运行
- 转换前请备份重要文件
"""

import sys
import os

def check_dependencies():
    """检查必要的依赖库"""
    missing_deps = []
    
    try:
        import tkinter
    except ImportError:
        missing_deps.append('tkinter (通常随Python一起安装)')
    
    try:
        import win32com.client
        import pythoncom
        import win32file
        import win32con
        import pywintypes
    except ImportError:
        missing_deps.append('pywin32 (运行: pip install pywin32)')
    
    if missing_deps:
        print("错误：缺少以下依赖库：")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("\n请安装缺少的依赖库后重试。")
        input("按回车键退出...")
        return False
    
    return True

def check_office():
    """检查Microsoft Office是否可用"""
    try:
        import win32com.client
        import pythoncom
        
        pythoncom.CoInitialize()
        
        # 尝试创建Word应用程序
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            print("✓ Microsoft Word 可用")
        except Exception as e:
            print(f"⚠ Microsoft Word 不可用: {e}")
            return False
        
        # 尝试创建Excel应用程序
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
            print("✓ Microsoft Excel 可用")
        except Exception as e:
            print(f"⚠ Microsoft Excel 不可用: {e}")
            return False
            
        pythoncom.CoUninitialize()
        return True
        
    except Exception as e:
        print(f"检查Office时发生错误: {e}")
        return False

def main():
    print("Office格式批量转换工具 - GUI版本")
    print("=" * 40)
    
    # 检查依赖
    print("正在检查依赖库...")
    if not check_dependencies():
        return
    
    print("✓ 所有依赖库已安装")
    
    # 检查Office
    print("\n正在检查Microsoft Office...")
    if not check_office():
        print("\n警告：Microsoft Office可能未正确安装或未激活。")
        print("转换功能可能无法正常工作。")
        response = input("\n是否继续启动GUI？(y/n): ")
        if response.lower() != 'y':
            return
    
    print("\n启动图形界面...")
    
    try:
        # 导入并运行GUI
        from office_converter_gui import main as gui_main
        gui_main()
    except ImportError:
        print("错误：无法找到office_converter_gui.py文件")
        print("请确保office_converter_gui.py文件在同一目录下")
        input("按回车键退出...")
    except Exception as e:
        print(f"启动GUI时发生错误: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"\n程序发生未预期的错误: {e}")
        input("按回车键退出...")