import os
import shutil
import win32com.client
import pythoncom
import win32file
import win32con
import pywintypes
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
import queue
from datetime import datetime

# 现代化主题配色
COLORS = {
    'primary': '#2563eb',      # 蓝色主色调
    'primary_dark': '#1d4ed8', # 深蓝色
    'secondary': '#10b981',    # 绿色
    'danger': '#ef4444',       # 红色
    'warning': '#f59e0b',      # 橙色
    'info': '#3b82f6',         # 信息蓝色
    'background': '#f8fafc',   # 浅灰背景
    'surface': '#ffffff',      # 白色表面
    'text': '#1f2937',         # 深灰文字
    'text_light': '#6b7280',   # 浅灰文字
    'border': '#e5e7eb',       # 边框颜色
    'hover': '#f3f4f6'         # 悬停颜色
}

class OfficeConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Office格式批量转换工具 - 现代版")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 设置现代化样式
        self.setup_styles()
        
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap(default="office.ico")
        except:
            pass
        
        # 创建变量
        self.source_dir = tk.StringVar()
        self.convert_doc = tk.BooleanVar(value=True)
        self.convert_xls = tk.BooleanVar(value=True)
        self.preserve_timestamps = tk.BooleanVar(value=True)
        self.archive_originals = tk.BooleanVar(value=True)
        self.custom_archive_dir = tk.StringVar()
        self.use_custom_archive = tk.BooleanVar(value=False)
        self.language = tk.StringVar(value="中文")
        
        # 初始化队列
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.stats_queue = queue.Queue()
        
        # 转换状态
        self.is_converting = False
        self.total_files = 0
        self.converted_files = 0
        self.skipped_files = 0
        self.error_files = 0
        
        self.create_menu()
        self.create_widgets()
        self.setup_layout()
        
        # 启动日志更新定时器
        self.root.after(100, self.update_log)
        
        # 初始化底部统计显示（在create_widgets之后）
        self.root.after(200, self.init_stats_display)
        
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 帮助菜单（左侧）
        self.help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=self.help_menu)
        self.help_menu.add_command(label="使用帮助", command=self.show_help_wrapper)
        self.help_menu.add_separator()
        self.help_menu.add_command(label="关于", command=self.show_about_wrapper)
        
        # 语言菜单（右侧）
        language_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Language", menu=language_menu)
        language_menu.add_radiobutton(label="中文", variable=self.language, value="中文", command=self.change_language)
        language_menu.add_radiobutton(label="English", variable=self.language, value="English", command=self.change_language)
        
    def show_help(self):
        """显示帮助信息"""
        help_text = """📖 Office格式批量转换工具使用帮助

🔧 基本操作：
1. 点击"浏览"按钮选择包含Office文件的目录
2. 选择需要转换的文件类型（DOC→DOCX、XLS→XLSX）
3. 根据需要选择其他选项：
   • 备份到默认文件夹：转换后将原文件移动到源目录下的"旧格式文件"文件夹
    • 备份到自定义文件夹：转换后将原文件移动到您指定的文件夹
   • 保留原始时间戳：新文件保持与原文件相同的时间信息
4. 点击"开始转换"按钮启动转换过程
5. 可随时点击"停止转换"按钮中止操作

📊 进度监控：
• 底部面板实时显示转换进度和统计信息
• 右侧日志区域显示详细的转换过程
• 转换完成后会弹出统计报告

⚠️ 注意事项：
• 程序会自动跳过受密码保护的文件
• 转换过程中请勿关闭Microsoft Office应用程序
• 建议在转换前备份重要文件
• 程序支持递归处理子目录中的文件"""
        
        self.show_text_dialog("使用帮助", help_text)
        
    def show_about(self):
         """显示关于信息"""
         about_text = """📄 Office格式批量转换工具
 
 版本：2.0
 作者：张鑫鑫
 单位：蚌埠市蚌山区燕山乡人民政府
 
 功能特点：
 • 批量转换DOC到DOCX格式
 • 批量转换XLS到XLSX格式
 • 保留原始文件时间戳
 • 自动备份原始文件
 • 现代化图形界面
 • 实时进度监控
 
 技术支持：
 基于Python + tkinter开发
 使用Microsoft Office COM组件进行文件转换
 
 项目地址：
 https://github.com/EatWorld/BatchOfficeFormatConverter
 
 © 2024 蚌埠市蚌山区燕山乡人民政府"""
         
         self.show_text_dialog("关于", about_text)
    
    def show_help_english(self):
        """显示英文帮助信息"""
        help_text = """📖 Office Format Batch Converter Help

🔧 Basic Operations:
1. Click "Browse" button to select directory containing Office files
2. Choose file types to convert (DOC→DOCX, XLS→XLSX)
3. Select additional options as needed:
   • Backup to Default Folder: Move original files to "Old Format Files" folder in source directory
   • Backup to Custom Folder: Move original files to your specified folder
   • Preserve Original Timestamps: Keep same time information as original files
4. Click "Start Conversion" button to begin conversion process
5. Click "Stop Conversion" button to abort operation at any time

📊 Progress Monitoring:
• Bottom panel shows real-time conversion progress and statistics
• Right log area displays detailed conversion process
• Statistical report will pop up after conversion completes

⚠️ Important Notes:
• Program automatically skips password-protected files
• Do not close Microsoft Office applications during conversion
• Recommend backing up important files before conversion
• Program supports recursive processing of files in subdirectories"""
        
        self.show_text_dialog("Help", help_text)
        
    def show_about_english(self):
        """显示英文关于信息"""
        about_text = """📄 Office Format Batch Converter

Version: 2.0
Author: Zhang Xinxin
Organization: Yanshan Township People's Government, Bengshan District, Bengbu City

Features:
• Batch convert DOC to DOCX format
• Batch convert XLS to XLSX format
• Preserve original file timestamps
• Automatic archiving of original files
• Modern graphical interface
• Real-time progress monitoring

Technical Support:
Developed with Python + tkinter
Uses Microsoft Office COM components for file conversion

Project Repository:
https://github.com/EatWorld/BatchOfficeFormatConverter

© 2024 Yanshan Township People's Government, Bengshan District, Bengbu City"""
        
        self.show_text_dialog("About", about_text)
    
    def show_help_wrapper(self):
        """帮助功能包装器"""
        if self.language.get() == "English":
            self.show_help_english()
        else:
            self.show_help()
    
    def show_about_wrapper(self):
        """关于功能包装器"""
        if self.language.get() == "English":
            self.show_about_english()
        else:
            self.show_about()
    
    def change_language(self):
        """切换语言"""
        selected_lang = self.language.get()
        if selected_lang == "English":
            self.update_interface_language("English")
        else:
            self.update_interface_language("中文")
    
    def update_interface_language(self, lang):
        """更新界面语言"""
        if lang == "English":
            # 更新窗口标题
            self.root.title("Office File Batch Converter")
            
            # 重新创建界面
            self.recreate_interface_english()
            
            # 更新菜单标签
            try:
                menubar = self.root.nametowidget(self.root['menu'])
                menubar.entryconfig(0, label="Help")
                menubar.entryconfig(1, label="Language")
            except:
                pass
        else:
            # 更新窗口标题
            self.root.title("Office文档批量转换工具")
            
            # 重新创建界面
            self.recreate_interface_chinese()
            
            # 更新菜单标签
            try:
                menubar = self.root.nametowidget(self.root['menu'])
                menubar.entryconfig(0, label="帮助")
                menubar.entryconfig(1, label="Language")
            except:
                pass
    
    def recreate_interface_english(self):
        """重新创建英文界面"""
        # 清除现有内容
        for widget in self.left_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()
        for widget in self.bottom_frame.winfo_children():
            widget.destroy()
        
        # 更新窗口标题
        self.root.title("Office File Batch Converter")
        
        # 更新菜单标签
        try:
            menubar = self.root.nametowidget(self.root['menu'])
            menubar.entryconfig("帮助", label="Help")
            
            # 更新帮助菜单项
            self.help_menu.entryconfig("使用帮助", label="User Guide")
            self.help_menu.entryconfig("关于", label="About")
        except:
            pass
        
        # 重新创建英文界面
        self.create_left_panel_english()
        self.create_right_panel_english()
        self.create_bottom_panel_english()
    
    def recreate_interface_chinese(self):
        """重新创建中文界面"""
        # 清除现有内容
        for widget in self.left_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()
        for widget in self.bottom_frame.winfo_children():
            widget.destroy()
        
        # 更新窗口标题
        self.root.title("Office文档批量转换工具")
        
        # 更新菜单标签
        try:
            menubar = self.root.nametowidget(self.root['menu'])
            menubar.entryconfig("Help", label="帮助")
            
            # 更新帮助菜单项
            self.help_menu.entryconfig("User Guide", label="使用帮助")
            self.help_menu.entryconfig("About", label="关于")
        except:
            pass
        
        # 重新创建中文界面
        self.create_left_panel()
        self.create_right_panel()
        self.create_bottom_panel()
         
    def show_text_dialog(self, title, text):
        """显示可选择文本的对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x500")
        dialog.configure(bg=COLORS['background'])
        dialog.resizable(True, True)
        
        # 设置对话框居中
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 创建主框架
        main_frame = tk.Frame(dialog, bg=COLORS['background'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 创建文本框和滚动条
        text_frame = tk.Frame(main_frame, bg=COLORS['background'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文本控件
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Microsoft YaHei", 10),
            bg=COLORS['surface'],
            fg=COLORS['text'],
            relief='solid',
            bd=1,
            padx=15,
            pady=15,
            selectbackground=COLORS['primary'],
            selectforeground='white',
            state=tk.DISABLED,
            cursor="arrow"
        )
        
        # 滚动条
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 插入文本
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, text)
        text_widget.config(state=tk.DISABLED)  # 设置为只读但可选择
        
        # 按钮框架
        button_frame = tk.Frame(main_frame, bg=COLORS['background'])
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        # 复制按钮
        copy_button = tk.Button(
            button_frame,
            text="📋 复制全部",
            font=("Microsoft YaHei", 9),
            bg=COLORS['primary'],
            fg='white',
            relief='flat',
            padx=20,
            pady=8,
            cursor='hand2',
            command=lambda: self.copy_text_to_clipboard(text_widget)
        )
        copy_button.pack(side=tk.LEFT)
        
        # 关闭按钮
        close_button = tk.Button(
            button_frame,
            text="关闭",
            font=("Microsoft YaHei", 9),
            bg=COLORS['secondary'],
            fg=COLORS['text'],
            relief='flat',
            padx=20,
            pady=8,
            cursor='hand2',
            command=dialog.destroy
        )
        close_button.pack(side=tk.RIGHT)
        
        # 居中显示对话框
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # 设置焦点
        text_widget.focus_set()
        
    def copy_text_to_clipboard(self, text_widget):
        """复制文本到剪贴板"""
        try:
            # 获取所有文本
            text = text_widget.get(1.0, tk.END).strip()
            # 复制到剪贴板
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            # 显示提示
            messagebox.showinfo("复制成功", "文本已复制到剪贴板")
        except Exception as e:
            messagebox.showerror("复制失败", f"复制文本时出错：{str(e)}")
            
    def on_custom_archive_change(self, *args):
        """处理自定义归档选项变化"""
        if self.use_custom_archive.get():
            # 启用自定义文件夹选择
            self.custom_dir_entry.config(state='normal')
            self.custom_dir_button.config(state='normal', bg=COLORS['secondary'])
            # 禁用默认归档选项
            self.archive_originals.set(False)
        else:
            # 禁用自定义文件夹选择
            self.custom_dir_entry.config(state='disabled')
            self.custom_dir_button.config(state='disabled', bg=COLORS['border'])
            
    def on_default_archive_change(self, *args):
        """处理默认归档选项变化"""
        if self.archive_originals.get():
            # 禁用自定义归档选项
            self.use_custom_archive.set(False)
            
    def select_custom_archive_dir(self):
        """选择自定义备份文件夹"""
        directory = filedialog.askdirectory(
            title="选择备份文件夹",
            initialdir=self.source_dir.get() if self.source_dir.get() else os.getcwd()
        )
        if directory:
            self.custom_archive_dir.set(directory)
        
    def init_stats_display(self):
        """初始化统计显示"""
        initial_stats = "📈 可转换文件: 0 | ✅ 已转换: 0 | ⏭️ 跳过: 0 | ❌ 错误: 0"
        self.stats_queue.put(initial_stats)
        
    def setup_styles(self):
        """设置现代化样式主题"""
        style = ttk.Style()
        
        # 配置主题
        style.theme_use('clam')
        
        # 自定义样式
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 16, 'bold'),
                       foreground=COLORS['text'])
        
        style.configure('Heading.TLabel',
                       font=('Segoe UI', 12, 'bold'),
                       foreground=COLORS['text'])
        
        style.configure('Modern.TButton',
                       font=('Segoe UI', 10),
                       padding=(20, 10))
        
        style.configure('Primary.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       padding=(20, 12))
        
        style.configure('Success.TButton',
                       font=('Segoe UI', 10),
                       padding=(15, 8))
        
        style.configure('Modern.TCheckbutton',
                       font=('Segoe UI', 10),
                       foreground=COLORS['text'])
        
        style.configure('Stats.TLabel',
                       font=('Segoe UI', 9),
                       foreground=COLORS['text_light'])
        
        # 设置根窗口背景
        self.root.configure(bg=COLORS['background'])
        
    def create_widgets(self):
        """创建界面组件"""
        # 设置窗口样式 - 1200x750像素（增加宽度以容纳更多按钮）
        self.root.title("Office格式批量转换工具")
        self.root.geometry("1200x750")
        self.root.configure(bg=COLORS['background'])
        self.root.resizable(False, False)  # 固定窗口大小
        
        # 创建标题区域
        title_frame = tk.Frame(self.root, bg=COLORS['background'], height=50)
        title_frame.pack(fill=tk.X, pady=(8, 0))
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📄 Office 文档批量转换工具",
            font=("Microsoft YaHei", 16, "bold"),
            fg=COLORS['primary'],
            bg=COLORS['background']
        )
        title_label.pack(expand=True)
        
        # 创建主内容区域（调整高度：750-50标题-150底部进度=550像素）
        main_frame = tk.Frame(self.root, bg=COLORS['background'], height=550)
        main_frame.pack(fill=tk.X, padx=10, pady=(8, 0))
        main_frame.pack_propagate(False)
        
        # 左侧控制面板 - 590像素宽（增加宽度）
        self.left_frame = tk.Frame(main_frame, bg=COLORS['surface'], width=590, relief='solid', bd=1)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        self.left_frame.pack_propagate(False)
        
        # 右侧日志面板 - 590像素宽（增加宽度）
        self.right_frame = tk.Frame(main_frame, bg=COLORS['surface'], width=590, relief='solid', bd=1)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        self.right_frame.pack_propagate(False)
        
        # 底部进度面板 - 增加高度到150像素，确保统计信息可见
        self.bottom_frame = tk.Frame(self.root, bg=COLORS['surface'], height=150, relief='solid', bd=1)
        self.bottom_frame.pack(fill=tk.X, padx=10, pady=(8, 8))
        self.bottom_frame.pack_propagate(False)
        
        # 创建各个区域内容
        self.create_left_panel()
        self.create_right_panel()
        self.create_bottom_panel()
        
    def create_left_panel(self):
        # 左侧面板内容区域
        content_frame = tk.Frame(self.left_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 目录选择区域
        dir_label = tk.Label(
            content_frame,
            text="📁 选择转换目录",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        dir_label.pack(anchor="w", pady=(0, 5))
        
        dir_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        dir_frame.pack(fill="x", pady=(0, 15))
        
        self.directory_entry = tk.Entry(
            dir_frame,
            textvariable=self.source_dir,
            font=("Microsoft YaHei", 10),
            bg=COLORS['background'],
            fg=COLORS['text'],
            relief="flat",
            bd=1
        )
        self.directory_entry.pack(side="left", fill="x", expand=True, ipady=8)
        
        browse_btn = tk.Button(
            dir_frame,
            text="浏览",
            command=self.browse_directory,
            font=("Microsoft YaHei", 10),
            bg=COLORS['primary'],
            fg="white",
            relief="flat",
            padx=20
        )
        browse_btn.pack(side="right", padx=(10, 0))
        
        # 转换选项区域
        options_label = tk.Label(
            content_frame,
            text="⚙️ 转换选项",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        options_label.pack(anchor="w", pady=(15, 5))
        
        options_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        options_frame.pack(fill="x", pady=(0, 15))
        
        # 转换类型选项（横向排列）
        convert_frame = tk.Frame(options_frame, bg=COLORS['surface'])
        convert_frame.pack(fill="x", pady=(0, 10))
        
        # 第一行：转换选项
        convert_row1 = tk.Frame(convert_frame, bg=COLORS['surface'])
        convert_row1.pack(fill="x", pady=2)
        
        doc_cb = self.create_modern_checkbox(convert_row1, "转换 DOC → DOCX", self.convert_doc)
        doc_cb.pack(side="left", padx=(0, 20))
        
        xls_cb = self.create_modern_checkbox(convert_row1, "转换 XLS → XLSX", self.convert_xls)
        xls_cb.pack(side="left")
        
        # 第二行：时间戳选项（与第一行对齐）
        convert_row2 = tk.Frame(convert_frame, bg=COLORS['surface'])
        convert_row2.pack(fill="x", pady=2)
        
        timestamp_cb = self.create_modern_checkbox(convert_row2, "保留原始时间戳", self.preserve_timestamps)
        timestamp_cb.pack(side="left", padx=(0, 0))
        
        # 分隔线
        separator = tk.Frame(options_frame, height=1, bg=COLORS['border'])
        separator.pack(fill="x", pady=(5, 10))
        
        # 归档选项标题
        archive_label = tk.Label(
            options_frame,
            text="📁 原文件处理",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        archive_label.pack(anchor="w", pady=(0, 5))
        
        # 归档选项框架（横向排列）
        archive_frame = tk.Frame(options_frame, bg=COLORS['surface'])
        archive_frame.pack(fill="x", pady=(0, 10))
        
        archive_row = tk.Frame(archive_frame, bg=COLORS['surface'])
        archive_row.pack(fill="x", pady=2)
        
        archive_cb = self.create_modern_checkbox(archive_row, "备份到默认文件夹", self.archive_originals)
        archive_cb.pack(side="left", padx=(0, 20))
        
        custom_archive_cb = self.create_modern_checkbox(archive_row, "备份到自定义文件夹", self.use_custom_archive)
        custom_archive_cb.pack(side="left")
        
        # 自定义文件夹选择框架
        custom_dir_frame = tk.Frame(archive_frame, bg=COLORS['surface'])
        custom_dir_frame.pack(fill="x", pady=(5, 0), padx=(20, 0))
        
        self.custom_dir_entry = tk.Entry(
            custom_dir_frame,
            textvariable=self.custom_archive_dir,
            font=("Microsoft YaHei", 9),
            bg=COLORS['background'],
            fg=COLORS['text'],
            relief='solid',
            bd=1,
            state='disabled'
        )
        self.custom_dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.custom_dir_button = tk.Button(
            custom_dir_frame,
            text="📁 选择",
            command=self.select_custom_archive_dir,
            font=("Microsoft YaHei", 9),
            bg=COLORS['border'],
            fg=COLORS['text'],
            relief='solid',
            bd=1,
            state='disabled',
            cursor='hand2'
        )
        self.custom_dir_button.pack(side="right")
        
        # 绑定归档选项变化事件
        self.use_custom_archive.trace('w', self.on_custom_archive_change)
        self.archive_originals.trace('w', self.on_default_archive_change)
        

        
        # 控制面板
        control_label = tk.Label(
            content_frame,
            text="🎮 控制面板",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        control_label.pack(anchor="w", pady=(15, 5))
        
        control_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        control_frame.pack(fill="x", pady=(0, 20))
        
        self.start_button = tk.Button(
            control_frame,
            text="🚀 开始转换",
            command=self.start_conversion,
            font=("Microsoft YaHei", 10),
            bg=COLORS['secondary'],
            fg="white",
            relief="flat",
            width=12,
            height=2
        )
        self.start_button.pack(side="left", padx=(0, 10))
        
        self.stop_button = tk.Button(
            control_frame,
            text="⏹️ 停止转换",
            command=self.stop_conversion,
            font=("Microsoft YaHei", 10),
            bg=COLORS['border'],
            fg=COLORS['text'],
            relief="flat",
            width=12,
            height=2,
            state="disabled"
        )
        self.stop_button.pack(side="left")
    
    def create_left_panel_english(self):
        # Left panel content area
        content_frame = tk.Frame(self.left_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Directory selection area
        dir_label = tk.Label(
            content_frame,
            text="📁 Select Conversion Directory",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        dir_label.pack(anchor="w", pady=(0, 5))
        
        dir_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        dir_frame.pack(fill="x", pady=(0, 15))
        
        self.directory_entry = tk.Entry(
            dir_frame,
            textvariable=self.source_dir,
            font=("Microsoft YaHei", 10),
            bg=COLORS['background'],
            fg=COLORS['text'],
            relief="flat",
            bd=1
        )
        self.directory_entry.pack(side="left", fill="x", expand=True, ipady=8)
        
        browse_btn = tk.Button(
            dir_frame,
            text="Browse",
            command=self.browse_directory,
            font=("Microsoft YaHei", 10),
            bg=COLORS['primary'],
            fg="white",
            relief="flat",
            padx=20
        )
        browse_btn.pack(side="right", padx=(10, 0))
        
        # Conversion options area
        options_label = tk.Label(
            content_frame,
            text="⚙️ Conversion Options",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        options_label.pack(anchor="w", pady=(15, 5))
        
        options_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        options_frame.pack(fill="x", pady=(0, 15))
        
        # Conversion type options (horizontal layout)
        convert_frame = tk.Frame(options_frame, bg=COLORS['surface'])
        convert_frame.pack(fill="x", pady=(0, 10))
        
        # First row: conversion options
        convert_row1 = tk.Frame(convert_frame, bg=COLORS['surface'])
        convert_row1.pack(fill="x", pady=2)
        
        doc_cb = self.create_modern_checkbox(convert_row1, "Convert DOC → DOCX", self.convert_doc)
        doc_cb.pack(side="left", padx=(0, 20))
        
        xls_cb = self.create_modern_checkbox(convert_row1, "Convert XLS → XLSX", self.convert_xls)
        xls_cb.pack(side="left")
        
        # Second row: timestamp option (aligned with first row)
        convert_row2 = tk.Frame(convert_frame, bg=COLORS['surface'])
        convert_row2.pack(fill="x", pady=2)
        
        timestamp_cb = self.create_modern_checkbox(convert_row2, "Preserve Original Timestamps", self.preserve_timestamps)
        timestamp_cb.pack(side="left", padx=(0, 0))
        
        # Separator
        separator = tk.Frame(options_frame, height=1, bg=COLORS['border'])
        separator.pack(fill="x", pady=(5, 10))
        
        # Archive options title
        archive_label = tk.Label(
            options_frame,
            text="📁 Original File Handling",
            font=("Microsoft YaHei", 10, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        archive_label.pack(anchor="w", pady=(0, 5))
        
        # Archive options frame (horizontal layout)
        archive_frame = tk.Frame(options_frame, bg=COLORS['surface'])
        archive_frame.pack(fill="x", pady=(0, 10))
        
        archive_row = tk.Frame(archive_frame, bg=COLORS['surface'])
        archive_row.pack(fill="x", pady=2)
        
        archive_cb = self.create_modern_checkbox(archive_row, "Backup to Default Folder", self.archive_originals)
        archive_cb.pack(side="left", padx=(0, 20))
        
        custom_archive_cb = self.create_modern_checkbox(archive_row, "Backup to Custom Folder", self.use_custom_archive)
        custom_archive_cb.pack(side="left")
        
        # Custom folder selection frame
        custom_dir_frame = tk.Frame(archive_frame, bg=COLORS['surface'])
        custom_dir_frame.pack(fill="x", pady=(5, 0), padx=(20, 0))
        
        self.custom_dir_entry = tk.Entry(
            custom_dir_frame,
            textvariable=self.custom_archive_dir,
            font=("Microsoft YaHei", 9),
            bg=COLORS['background'],
            fg=COLORS['text'],
            relief='solid',
            bd=1,
            state='disabled'
        )
        self.custom_dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.custom_dir_button = tk.Button(
            custom_dir_frame,
            text="📁 Select",
            command=self.select_custom_archive_dir,
            font=("Microsoft YaHei", 9),
            bg=COLORS['border'],
            fg=COLORS['text'],
            relief='solid',
            bd=1,
            state='disabled',
            cursor='hand2'
        )
        self.custom_dir_button.pack(side="right")
        
        # Bind archive option change events
        self.use_custom_archive.trace('w', self.on_custom_archive_change)
        self.archive_originals.trace('w', self.on_default_archive_change)
        
        # Control panel
        control_label = tk.Label(
            content_frame,
            text="🎮 Control Panel",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        control_label.pack(anchor="w", pady=(15, 5))
        
        control_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        control_frame.pack(fill="x", pady=(0, 20))
        
        self.start_button = tk.Button(
            control_frame,
            text="🚀 Start Conversion",
            command=self.start_conversion,
            font=("Microsoft YaHei", 10),
            bg=COLORS['secondary'],
            fg="white",
            relief="flat",
            width=15,
            height=2
        )
        self.start_button.pack(side="left", padx=(0, 10))
        
        self.stop_button = tk.Button(
            control_frame,
            text="⏹️ Stop Conversion",
            command=self.stop_conversion,
            font=("Microsoft YaHei", 10),
            bg=COLORS['border'],
            fg=COLORS['text'],
            relief="flat",
            width=15,
            height=2,
            state="disabled"
        )
        self.stop_button.pack(side="left")
    
    def create_right_panel_english(self):
        # Right panel content area
        content_frame = tk.Frame(self.right_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Log title
        log_label = tk.Label(
            content_frame,
            text="📋 Conversion Log",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        log_label.pack(anchor="w", pady=(0, 5))
        
        # Log text area
        log_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        log_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            font=("Consolas", 9),
            bg="#f8f9fa",
            fg=COLORS['text'],
            relief="flat",
            bd=5
        )
        self.log_text.pack(fill="both", expand=True, pady=(0, 10))
        
        # Log control buttons
        log_control_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        log_control_frame.pack(fill="x", pady=(0, 0))
        
        clear_log_btn = tk.Button(
            log_control_frame,
            text="🗑️Clear Log",
            command=self.clear_log,
            font=("Microsoft YaHei", 10),
            bg=COLORS['warning'],
            fg="white",
            relief="flat",
            width=15,
            height=2,
            compound="left"
        )
        clear_log_btn.pack(side="left", padx=(0, 10))
        
        save_log_btn = tk.Button(
            log_control_frame,
            text="💾Save Log",
            command=self.save_log,
            font=("Microsoft YaHei", 10),
            bg=COLORS['info'],
            fg="white",
            relief="flat",
            width=15,
            height=2,
            compound="left"
        )
        save_log_btn.pack(side="left")
    
    def create_bottom_panel_english(self):
        # Bottom panel content
        content_frame = tk.Frame(self.bottom_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # Progress title
        progress_label = tk.Label(
            content_frame,
            text="📊 Conversion Progress",
            font=("Microsoft YaHei", 10, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        progress_label.pack(anchor="w", pady=(0, 5))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            content_frame,
            mode='determinate',
            length=400
        )
        self.progress_bar.pack(fill="x", pady=(0, 10))
        
        # Statistics frame
        stats_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        stats_frame.pack(fill="x")
        
        # Statistics labels
        self.stats_label = tk.Label(
            stats_frame,
            text="Ready",
            font=("Microsoft YaHei", 9),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        self.stats_label.pack(side="left")
        
        # Status indicators
        status_frame = tk.Frame(stats_frame, bg=COLORS['surface'])
        status_frame.pack(side="right")
        
        self.status_indicators = {
            'convertible': tk.Label(status_frame, text="🟢 Convertible: 0", font=("Microsoft YaHei", 8), fg=COLORS['text'], bg=COLORS['surface']),
            'converted': tk.Label(status_frame, text="🔵 Converted: 0", font=("Microsoft YaHei", 8), fg=COLORS['text'], bg=COLORS['surface']),
            'skipped': tk.Label(status_frame, text="🟡 Skipped: 0", font=("Microsoft YaHei", 8), fg=COLORS['text'], bg=COLORS['surface']),
            'errors': tk.Label(status_frame, text="🔴 Errors: 0", font=("Microsoft YaHei", 8), fg=COLORS['text'], bg=COLORS['surface'])
        }
        
        for indicator in self.status_indicators.values():
            indicator.pack(side="left", padx=(0, 10))

    def create_right_panel(self):
        # 右侧面板内容区域
        content_frame = tk.Frame(self.right_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 日志标题
        log_label = tk.Label(
            content_frame,
            text="📋 转换日志",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        log_label.pack(anchor="w", pady=(0, 5))
        
        # 日志文本区域
        log_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        log_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            font=("Consolas", 9),
            bg="#f8f9fa",
            fg=COLORS['text'],
            relief="flat",
            bd=5
        )
        self.log_text.pack(fill="both", expand=True, pady=(0, 10))
        
        # 日志控制按钮
        log_control_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        log_control_frame.pack(fill="x", pady=(0, 0))
        
        clear_log_btn = tk.Button(
            log_control_frame,
            text="🗑️清空日志",
            command=self.clear_log,
            font=("Microsoft YaHei", 10),
            bg=COLORS['warning'],
            fg="white",
            relief="flat",
            width=12,
            height=2,
            compound="left"
        )
        clear_log_btn.pack(side="left", padx=(0, 10))
        
        save_log_btn = tk.Button(
            log_control_frame,
            text="💾保存日志",
            command=self.save_log,
            font=("Microsoft YaHei", 10),
            bg=COLORS['info'],
            fg="white",
            relief="flat",
            width=12,
            height=2,
            compound="left"
        )
        save_log_btn.pack(side="left")
        
    def create_bottom_panel(self):
        # 底部面板内容区域
        content_frame = tk.Frame(self.bottom_frame, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 进度标题
        progress_label = tk.Label(
            content_frame,
            text="📊 转换进度",
            font=("Microsoft YaHei", 12, "bold"),
            fg=COLORS['text'],
            bg=COLORS['surface']
        )
        progress_label.pack(anchor="w", pady=(0, 5))
        
        # 进度条
        progress_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        progress_frame.pack(fill="x", pady=(0, 8))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill="x")
        
        # 状态信息
        status_frame = tk.Frame(content_frame, bg=COLORS['surface'])
        status_frame.pack(fill="x", pady=(5, 0))
        
        self.status_label = tk.Label(
            status_frame,
            text="✅ 就绪",
            font=("Microsoft YaHei", 11, "bold"),
            fg=COLORS['secondary'],
            bg=COLORS['surface']
        )
        self.status_label.pack(side="left")
        
        self.stats_label = tk.Label(
            status_frame,
            text="📈 可转换文件: 0 | ✅ 已转换: 0 | ⏭️ 跳过: 0 | ❌ 错误: 0",
            font=("Microsoft YaHei", 10),
            fg=COLORS['text_light'],
            bg=COLORS['surface']
        )
        self.stats_label.pack(side="right")
        
    def create_card_frame(self, parent, title=None):
        """创建现代化卡片样式框架"""
        card = tk.Frame(parent, bg=COLORS['surface'], relief='flat', bd=0)
        card.pack(fill="x", pady=(0, 20))
        
        # 添加阴影效果（通过边框模拟）
        shadow = tk.Frame(parent, bg=COLORS['border'], height=2)
        shadow.pack(fill="x", pady=(0, 18))
        
        if title:
            title_frame = tk.Frame(card, bg=COLORS['surface'])
            title_frame.pack(fill="x", padx=25, pady=(20, 10))
            
            title_label = tk.Label(title_frame, text=title, 
                                 font=('Segoe UI', 14, 'bold'),
                                 fg=COLORS['text'], bg=COLORS['surface'])
            title_label.pack(anchor="w")
            
            # 分隔线
            separator = tk.Frame(card, bg=COLORS['border'], height=1)
            separator.pack(fill="x", padx=25, pady=(0, 15))
        
        content_frame = tk.Frame(card, bg=COLORS['surface'])
        content_frame.pack(fill="both", expand=True, padx=25, pady=(0, 25))
        
        return content_frame
        
    def create_header(self, parent):
        """创建标题区域"""
        header_frame = tk.Frame(parent, bg=COLORS['background'])
        header_frame.pack(fill="x", pady=(0, 30))
        
        # 主标题
        title = tk.Label(header_frame, text="📄 Office格式批量转换工具",
                        font=('Segoe UI', 24, 'bold'),
                        fg=COLORS['primary'], bg=COLORS['background'])
        title.pack(anchor="w")
        
        # 副标题
        subtitle = tk.Label(header_frame, text="快速转换您的Office文档到现代格式",
                           font=('Segoe UI', 12),
                           fg=COLORS['text_light'], bg=COLORS['background'])
        subtitle.pack(anchor="w", pady=(5, 0))
        

        
    def create_modern_checkbox(self, parent, text, variable):
        """创建现代化复选框"""
        cb_frame = tk.Frame(parent, bg=COLORS['surface'])
        cb_frame.pack(fill="x", pady=8, padx=10)
        
        cb = tk.Checkbutton(cb_frame, text=text, variable=variable,
                           font=('Segoe UI', 11), fg=COLORS['text'],
                           bg=COLORS['surface'], activebackground=COLORS['surface'],
                           relief='flat', cursor='hand2')
        cb.pack(anchor="w")
        return cb_frame
        

        
    def setup_layout(self):
        # 新的布局使用pack管理器，不需要网格配置
        pass
        
    def browse_directory(self):
        directory = filedialog.askdirectory(title="选择要转换的目录")
        if directory:
            self.source_dir.set(directory)
            
    def log_message(self, message):
        """线程安全的日志记录"""
        self.log_queue.put(message)
        
    def update_progress(self, current, total):
        """更新进度条和统计信息"""
        if total > 0:
            progress = (current / total) * 100
            self.progress_queue.put(progress)
            # 更新统计信息
            stats_text = f"📈 可转换文件: {total} | 🔄 进度: {current}/{total} | ✅ 已转换: {self.converted_files} | ⏭️ 跳过: {self.skipped_files} | ❌ 错误: {self.error_files}"
            self.stats_queue.put(stats_text)
            self.root.update_idletasks()
            
    def update_log(self):
        """更新日志显示"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
            
        try:
            while True:
                progress = self.progress_queue.get_nowait()
                self.progress_var.set(progress)
        except queue.Empty:
            pass
            
        try:
            while True:
                stats_text = self.stats_queue.get_nowait()
                self.stats_label.config(text=stats_text)
        except queue.Empty:
            pass
            
        # 继续定时更新
        self.root.after(100, self.update_log)
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        
    def save_log(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            title="保存日志文件"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                messagebox.showinfo("成功", "日志已保存")
            except Exception as e:
                messagebox.showerror("错误", f"保存日志失败: {e}")
                
    def start_conversion(self):
        if not self.source_dir.get():
            messagebox.showerror("错误", "请选择要转换的目录")
            return
            
        if not os.path.exists(self.source_dir.get()):
            messagebox.showerror("错误", "选择的目录不存在")
            return
            
        if not self.convert_doc.get() and not self.convert_xls.get():
            messagebox.showerror("错误", "请至少选择一种转换类型")
            return
            
        # 更新UI状态
        self.is_converting = True
        self.start_button.config(state=tk.DISABLED, bg=COLORS['border'])
        self.stop_button.config(state=tk.NORMAL, bg=COLORS['danger'])
        self.status_label.config(text="🔄 正在转换...", fg=COLORS['warning'])
        self.progress_var.set(0)
        
        # 清空统计
        self.total_files = 0
        self.converted_files = 0
        self.skipped_files = 0
        self.error_files = 0
        
        # 初始化统计显示
        initial_stats = "📈 可转换文件: 0 | 🔄 进度: 0/0 | ✅ 已转换: 0 | ⏭️ 跳过: 0 | ❌ 错误: 0"
        self.stats_queue.put(initial_stats)
        
        # 在新线程中执行转换
        self.conversion_thread = Thread(target=self.run_conversion, daemon=True)
        self.conversion_thread.start()
        
    def stop_conversion(self):
        self.is_converting = False
        self.status_label.config(text="⏸️ 正在停止...", fg=COLORS['warning'])
        
    def run_conversion(self):
        try:
            self.log_message(f"🚀 开始转换目录: {self.source_dir.get()}")
            
            # 创建归档文件夹
            old_files_path = None
            if self.archive_originals.get():
                old_files_path = self.create_old_files_folder(self.source_dir.get())
            elif self.use_custom_archive.get() and self.custom_archive_dir.get():
                old_files_path = self.custom_archive_dir.get()
                if not os.path.exists(old_files_path):
                    try:
                        os.makedirs(old_files_path)
                        self.log_message(f"创建自定义备份文件夹: {old_files_path}")
                    except OSError as e:
                        self.log_message(f"错误：无法创建自定义备份文件夹 '{old_files_path}': {e}")
                        old_files_path = None
                
            # 统计文件数量
            self.count_files()
            self.log_message(f"📊 统计完成，共找到 {self.total_files} 个文件需要转换")
            
            # 初始化统计显示
            self.update_stats()
            
            current_file = 0
            
            # 转换DOC文件
            if self.convert_doc.get():
                current_file = self.convert_doc_files(self.source_dir.get(), old_files_path, current_file)
                
            # 转换XLS文件
            if self.convert_xls.get():
                current_file = self.convert_xls_files(self.source_dir.get(), old_files_path, current_file)
                
            if self.is_converting:
                self.log_message("🎉 转换完成！")
                self.status_label.config(text="✅ 转换完成", fg=COLORS['secondary'])
                # 显示完成提示弹窗
                completion_message = f"转换任务已完成！\n\n📊 转换统计：\n• 可转换文件数：{self.total_files}\n• 成功转换：{self.converted_files}\n• 跳过文件：{self.skipped_files}\n• 错误文件：{self.error_files}"
                messagebox.showinfo("转换完成", completion_message)
            else:
                self.log_message("⏹️ 转换已停止")
                self.status_label.config(text="⏹️ 已停止", fg=COLORS['text_light'])
                
        except Exception as e:
            self.log_message(f"❌ 转换过程中发生错误: {e}")
            self.status_label.config(text="❌ 转换失败", fg=COLORS['danger'])
        finally:
            # 恢复UI状态
            self.start_button.config(state=tk.NORMAL, bg=COLORS['secondary'])
            self.stop_button.config(state=tk.DISABLED, bg=COLORS['border'])
            self.update_stats()
            
    def count_files(self):
        """统计需要转换的文件数量"""
        self.total_files = 0
        for root, _, files in os.walk(self.source_dir.get()):
            for file in files:
                if file.lower().endswith('.doc') and self.convert_doc.get():
                    self.total_files += 1
                elif file.lower().endswith('.xls') and self.convert_xls.get():
                    self.total_files += 1
                    
    def update_stats(self):
        """更新统计信息"""
        stats_text = f"📈 可转换文件: {self.total_files} | ✅ 已转换: {self.converted_files} | ⏭️ 跳过: {self.skipped_files} | ❌ 错误: {self.error_files}"
        self.stats_queue.put(stats_text)
        
    def create_old_files_folder(self, source_directory):
        old_files_folder_name = "旧格式文件"
        old_files_path = os.path.join(source_directory, old_files_folder_name)
        if not os.path.exists(old_files_path):
            try:
                os.makedirs(old_files_path)
                self.log_message(f"创建备份文件夹: {old_files_path}")
            except OSError as e:
                self.log_message(f"错误：无法创建文件夹 '{old_files_path}': {e}")
                return None
        return old_files_path
        
    def set_file_times(self, target_path, source_path):
        if not self.preserve_timestamps.get():
            return
            
        max_retries = 5
        retry_delay = 0.5
        
        for attempt in range(max_retries):
            try:
                source_stat = os.stat(source_path)
                win_creation_time = pywintypes.Time(source_stat.st_ctime)
                win_access_time = pywintypes.Time(source_stat.st_atime)
                win_modify_time = pywintypes.Time(source_stat.st_mtime)
                
                handle = win32file.CreateFile(
                    target_path,
                    win32con.GENERIC_WRITE,
                    win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                    None,
                    win32con.OPEN_EXISTING,
                    win32con.FILE_ATTRIBUTE_NORMAL,
                    None
                )
                win32file.SetFileTime(handle, win_creation_time, win_access_time, win_modify_time)
                win32file.CloseHandle(handle)
                return
            except pywintypes.error as e:
                if e.winerror == 32 and attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    self.log_message(f"警告: 无法设置时间戳 {os.path.basename(target_path)}: {e}")
                    return
            except Exception as e:
                self.log_message(f"警告: 设置时间戳时发生错误: {e}")
                return
                
    def convert_doc_files(self, source_directory, old_files_path, current_file):
        if not self.is_converting:
            return current_file
            
        self.log_message("开始处理 DOC 文件...")
        word_app = None
        
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            
            for root, _, files in os.walk(source_directory):
                if not self.is_converting:
                    break
                    
                normalized_root = os.path.normpath(root)
                if old_files_path:
                    normalized_old_files_path = os.path.normpath(old_files_path)
                    if normalized_root == normalized_old_files_path or normalized_root.startswith(normalized_old_files_path + os.sep):
                        continue
                        
                for file in files:
                    if not self.is_converting:
                        break
                        
                    if file.lower().endswith(".doc") and not file.lower().startswith("~"):
                        current_file += 1
                        self.update_progress(current_file, self.total_files)
                        
                        doc_file_path = os.path.join(root, file)
                        docx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".docx")
                        
                        # 规范化文件路径，处理特殊字符
                        doc_file_path = os.path.normpath(doc_file_path)
                        docx_file_path = os.path.normpath(docx_file_path)
                        
                        self.log_message(f"处理: {doc_file_path}")
                        
                        doc = None
                        should_move_original = False
                        
                        try:
                            # 检查文件是否真实存在
                            if not os.path.exists(doc_file_path):
                                self.log_message(f"跳过（文件不存在）: {doc_file_path}")
                                continue
                                
                            if os.path.exists(docx_file_path):
                                self.log_message(f"跳过（目标文件已存在）: {docx_file_path}")
                                should_move_original = True
                                self.skipped_files += 1
                                self.update_stats()
                            else:
                                # 使用原始路径，但确保路径格式正确
                                normalized_doc_path = doc_file_path.replace('/', '\\')
                                doc = word_app.Documents.Open(normalized_doc_path, ReadOnly=True, PasswordDocument="")
                                doc.SaveAs2(docx_file_path, FileFormat=12)
                                doc.Close(SaveChanges=0)
                                doc = None
                                self.log_message(f"转换成功: {docx_file_path}")
                                self.set_file_times(docx_file_path, doc_file_path)
                                should_move_original = True
                                self.converted_files += 1
                                self.update_stats()
                                
                        except pythoncom.com_error as ce:
                            error_message = str(ce).lower()
                            if "password" in error_message or "密码" in error_message:
                                self.log_message(f"跳过（密码保护）: {doc_file_path}")
                            else:
                                self.log_message(f"跳过（无法打开）: {doc_file_path} - {ce}")
                            self.skipped_files += 1
                            self.update_stats()
                        except Exception as e:
                            self.log_message(f"错误: {doc_file_path} - {e}")
                            self.error_files += 1
                            self.update_stats()
                        finally:
                            if doc:
                                try:
                                    doc.Close(SaveChanges=0)
                                except:
                                    pass
                                    
                            if should_move_original and (self.archive_originals.get() or self.use_custom_archive.get()) and old_files_path:
                                if os.path.exists(doc_file_path):
                                    try:
                                        shutil.move(doc_file_path, os.path.join(old_files_path, file))
                                        self.log_message(f"已备份: {file}")
                                    except Exception as e_move:
                                        self.log_message(f"备份失败: {doc_file_path} - {e_move}")
                                        
        except Exception as e:
            self.log_message(f"DOC转换过程中发生错误: {e}")
        finally:
            if word_app:
                try:
                    word_app.Quit(SaveChanges=0)
                except:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass
                
        return current_file
        
    def convert_xls_files(self, source_directory, old_files_path, current_file):
        if not self.is_converting:
            return current_file
            
        self.log_message("开始处理 XLS 文件...")
        excel_app = None
        
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.DisplayAlerts = False
            
            for root, _, files in os.walk(source_directory):
                if not self.is_converting:
                    break
                    
                normalized_root = os.path.normpath(root)
                if old_files_path:
                    normalized_old_files_path = os.path.normpath(old_files_path)
                    if normalized_root == normalized_old_files_path or normalized_root.startswith(normalized_old_files_path + os.sep):
                        continue
                        
                for file in files:
                    if not self.is_converting:
                        break
                        
                    if file.lower().endswith(".xls") and not file.lower().startswith("~"):
                        current_file += 1
                        self.update_progress(current_file, self.total_files)
                        
                        xls_file_path = os.path.join(root, file)
                        xlsx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".xlsx")
                        
                        # 规范化文件路径，处理特殊字符
                        xls_file_path = os.path.normpath(xls_file_path)
                        xlsx_file_path = os.path.normpath(xlsx_file_path)
                        
                        self.log_message(f"处理: {xls_file_path}")
                        
                        workbook = None
                        should_move_original = False
                        
                        try:
                            # 检查文件是否真实存在
                            if not os.path.exists(xls_file_path):
                                self.log_message(f"跳过（文件不存在）: {xls_file_path}")
                                continue
                                
                            if os.path.exists(xlsx_file_path):
                                self.log_message(f"跳过（目标文件已存在）: {xlsx_file_path}")
                                should_move_original = True
                                self.skipped_files += 1
                                self.update_stats()
                            else:
                                # 使用原始路径，但确保路径格式正确
                                normalized_xls_path = xls_file_path.replace('/', '\\')
                                workbook = excel_app.Workbooks.Open(
                                    normalized_xls_path,
                                    UpdateLinks=0,
                                    ReadOnly=True,
                                    Format=None,
                                    Password="",
                                    IgnoreReadOnlyRecommended=True
                                )
                                workbook.SaveAs(xlsx_file_path, FileFormat=51)
                                workbook.Close(SaveChanges=False)
                                workbook = None
                                self.log_message(f"转换成功: {xlsx_file_path}")
                                self.set_file_times(xlsx_file_path, xls_file_path)
                                should_move_original = True
                                self.converted_files += 1
                                self.update_stats()
                                
                        except pythoncom.com_error as ce:
                            error_message = str(ce).lower()
                            if "password" in error_message or "密码" in error_message:
                                self.log_message(f"跳过（密码保护）: {xls_file_path}")
                            else:
                                self.log_message(f"跳过（无法打开）: {xls_file_path} - {ce}")
                            self.skipped_files += 1
                            self.update_stats()
                        except Exception as e:
                            self.log_message(f"错误: {xls_file_path} - {e}")
                            self.error_files += 1
                            self.update_stats()
                        finally:
                            if workbook:
                                try:
                                    workbook.Close(SaveChanges=False)
                                except:
                                    pass
                                    
                            if should_move_original and (self.archive_originals.get() or self.use_custom_archive.get()) and old_files_path:
                                if os.path.exists(xls_file_path):
                                    try:
                                        shutil.move(xls_file_path, os.path.join(old_files_path, file))
                                        self.log_message(f"已备份: {file}")
                                    except Exception as e_move:
                                        self.log_message(f"备份失败: {xls_file_path} - {e_move}")
                                        
        except Exception as e:
            self.log_message(f"XLS转换过程中发生错误: {e}")
        finally:
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass
                
        return current_file

def main():
    root = tk.Tk()
    app = OfficeConverterGUI(root)
    
    # 设置窗口关闭事件
    def on_closing():
        if app.is_converting:
            if messagebox.askokcancel("退出确认", "转换正在进行中，确定要退出程序吗？"):
                app.is_converting = False
                root.destroy()
        else:
            # 没有转换时直接关闭程序
            root.destroy()
            
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()