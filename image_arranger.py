import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import math
import sys
from tkinter.font import families
import json
import logging
from fontTools.ttLib import TTFont, TTCollection
import sqlite3
from datetime import datetime

class ImageArranger:
    def __init__(self):
        # Flat Remix 风格的颜色方案
        self.COLORS = {
            'primary': '#cadef0',         # 主色调（按钮）
            'primary_light': '#eef4fb',   # 主色调悬浮
            'bg_main': '#ffffff',         # 底层背景
            'bg_light': '#f5f6f7',        # 窗口背景
            'bg_dark': '#e1e3e6',         # 深色背景
            'text_primary': '#1e1e1e',    # 主要文字
            'text_secondary': '#6f7981',  # 次要文字
            'border': '#dde1e5',          # 边框颜色
            'success': '#2ecc71',         # 成功色
            'warning': '#f1c40f',         # 警告色
            'error': '#e74c3c'           # 错误色
        }
        
        self.window = tk.Tk()
        self.window.title("头像排版工具")
        self.window.geometry("1000x800")
        
        # 添加以下代码来自定义标题栏颜色
        self.window.configure(bg=self.COLORS['primary'])  # 使用主题色
        self.window.overrideredirect(True)  # 移除默认标题栏
        
        # 创建自定义标题栏
        title_bar = tk.Frame(
            self.window,
            bg=self.COLORS['primary'],  # 标题栏背景色
            relief='flat',
            bd=0,
            height=30
        )
        title_bar.pack(fill='x')
        
        # 添加标题文本
        title_label = tk.Label(
            title_bar,
            text="头像排版工具",
            bg=self.COLORS['primary'],
            fg=self.COLORS['text_primary'],  # 使用主要文字颜色
            font=('微软雅黑', 10)
        )
        title_label.pack(side='left', padx=10)
        
        # 添加切换最大化的方法
        def toggle_maximize():
            if self.window.state() == 'zoomed':
                self.window.state('normal')
                btn_maximize.configure(text='□')
            else:
                self.window.state('zoomed')
                btn_maximize.configure(text='❐')
        
        self.toggle_maximize = toggle_maximize
        
        # 创建按钮容器来确保水平对齐
        buttons_frame = tk.Frame(
            title_bar,
            bg=self.COLORS['primary'],
            height=30
        )
        buttons_frame.pack(side='right', fill='y')
        
        # 关闭按钮
        btn_close = tk.Button(
            buttons_frame,
            text='×',
            command=self.window.destroy,
            bg=self.COLORS['primary'],
            fg=self.COLORS['text_primary'],
            bd=0,
            font=('微软雅黑', 12),
            width=3,
            cursor='hand2'
        )
        btn_close.pack(side='right', pady=1)
        
        # 为关闭按钮添加悬浮效果（使用错误色）
        def on_close_enter(e):
            btn_close.configure(bg=self.COLORS['error'])
        def on_close_leave(e):
            btn_close.configure(bg=self.COLORS['primary'])
        btn_close.bind('<Enter>', on_close_enter)
        btn_close.bind('<Leave>', on_close_leave)

        # 最大化按钮
        btn_maximize = tk.Button(
            buttons_frame,
            text='▢',
            command=self.toggle_maximize,
            bg=self.COLORS['primary'],
            fg=self.COLORS['text_primary'],
            bd=0,
            font=('微软雅黑', 12),
            width=3,
            cursor='hand2',
            anchor='n'  # 添加这个属性使文本向上对齐
        )
        btn_maximize.pack(side='right', pady=(4, 2))  # 上边距2像素，下边距0像素
        
        # 为最大化按钮添加悬浮效果
        def on_maximize_enter(e):
            btn_maximize.configure(bg=self.COLORS['primary_light'])
        def on_maximize_leave(e):
            btn_maximize.configure(bg=self.COLORS['primary'])
        btn_maximize.bind('<Enter>', on_maximize_enter)
        btn_maximize.bind('<Leave>', on_maximize_leave)

        # 最小化按钮
        btn_minimize = tk.Button(
            buttons_frame,
            text='─',
            command=self.window.iconify,
            bg=self.COLORS['primary'],
            fg=self.COLORS['text_primary'],
            bd=0,
            font=('微软雅黑', 12),
            width=3,
            cursor='hand2'
        )
        btn_minimize.pack(side='right', pady=1)
        
        # 为最小化按钮添加悬浮效果
        def on_minimize_enter(e):
            btn_minimize.configure(bg=self.COLORS['primary_light'])
        def on_minimize_leave(e):
            btn_minimize.configure(bg=self.COLORS['primary'])
        btn_minimize.bind('<Enter>', on_minimize_enter)
        btn_minimize.bind('<Leave>', on_minimize_leave)

        # 添加拖动功能
        def get_pos(event):
            xwin = self.window.winfo_x()
            ywin = self.window.winfo_y()
            startx = event.x_root
            starty = event.y_root
            
            ywin = ywin - starty
            xwin = xwin - startx
            
            def move_window(event):
                self.window.geometry(f'+{event.x_root + xwin}+{event.y_root + ywin}')
            
            title_bar.bind('<B1-Motion>', move_window)
        
        title_bar.bind('<Button-1>', get_pos)
        
        # 统一的控件样式
        self.STYLES = {
            'button': {
                'bg': self.COLORS['primary'],
                'fg': 'white',
                'font': ('微软雅黑', 10),
                'relief': 'flat',
                'padx': 15,
                'pady': 8,
                'cursor': 'hand2'  # 鼠标悬浮时显示手型
            },
            'button_secondary': {
                'bg': self.COLORS['bg_light'],
                'fg': self.COLORS['text_primary'],
                'font': ('微软雅黑', 10),
                'relief': 'flat',
                'padx': 15,
                'pady': 8,
                'cursor': 'hand2'
            },
            'label': {
                'bg': self.COLORS['bg_main'],
                'fg': self.COLORS['text_primary'],
                'font': ('微软雅黑', 10)
            },
            'entry': {
                'relief': 'flat',
                'bd': 1,
                'highlightthickness': 1,
                'highlightcolor': self.COLORS['primary'],
                'highlightbackground': self.COLORS['border']
            }
        }
        
        # 配置主窗口样式
        self.window.configure(bg=self.COLORS['bg_main'])
        
        # 获取系统字体列表并处理
        self.available_fonts = []
        self.font_paths = {}
        self.init_fonts()
        
        # 初始化所有变量
        self.init_variables()
        
        # 初始化配置文件
        self.config_file = "arranger_config.json"
        self.load_config()
        
        # 设置统一的样式
        self.setup_styles()
        
        # 创建UI（包括日志文本框）
        self.create_ui()
        
        # 初始化日志（在创建UI之后）
        self.setup_logging()
        
        # 应用已保存的配置
        if hasattr(self, 'config_to_apply') and self.config_to_apply:
            self.apply_config(self.config_to_apply)
        
        # 初始化预览窗口
        self.preview_window = None
        self.preview_canvas = None
        self.preview_image = None
    
    def setup_styles(self):
        """设置全局样式"""
        style = ttk.Style()
        
        # 完全重写 LabelFrame 的布局，使用最简单的结构
        style.layout('Custom.TLabelframe', [
            # 移除 Custom. 前缀，直接使用 Labelframe.Label
            ('Labelframe.Label', {'sticky': 'nw'})  # 只保留标签，放在左上角
        ])
        
        # 配置标题样式
        style.configure(
            'Custom.TLabelframe',
            background=self.COLORS['bg_main'],
            borderwidth=0
        )
        
        style.configure(
            'Custom.TLabelframe.Label',
            background=self.COLORS['bg_main'],
            foreground=self.COLORS['text_primary'],
            font=('微软雅黑', 10),
            padding=(10, 0, 0, 0)  # 左边距10像素
        )
    
    def process_font_record(self, record, encoding):
        """处理字体记录"""
        try:
            name = record.string.decode(encoding)
            if all(ord(c) < 0x10000 for c in name):
                return name
        except Exception:
            return None
        return None
    
    def init_fonts(self):
        """初始化字体列表和路径"""
        import sqlite3
        from fontTools.ttLib import TTFont, TTCollection
        
        # 创建/连接数据库
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'font_metadata.db')
        
        # 如果数据库文件存在，先删除它以确保表结构正确
        if os.path.exists(db_path):
            os.remove(db_path)
        
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 创建新的表格
        cursor.execute('''CREATE TABLE IF NOT EXISTS fonts
                         (font_name TEXT PRIMARY KEY, 
                          font_path TEXT, 
                          font_index INTEGER DEFAULT 0,
                          last_modified REAL)''')
        
        # 字体目录列表
        fonts_dirs = [
            os.path.join(os.environ['WINDIR'], 'Fonts'),  # 系统字体目录
            os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts')  # 用户字体目录
        ]
        
        # 预设常用字体映射
        default_fonts = {
            '微软雅黑': (os.path.join(os.environ['WINDIR'], 'Fonts', 'msyh.ttc'), 0),
            'Microsoft YaHei': (os.path.join(os.environ['WINDIR'], 'Fonts', 'msyh.ttc'), 0),
            '宋体': (os.path.join(os.environ['WINDIR'], 'Fonts', 'simsun.ttc'), 0),
            'SimSun': (os.path.join(os.environ['WINDIR'], 'Fonts', 'simsun.ttc'), 0),
        }
        
        # 初始化字体映射
        self.font_paths = {}
        self.available_fonts = []
        
        try:
            # 先添加预设字体
            for name, (path, index) in default_fonts.items():
                if os.path.exists(path):
                    cursor.execute(
                        "INSERT OR REPLACE INTO fonts VALUES (?, ?, ?, ?)",
                        (name, path, index, os.path.getmtime(path))
                    )
                    self.font_paths[name] = (path, index)
                    self.available_fonts.append(name)
            
            # 扫描所有字体
            for fonts_dir in fonts_dirs:
                if os.path.exists(fonts_dir):
                    for file in os.listdir(fonts_dir):
                        if file.lower().endswith(('.ttf', '.ttc', '.otf')):
                            try:
                                full_path = os.path.join(fonts_dir, file)
                                mtime = os.path.getmtime(full_path)
                                
                                # 处理 .ttc 文件
                                if file.lower().endswith('.ttc'):
                                    ttc = TTCollection(full_path)
                                    for i, font in enumerate(ttc.fonts):
                                        name_table = font['name']
                                        for record in name_table.names:
                                            if record.nameID in (1, 4, 6):
                                                for encoding in ('utf-16-be', 'utf-8', 'ascii'):
                                                    if name := self.process_font_record(record, encoding):
                                                        cursor.execute(
                                                            "INSERT OR REPLACE INTO fonts VALUES (?, ?, ?, ?)",
                                                            (name, full_path, i, mtime)
                                                        )
                                                        break
                                else:
                                    # 处理单字体文件
                                    font = TTFont(full_path)
                                    name_table = font['name']
                                    for record in name_table.names:
                                        if record.nameID in (1, 4, 6):
                                            try:
                                                for encoding in ('utf-16-be', 'utf-8', 'ascii'):
                                                    try:
                                                        name = record.string.decode(encoding)
                                                        if all(ord(c) < 0x10000 for c in name):
                                                            cursor.execute(
                                                                "INSERT OR REPLACE INTO fonts VALUES (?, ?, ?, ?)",
                                                                (name, full_path, 0, mtime)
                                                            )
                                                            break
                                                    except:
                                                        continue
                                            except:
                                                continue
                                    font.close()
                                
                            except Exception as e:
                                print(f"处理字体文件 {file} 时出错: {e}")
                                continue
            
            # 提交更改
            conn.commit()
            
            # 从数据库加载字体信息
            cursor.execute("SELECT font_name, font_path, font_index FROM fonts")
            for name, path, index in cursor.fetchall():
                self.font_paths[name] = (path, index)
                self.available_fonts.append(name)
            
            # 对字体列表进行排序
            self.available_fonts = sorted(list(set(self.available_fonts)))
            
            # 确保常用字体在列表前面
            preferred_fonts = ['微软雅黑', '宋体', '黑体', '楷体', 'SimSun', 'SimHei', 'Microsoft YaHei']
            for font in reversed(preferred_fonts):
                if font in self.available_fonts:
                    self.available_fonts.remove(font)
                    self.available_fonts.insert(0, font)
                
        finally:
            # 关闭数据库连接
            conn.close()
    
    def init_variables(self):
        """初始化所有变量"""
        # 文件路径
        self.background_path = None
        self.avatars_folder = None
        
        # 字体设置
        self.name_font_var = tk.StringVar(value="微软雅黑")  # 姓名字体
        self.class_font_var = tk.StringVar(value="微软雅黑")  # 标题字体
        self.name_size_var = tk.StringVar(value="40")       # 姓名字号
        self.class_size_var = tk.StringVar(value="120")      # 标题字号
        
        # 布局设置
        self.ratio_var = tk.StringVar(value="4:5")
        self.layout_type_var = tk.StringVar(value="左对齐布局")
        
        # 边框设置
        self.border_enabled = tk.BooleanVar(value=True)
        self.border_width_var = tk.StringVar(value="2")
        
        # 标题设置
        self.title_align_var = tk.StringVar(value="居中")
        self.title_bottom_margin_var = tk.StringVar(value="200")
        self.title_side_margin_var = tk.StringVar(value="0")
        
        # 颜色设置
        self.title_color = '#000000' #标题颜色
        self.border_color = '#000000' #边框颜色
        self.name_color = '#000000'  # 姓名颜色
        
        # 避让设置
        self.avoid_area_var = tk.StringVar(value="无")
        self.avoid_count_var = tk.StringVar(value="2")  # 新增避让数量变量
        
        # 状态变量
        self.status_var = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        
        # 路径显示
        self.bg_path_var = tk.StringVar()
        self.avatar_path_var = tk.StringVar()
        
        # 批处理模式
        self.batch_mode = tk.BooleanVar(value=False)
        
        # 头像设置相关变量
        self.border_enabled = tk.BooleanVar(value=False)
        self.border_width_var = tk.StringVar(value="2")
        self.corner_radius_var = tk.StringVar(value="0")  # 添加圆角变量
    
    def create_ui(self):
        # 创建主容器，使用内边距
        main_container = tk.PanedWindow(
            self.window,
            orient=tk.HORIZONTAL,
            bg=self.COLORS['bg_main'],
            sashwidth=4,
            sashpad=2
        )
        main_container.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 左侧控制面板（带圆角和阴影效果的框架）
        left_frame = tk.Frame(
            main_container,
            bg=self.COLORS['bg_main'],
            highlightthickness=1,
            highlightbackground=self.COLORS['border']
        )
        main_container.add(left_frame)
        
        # 右侧日志窗口
        log_frame = tk.Frame(
            main_container,
            bg=self.COLORS['bg_main'],
            highlightthickness=1,
            highlightbackground=self.COLORS['border']
        )
        main_container.add(log_frame)
        
        # 创建日志文本框
        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            width=50,
            height=30,
            bg=self.COLORS['bg_main'],
            fg=self.COLORS['text_primary'],
            font=('微软雅黑', 9),
            relief='flat',
            padx=10,
            pady=10
        )
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 添加现代风格的滚动条
        scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 在左侧创建UI元素
        self.create_file_selection(left_frame)
        self.create_layout_settings(left_frame)  # 改名：布局设置
        self.create_avatar_settings(left_frame)  # 新增：头像设置
        self.create_title_settings(left_frame)   # 新增：标题设置
        
        # 添加按钮区域
        button_frame = tk.Frame(left_frame, bg=self.COLORS['bg_main'])
        button_frame.pack(pady=20)
        
        # 预览按钮
        preview_btn = tk.Button(button_frame,
                              text="预览效果",
                              bg='white',
                              relief='solid',
                              borderwidth=1,
                              width=20,
                              height=2,
                              command=self.preview_layout)
        preview_btn.pack(side='left', padx=20)
        
        # 生成按钮
        generate_btn = tk.Button(button_frame,
                               text="生成排版",
                               bg='white',
                               relief='solid',
                               borderwidth=1,
                               width=20,
                               height=2,
                               command=self.generate_layout)
        generate_btn.pack(side='left', padx=20)
        
        # 添加状态栏
        status_frame = tk.Frame(left_frame, bg=self.COLORS['bg_main'])
        status_frame.pack(fill='x', pady=10)
        
        # 状态文本
        self.status_var = tk.StringVar()
        status_label = tk.Label(status_frame, 
                               textvariable=self.status_var,
                               bg=self.COLORS['bg_main'])
        status_label.pack()
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(status_frame,
                                     variable=self.progress_var,
                                     maximum=100)
        progress_bar.pack(fill='x', pady=5)
    
    def create_file_selection(self, parent):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(
            parent, 
            text="文件选择",
            style='Custom.TLabelframe'
        )
        file_frame.pack(fill='x', pady=(0, 10))
        
        # 创建内容容器，减小顶部边距
        content_frame = tk.Frame(file_frame, bg=self.COLORS['bg_main'])
        content_frame.pack(fill='x', padx=10, pady=(0, 10))  # 改为10像素的统一边距
        
        # 背景图片选择
        bg_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        bg_frame.pack(fill='x', pady=5)
        
        # 选择背景图片按钮
        select_bg_btn = RoundedButton(
            bg_frame,
            text="选择背景图片",
            command=self.select_background,
            width=120,
            height=30,
            corner_radius=8,
            colors={
                'normal': self.COLORS['primary'],
                'hover': self.COLORS['primary_light'],
                'pressed': self.COLORS['primary_light'],
                'text': self.COLORS['text_primary']
            }
        )
        select_bg_btn.pack(side='left', padx=(0, 10))
        
        # 背景图片路径输入框
        bg_entry = tk.Entry(
            bg_frame,
            textvariable=self.bg_path_var,
            **self.STYLES['entry']
        )
        bg_entry.pack(side='left', fill='x', expand=True)
        
        # 头像文件夹选择
        folder_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        folder_frame.pack(fill='x', pady=5)
        
        # 选择头像文件夹按钮
        select_avatar_btn = RoundedButton(
            folder_frame,
            text="选择头像文件夹",
            command=self.select_avatars_folder,
            width=120,
            height=30,
            corner_radius=8,
            colors={
                'normal': self.COLORS['primary'],
                'hover': self.COLORS['primary_light'],
                'pressed': self.COLORS['primary_light'],
                'text': self.COLORS['text_primary']
            }
        )
        select_avatar_btn.pack(side='left', padx=(0, 10))
        
        # 头像文件夹路径输入框
        avatar_entry = tk.Entry(
            folder_frame,
            textvariable=self.avatar_path_var,
            **self.STYLES['entry']
        )
        avatar_entry.pack(side='left', fill='x', expand=True)
    
    
    def create_layout_settings(self, parent):
        """创建布局设置区域"""
        # 布局设置区域
        layout_frame = ttk.LabelFrame(
            parent, 
            text="布局设置",
            padding=(10, 0, 10, 10),
            style='Custom.TLabelframe'
        )
        layout_frame.pack(fill='x', pady=(0, 10))

        # 设置自定义样式
        style = ttk.Style()
        # 完全移除 LabelFrame 的边框
        style.layout('Custom.TLabelframe', [
            ('Labelframe.padding', {
                'sticky': 'nswe',
                'children': [('Labelframe.text', {'sticky': ''})],
            })
        ])
        style.configure(
            'Custom.TLabelframe',
            background=self.COLORS['bg_main'],
        )
        style.configure(
            'Custom.TLabelframe.Label',
            background=self.COLORS['bg_main'],
            padding=(0, 0, 0, 0)
        )

        # 创建单行布局
        settings_frame = tk.Frame(layout_frame, bg=self.COLORS['bg_main'])
        settings_frame.pack(fill='x', pady=5)
        
        # 头像比例
        tk.Label(settings_frame, text="头像比例:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Combobox(settings_frame, 
                    textvariable=self.ratio_var, 
                    values=["4:5", "1:1"],
                    width=10,
                    state="readonly").pack(side='left', padx=(0, 20))
        
        # 布局方式
        tk.Label(settings_frame, text="布局方式:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Combobox(settings_frame,
                    textvariable=self.layout_type_var,
                    values=["左对齐布局", "居中布局"],
                    width=10,
                    state="readonly").pack(side='left')
        
        # 避让区域设置
        avoid_frame = tk.Frame(layout_frame, bg=self.COLORS['bg_main'])
        avoid_frame.pack(fill='x', pady=5)
        tk.Label(avoid_frame, text="避让区域:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Combobox(avoid_frame,
                    textvariable=self.avoid_area_var,
                    values=["无", "中部", "下部"],
                    width=10,
                    state="readonly").pack(side='left', padx=(0, 20))
        
        tk.Label(avoid_frame, text="避让数量:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Combobox(avoid_frame,
                    textvariable=self.avoid_count_var,
                    values=["1", "2", "3", "4"],
                    width=10,
                    state="readonly").pack(side='left')
        
        # 边距设置
        margin_frame = tk.Frame(layout_frame, bg=self.COLORS['bg_main'])
        margin_frame.pack(fill='x', pady=5)
        
        # 上边距
        tk.Label(margin_frame, text="上边距:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        self.top_margin_var = tk.StringVar(value="270")
        ttk.Entry(margin_frame, textvariable=self.top_margin_var, width=8).pack(side='left', padx=(0, 20))
        
        # 下边距
        tk.Label(margin_frame, text="下边距:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        self.bottom_margin_var = tk.StringVar(value="650")
        ttk.Entry(margin_frame, textvariable=self.bottom_margin_var, width=8).pack(side='left', padx=(0, 20))
        
        # 左右边距（统一）
        tk.Label(margin_frame, text="左右边距:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        self.side_margin_var = tk.StringVar(value="130")
        ttk.Entry(margin_frame, textvariable=self.side_margin_var, width=8).pack(side='left')
    
    def create_avatar_settings(self, parent):
        """创建头像设置区域"""
        avatar_frame = ttk.LabelFrame(
            parent, 
            text="头像设置",
            style='Custom.TLabelframe'
        )
        avatar_frame.pack(fill='x', pady=(0, 10))
        
        # 创建内容容器
        content_frame = tk.Frame(avatar_frame, bg=self.COLORS['bg_main'])
        content_frame.pack(fill='x', padx=10, pady=(10, 10))
        
        # 姓名字体设置
        name_font_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        name_font_frame.pack(fill='x', pady=5)
        
        tk.Label(name_font_frame, text="姓名字体:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        font_frame, self.name_font_combo = self.create_font_combobox(name_font_frame, self.name_font_var)
        font_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(name_font_frame, text="字号:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Entry(name_font_frame, textvariable=self.name_size_var, width=8).pack(side='left', padx=(0, 20))
        
        # 将字体颜色按钮移到这里
        tk.Button(name_font_frame,
                  text="字体颜色",
                  bg='white',
                  relief='solid',
                  borderwidth=1,
                  command=self.choose_name_color).pack(side='left', padx=(0, 20))
        
        # 边框设置
        border_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        border_frame.pack(fill='x', pady=5)
        
        tk.Checkbutton(border_frame, 
                       text="启用边框",
                       variable=self.border_enabled,
                       bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 20))
        
        tk.Button(border_frame,
                  text="边框颜色",
                  bg='white',
                  relief='solid',
                  borderwidth=1,
                  command=self.choose_border_color).pack(side='left', padx=(0, 20))
        
        tk.Label(border_frame, text="边框宽度:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Entry(border_frame, textvariable=self.border_width_var, width=8).pack(side='left', padx=(0, 20))
        
        # 圆角设置
        tk.Label(border_frame, text="圆角系数:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Entry(border_frame, textvariable=self.corner_radius_var, width=8).pack(side='left')
    
    def create_title_settings(self, parent):
        """创建标题设置区域"""
        title_frame = ttk.LabelFrame(
            parent, 
            text="标题设置",
            style='Custom.TLabelframe'
        )
        title_frame.pack(fill='x', pady=(0, 10))
        
        # 创建内容容器
        content_frame = tk.Frame(title_frame, bg=self.COLORS['bg_main'])
        content_frame.pack(fill='x', padx=10, pady=(10, 10))
        
        # 标题字体设置
        font_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        font_frame.pack(fill='x', pady=5)
        
        tk.Label(font_frame, text="标题字体:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        font_frame, self.class_font_combo = self.create_font_combobox(font_frame, self.class_font_var)
        font_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(font_frame, text="字号:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Entry(font_frame, textvariable=self.class_size_var, width=8).pack(side='left', padx=(0, 20))
        
        # 将标题颜色按钮移到这里
        tk.Button(font_frame,
                  text="标题颜色",
                  bg='white',
                  relief='solid',
                  borderwidth=1,
                  command=self.choose_title_color).pack(side='left', padx=(0, 20))
        
        # 对齐设置
        align_frame = tk.Frame(content_frame, bg=self.COLORS['bg_main'])
        align_frame.pack(fill='x', pady=5)
        
        tk.Label(align_frame, text="对齐方式:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Combobox(align_frame, textvariable=self.title_align_var, values=["左对齐", "居中", "右对齐"], width=8).pack(side='left', padx=(0, 20))
        
        tk.Label(align_frame, text="底部边距:", bg=self.COLORS['bg_main']).pack(side='left', padx=(0, 10))
        ttk.Entry(align_frame, textvariable=self.title_bottom_margin_var, width=8).pack(side='left', padx=(0, 20))
        
    def create_generate_button(self, parent):
        tk.Button(parent,
                 text="生成排版",
                 bg='white',
                 relief='solid',
                 borderwidth=1,
                 width=20,
                 height=2,
                 command=self.generate_layout).pack(pady=20)
    
    def choose_border_color(self):
        color = colorchooser.askcolor(title="选择边框颜色")
        if color[1]:
            self.border_color = color[1]
    
    def choose_name_color(self):
        """选择姓名文字颜色"""
        color = colorchooser.askcolor(title="选择姓名文字颜色")
        if color[1]:
            self.name_color = color[1]
    
    def select_background(self):
        """选择背景图片"""
        file_path = filedialog.askopenfilename(
            title="选择背景图片",
            filetypes=[("图片文件", "*.jpg;*.jpeg;*.png")]
        )
        
        if file_path:
            self.background_path = file_path
            self.bg_path_var.set(file_path)
            # 导入该背景图片的布局设置
            self.import_settings(file_path)
            logging.info(f"已选择背景图片: {file_path}")
    
    def select_avatars_folder(self):
        """选择头像文件夹"""
        path = filedialog.askdirectory(title="选择头像文件夹")
        if path:
            self.avatars_folder = path
            self.avatar_path_var.set(path)
            logging.info(f"已选择头像文件夹: {path}")
            
            # 检查文件夹结构
            folders = [f for f in os.listdir(path) 
                      if os.path.isdir(os.path.join(path, f))]
            if not folders:
                messagebox.showwarning(
                    "警告", 
                    "所选文件夹内没有找到任何子文件夹。\n请确保头像按照标题分类放在不同的子文件夹中。"
                )
            self.status_var.set(f"已选择头像文件夹，包含 {len(folders)} 个标题")
        else:
            self.avatars_folder = None
            self.avatar_path_var.set("")
            logging.info("取消选择头像文件夹")
    
    def generate_layout(self):
        """生成布局"""
        if not self.background_path or not self.avatars_folder:
            messagebox.showwarning("提示", "请先选择背景图片和头像文件夹")
            return
        
        try:
            logging.info("开始生成排版...")
            
            # 在生成开始时就导出设置
            self.export_settings(self.background_path)
            
            # 获取所有标题文件夹
            class_folders = [f for f in os.listdir(self.avatars_folder) 
                            if os.path.isdir(os.path.join(self.avatars_folder, f))
                            and f != "排版结果"]
            
            if not class_folders:
                logging.error("未找到任何标题文件夹")
                messagebox.showerror("错误", "未找到任何标题文件夹")
                return
            
            class_folders.sort()  # 按名称排序
            
            # 创建输出文件夹（在上一级目录）
            parent_dir = os.path.dirname(self.avatars_folder)
            output_folder = os.path.join(parent_dir, "排版完成")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # 处理每个标题
            total_classes = len(class_folders)
            for i, class_folder in enumerate(class_folders, 1):
                try:
                    # 更新进度
                    self.progress_var.set((i-1) / total_classes * 100)
                    self.status_var.set(f"正在处理：{class_folder} ({i}/{total_classes})")
                    self.window.update()
                    
                    # 打开背景图片
                    background = Image.open(self.background_path)
                    if background.size != (4800, 3200):
                        background = background.resize((4800, 3200))
                    
                    # 处理当前标题
                    class_path = os.path.join(self.avatars_folder, class_folder)
                    self.process_class(background, class_path, class_folder)
                    
                    # 使用班级文件夹名作为文件名
                    output_path = os.path.join(output_folder, f"{class_folder}.jpg")
                    background.save(output_path, quality=95)
                    
                except Exception as e:
                    logging.error(f"处理标题 {class_folder} 时出错: {e}")
                    messagebox.showerror("错误", f"处理标题 {class_folder} 时出错：\n{str(e)}")
            
            # 完成处理
            self.progress_var.set(100)
            self.status_var.set("处理完成！")
            logging.info("排版生成完成！")
            
            # 打开输出目录
            os.startfile(output_folder)
            messagebox.showinfo("完成", "排版已完成！")
            
        except Exception as e:
            logging.error(f"生成布局时出错: {e}")
            messagebox.showerror("错误", f"生成布局时出错: {e}")

    def process_class(self, background, class_path, class_name):
        """处理单个班级的照片"""
        try:
            # 获取所有头像文件
            avatars = []
            for filename in os.listdir(class_path):
                if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                    filepath = os.path.join(class_path, filename)
                    avatars.append((filename, filepath))
            
            if not avatars:
                logging.warning(f"文件夹 {class_name} 中没有找到头像文件")
                return
            
            logging.info(f"开始处理标题: {class_name}")
            logging.info(f"找到 {len(avatars)} 个头像文件")
            
            # 定义基本参数
            side_margin = int(self.side_margin_var.get())
            top_margin = int(self.top_margin_var.get())
            bottom_margin = int(self.bottom_margin_var.get())
            available_width = 4800 - 2 * side_margin
            available_height = 3200 - top_margin - bottom_margin
            h_spacing = 50  # 固定水平间距 - 移到这里定义
            
            # 计算名字所需的空间
            name_font_size = int(self.name_size_var.get())
            name_margin = 10  # 照片和名字之间的间距
            name_height = name_font_size * 2  # 预留两行文字的空间
            min_row_spacing = 50  # 最小行间距
            
            # 计算照片尺寸和布局
            def calculate_layout(total_avatars, test_photo_width):
                # 计算照片高度
                test_photo_height = int(test_photo_width * 5 / 4) if self.ratio_var.get() == "4:5" else test_photo_width
                
                # 计算一个照片块的总高度
                block_height = test_photo_height + name_margin + name_height
                
                # 检查是否符合高度限制
                total_block_height = block_height * 3 + min_row_spacing * 2
                if total_block_height > available_height:
                    return None
                
                # 计算每行最多能放几张照片（使用外部的h_spacing）
                max_per_row = (available_width + h_spacing) // (test_photo_width + h_spacing)
                
                # 确保至少能放下所有照片
                if max_per_row * 3 < total_avatars:
                    return None
                
                return {
                    'photo_width': test_photo_width,
                    'photo_height': test_photo_height,
                    'block_height': block_height,
                    'max_per_row': max_per_row
                }
            
            # 二分查找最大可行的照片尺寸
            total_avatars = len(avatars)
            left = 100  # 最小照片宽度
            right = min(available_width // 2, available_height // 4)  # 最大照片宽度
            optimal_layout = None
            
            while left <= right:
                mid = (left + right) // 2
                layout = calculate_layout(total_avatars, mid)
                
                if layout:
                    optimal_layout = layout
                    left = mid + 1
                else:
                    right = mid - 1
            
            if not optimal_layout:
                raise Exception("无法找到合适的布局方案")
            
            # 修改计算水平间距的逻辑
            def calculate_optimal_spacing(max_row_count, photo_width):
                # 计算最长行的总照片宽度
                total_photos_width = photo_width * max_row_count
                # 计算理想间距（均匀分布）
                optimal_spacing = (available_width - total_photos_width) // (max_row_count + 1)
                return optimal_spacing
            
            # 使用计算出的最优布局
            photo_width = optimal_layout['photo_width']
            photo_height = optimal_layout['photo_height']
            block_height = optimal_layout['block_height']
            max_per_row = optimal_layout['max_per_row']
            
            # 重新计算每行照片数
            if self.layout_type_var.get() == "左对齐布局":
                row_counts = self.calculate_left_aligned_layout(total_avatars, max_per_row)
            else:
                row_counts = self.calculate_centered_layout(total_avatars, max_per_row)
            
            # 计算最优水平间距
            max_row_count = max(row_counts)
            h_spacing = calculate_optimal_spacing(max_row_count, photo_width)
            
            # 计算实际的行间距
            total_blocks_height = block_height * 3
            row_spacing = (available_height - total_blocks_height) // 2
            
            # 4. 处理每一行照片
            draw = ImageDraw.Draw(background)
            current_index = 0
            
            # 计算名字所需的空间
            name_font_size = int(self.name_size_var.get())
            name_margin = 10  # 照片和名字之间的间距
            name_height = name_font_size * 2  # 预留两行文字的空间
            
            # 新的间距计算逻辑
            min_row_spacing = 50  # 最小行间距
            photo_block_height = photo_height + name_margin + name_height
            
            # 计算理想的行间距
            total_height = 3200 - top_margin - bottom_margin
            available_height = total_height - (photo_block_height * 3)
            
            if available_height < min_row_spacing * 2:
                # 如果空间不足，调整照片块大小
                total_spacing = min_row_spacing * 2
                max_block_height = (total_height - total_spacing) // 3
                
                # 重新计算照片高度和名字空间
                name_height = min(name_height, max_block_height * 0.2)  # 名字区域最多占20%
                name_margin = min(name_margin, max_block_height * 0.05)  # 间距最多占5%
                photo_height = max_block_height - name_height - name_margin
                photo_width = photo_height if self.ratio_var.get() == "1:1" else int(photo_height * 4 / 5)
                
                row_spacing = min_row_spacing
            else:
                # 如果空间充足，使用均匀分配的间距
                row_spacing = available_height // 2
            
            # 在循环中使用计算好的间距
            for row, count in enumerate(row_counts):
                # 计算当前行的y坐标，考虑照片块的总高度
                y = top_margin + row * (photo_block_height + row_spacing)
                
                # 计算当前行的起始x坐标
                if self.layout_type_var.get() == "居中布局":
                    row_width = count * photo_width + (count - 1) * h_spacing
                    start_x = side_margin + (available_width - row_width) // 2
                else:
                    start_x = side_margin + h_spacing
                
                # 处理当前行的每张照片
                for col in range(count):
                    if current_index >= len(avatars):
                        break
                        
                    x = start_x + col * (photo_width + h_spacing)
                    
                    if self.should_avoid_position(x, row, photo_width):
                        continue
                    
                    filename, filepath = avatars[current_index]
                    try:
                        # 处理头像（包括圆角和边框）
                        avatar = self.process_avatar(
                            filepath,
                            (photo_width, photo_height),
                            self.border_color if self.border_enabled.get() else None,
                            int(self.border_width_var.get()) if self.border_enabled.get() else 0
                        )
                        
                        # 使用 alpha 通道粘贴处理后的头像
                        background.paste(avatar, (x, y), avatar)  # 使用 avatar 作为 mask
                        
                        # 添加姓名 - 恢复姓名显示
                        name = os.path.splitext(filename)[0]
                        name_font = ImageFont.truetype(
                            self.get_font_path(self.name_font_var.get()),
                            name_font_size
                        )
                        
                        # 将名字按照照片宽度换行
                        wrapped_lines = self.wrap_text(draw, name, name_font, photo_width - 10)  # 留出左右各5像素边距
                        
                        # 计算文本总高度
                        line_spacing = 5  # 行间距
                        total_text_height = len(wrapped_lines) * (name_font_size + line_spacing) - line_spacing
                        
                        # 计算文本起始y坐标（改为上对齐）
                        text_start_y = y + photo_height + name_margin  # 直接从照片底部开始
                        
                        # 绘制每一行文字
                        for i, line in enumerate(wrapped_lines):
                            # 计算当前行的宽度
                            line_width = draw.textlength(line, font=name_font)
                            # 计算x坐标使文字水平居中
                            text_x = x + (photo_width - line_width) // 2
                            # 计算y坐标
                            text_y = text_start_y + i * (name_font_size + 5)  # 5是行间距
                            # 绘制文字
                            draw.text((text_x, text_y), line, font=name_font, fill=self.name_color)
                        
                        current_index += 1
                        
                    except Exception as e:
                        logging.error(f"处理照片 {filename} 时出错: {e}")
                        continue
            
            # 修改标题标题位置，确保下对齐
            class_font = ImageFont.truetype(
                self.get_font_path(self.class_font_var.get()),
                int(self.class_size_var.get())
            )
            
            # 计算文字位置，确保下对齐参考线
            text_width = draw.textlength(class_name, font=class_font)
            text_y = 3200 - int(self.title_bottom_margin_var.get()) - class_font.size  # 减去字体高度确保下对齐
            
            if self.title_align_var.get() == "居中":
                text_x = (4800 - text_width) // 2
            elif self.title_align_var.get() == "右对齐":
                text_x = 4800 - side_margin - text_width
            else:
                text_x = side_margin
            
            # 绘制文字
            draw.text((text_x, text_y), class_name, font=class_font, fill=self.title_color)
            
            logging.info(f"标题 {class_name} 处理完成")
            return True
            
        except Exception as e:
            logging.error(f"处理标题 {class_name} 时出错: {e}")
            raise

    def get_font_path(self, font_name):
        # sourcery skip: inline-immediately-returned-variable
        """获取字体文件路径"""
        # 直接从缓存中获取字体路径和索引
        if font_name in self.font_paths:
            path, index = self.font_paths[font_name]
            if os.path.exists(path):
                return path
        
        # 如果找不到，返回微软雅黑作为后备字体
        logging.warning(f"\n❌ 未找到字体文件: {font_name}")
        logging.warning("使用微软雅黑作为后备字体")
        msyh_path = os.path.join(os.environ['WINDIR'], 'Fonts', 'msyh.ttc')
        return msyh_path

    def choose_title_color(self):
        color = colorchooser.askcolor(title="选择标题颜色", color=self.title_color)
        if color[1]:
            self.title_color = color[1]

    def run(self):
        """运行程序"""
        self.window.mainloop()

    def calculate_layout(self, total_avatars):
        """计算照片布局，考虑避让区域"""
        # 基础计算：每行平均分配
        row_counts = [total_avatars // 3] * 3
        remainder = total_avatars % 3
        
        # 如果有避让区域
        if self.avoid_area_var.get() != "无":
            # 确定避让行
            avoid_row = 1 if self.avoid_area_var.get() == "中部" else 2
            
            # 计算避让区域占用的位置数
            avoid_width = int(self.avoid_width_var.get())
            side_margin = int(self.side_margin_var.get())
            available_width = 4800 - 2 * side_margin
            h_spacing = 50  # 固定间距
            avatar_width = (available_width - 8 * h_spacing) // 7  # 假设最多7个位置
            avoid_positions = math.ceil(avoid_width / (avatar_width + h_spacing))
            
            # 计算避让行需要的照片数（必须是双数）
            avoid_row_count = row_counts[avoid_row]
            if avoid_row_count - avoid_positions < 0:
                # 如果避让区域太大，调整其他行
                needed = avoid_positions - avoid_row_count
                for i in range(3):
                    if i != avoid_row and row_counts[i] > needed // 2:
                        row_counts[i] -= needed // 2
                        row_counts[avoid_row] += needed // 2
                        needed -= needed // 2
                    if needed == 0:
                        break
            
            # 确保避让行是双数
            if row_counts[avoid_row] % 2 == 1:
                row_counts[avoid_row] += 1
                # 从其他行借一个位置
                for i in range(3):
                    if i != avoid_row and row_counts[i] > 0:
                        row_counts[i] -= 1
                        break
        else:
            # 正常分配剩余照片
            for i in range(remainder):
                row_counts[i] += 1
        
        return row_counts

    def get_avoid_x_range(self):
        """计算避让区域的X坐标范围"""
        if self.avoid_area_var.get() == "无":
            return None
        
        # 获取避让照片数量
        avoid_count = int(self.avoid_count_var.get())
        
        # 计算单个照片的宽度
        side_margin = int(self.side_margin_var.get())
        available_width = 4800 - 2 * side_margin
        h_spacing = 50  # 固定间距
        max_photos_per_row = 7
        photo_width = (available_width - (max_photos_per_row + 1) * h_spacing) // max_photos_per_row
        
        # 计算避让区域总宽度（照片宽度 + 间距）* 照片数量
        avoid_width = (photo_width + h_spacing) * avoid_count
        
        # 计算避让区域的起始和结束位置
        center_x = 4800 // 2
        start_x = center_x - avoid_width // 2
        end_x = center_x + avoid_width // 2
        
        return (start_x, end_x)
    
    def should_avoid_position(self, x, row, avatar_width):
        """检查位置是否需要避让"""
        if self.avoid_area_var.get() == "无":
            return False
            
        avoid_range = self.get_avoid_x_range()
        if not avoid_range:
            return False
            
        # 只在需要避让的行进行检查
        target_row = 1 if self.avoid_area_var.get() == "中部" else 2
        if row != target_row:
            return False
            
        # 检查照片是否与避让区域有重叠
        photo_left = x
        photo_right = x + avatar_width
        
        # 检查是否与避让区域有任何重叠（增加容差）
        tolerance = 10  # 10像素的容差
        return not (photo_right <= avoid_range[0] - tolerance or 
                   photo_left >= avoid_range[1] + tolerance)

    def save_config(self):
        config = {
            'name_font': self.name_font_var.get(),
            'class_font': self.class_font_var.get(),
            'name_size': self.name_size_var.get(),
            'class_size': self.class_size_var.get(),
            'ratio': self.ratio_var.get(),
            'layout_type': self.layout_type_var.get(),
            'border_enabled': self.border_enabled.get(),
            'border_color': self.border_color,
            'border_width': self.border_width_var.get(),
            'title_color': self.title_color,
            'title_align': self.title_align_var.get(),
            'title_bottom_margin': self.title_bottom_margin_var.get(),
            'title_side_margin': self.title_side_margin_var.get(),
            'avoid_area': self.avoid_area_var.get(),
            'avoid_count': self.avoid_count_var.get(),
            'name_color': self.name_color  # 新增姓名颜色
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存配置失败: {e}")
    
    def load_config(self):
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 在创建UI之后应用配置
                self.config_to_apply = config
        except:
            self.config_to_apply = None

    def preview_layout(self):
        """预览排版效果"""
        if not self.background_path:  # 先检查背景图片路径是否存在
            messagebox.showerror("错误", "请先选择背景图片")
            return
        
        if not self.avatars_folder:  # 再检查头像文件夹路径是否存在
            messagebox.showerror("错误", "请先选择头像文件夹")
            return
        
        try:
            # 创建预览窗口
            self.show_preview()
            
            # 等待窗口渲染完成
            self.preview_window.update()
            self.preview_window.after(100)
            
            # 打开背景图片并创建副本
            background = Image.open(self.background_path)  # 确保这里 self.background_path 不为 None
            if background.size != (4800, 3200):
                background = background.resize((4800, 3200))
            
            # 获取第一个文件夹的信息
            class_folders = [f for f in os.listdir(self.avatars_folder) 
                            if os.path.isdir(os.path.join(self.avatars_folder, f))
                            and f != "排版结果"]
            
            if not class_folders:
                logging.error("未找到任何标题文件夹")
                messagebox.showerror("错误", "未找到任何标题文件夹")
                return
            
            logging.info(f"找到 {len(class_folders)} 个标题文件夹")
            first_class = class_folders[0]
            class_path = os.path.join(self.avatars_folder, first_class)
            
            # 处理第一个文件夹的头像和标题
            self.process_class(background, class_path, first_class)
            
            # 计算缩放比例（调整到80%）
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            width_scale = (screen_width * 0.5) / 4800  # 从 0.6 改为 0.8
            height_scale = (screen_height * 0.5) / 3200  # 从 0.6 改为 0.8
            scale = min(width_scale, height_scale)
            
            # 缩放图片
            preview_width = int(4800 * scale)
            preview_height = int(3200 * scale)
            background = background.resize((preview_width, preview_height))
            
            # 转换为PhotoImage
            self.preview_image = ImageTk.PhotoImage(background)
            
            # 在画布上显示图片
            self.preview_canvas.create_image(0, 0, image=self.preview_image, anchor='nw')
            
            # 创建预览内容（参考线等）
            self.create_preview(self.preview_canvas, scale)
            
            # 设置滚动区域
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox('all'))
            
        except Exception as e:
            logging.error(f"创建预览时出错: {e}")
            logging.error("错误详情:", exc_info=True)
            messagebox.showerror("错误", f"创建预览时出错: {e}")

    def setup_logging(self):
        """设置日志处理"""
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget
            
            def emit(self, record):
                msg = self.format(record)
                def append():
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.see(tk.END)
                self.text_widget.after(0, append)
        
        # 配置日志
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        
        # 添加文件处理器
        file_handler = logging.FileHandler('arranger.log')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
        
        # 添加文本框处理器
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def apply_config(self, config):
        """应用加载的配置"""
        try:
            # 设置字体
            if 'name_font' in config:
                self.name_font_var.set(config['name_font'])
            if 'class_font' in config:
                self.class_font_var.set(config['class_font'])
            
            # 设置字号
            if 'name_size' in config:
                self.name_size_var.set(config['name_size'])
            if 'class_size' in config:
                self.class_size_var.set(config['class_size'])
            
            # 设置其他选项
            if 'ratio' in config:
                self.ratio_var.set(config['ratio'])
            if 'layout_type' in config:
                self.layout_type_var.set(config['layout_type'])
            if 'border_enabled' in config:
                self.border_enabled.set(config['border_enabled'])
            if 'border_color' in config:
                self.border_color = config['border_color']
            if 'border_width' in config:
                self.border_width_var.set(config['border_width'])
            if 'title_color' in config:
                self.title_color = config['title_color']
            if 'title_align' in config:
                self.title_align_var.set(config['title_align'])
            if 'title_bottom_margin' in config:
                self.title_bottom_margin_var.set(config['title_bottom_margin'])
            if 'title_side_margin' in config:
                self.title_side_margin_var.set(config['title_side_margin'])
            if 'avoid_area' in config:
                self.avoid_area_var.set(config['avoid_area'])
            if 'avoid_count' in config:
                self.avoid_count_var.set(config['avoid_count'])
            if 'name_color' in config:
                self.name_color = config['name_color']
        except Exception as e:
            logging.error(f"应用配置时出错: {e}")

    def on_closing(self):
        """窗口关闭时保存配置"""
        try:
            self.save_config()
        except Exception as e:
            logging.error(f"保存配置时出错: {e}")
        self.window.destroy()

    def preview_first_class(self):
        """预览第一个文件夹的排版效果"""
        if not self.background_path or not self.avatars_folder:
            messagebox.showerror("错误", "请先选择背景图片和头像文件夹")
            return
        
        try:
            # 获取第一个文件夹
            folders = [f for f in os.listdir(self.avatars_folder) 
                      if os.path.isdir(os.path.join(self.avatars_folder, f))]
            if not folders:
                messagebox.showerror("错误", "未找到有效的标题文件夹")
                return
            
            first_folder = folders[0]
            class_path = os.path.join(self.avatars_folder, first_folder)
            
            # 如果预览窗口已存在，就更新它而不是创建新窗口
            if self.preview_window is None or not self.preview_window.winfo_exists():
                # 创建新的预览窗口
                self.preview_window = tk.Toplevel(self.window)
                self.preview_window.title(f"预览效果 - {first_folder}")
            
            # 设置窗口大小和位置
                screen_width = self.preview_window.winfo_screenwidth()
                screen_height = self.preview_window.winfo_screenheight()
            
            # 计算预览图大小，确保适应屏幕
            max_preview_width = int(screen_width * 0.8)
            max_preview_height = int(screen_height * 0.8)
            
            preview_width = max_preview_width
            preview_height = int(3200 * preview_width / 4800)
            
            if preview_height > max_preview_height:
                preview_height = max_preview_height
                preview_width = int(4800 * preview_height / 3200)
            
            # 设置窗口位置
            window_x = (screen_width - preview_width) // 2
            window_y = (screen_height - preview_height) // 2
            self.preview_window.geometry(f"{preview_width}x{preview_height+100}+{window_x}+{window_y}")
                
                # 创建预览图标签
            self.preview_label = tk.Label(self.preview_window)
            self.preview_label.pack(pady=10)
                
                # 添加按钮框架
            button_frame = tk.Frame(self.preview_window)
            button_frame.pack(pady=10)
                
                # 添加确认和取消按钮
            tk.Button(button_frame,
                         text="确认生成全部",
                         command=lambda: [self.preview_window.destroy(), self.generate_layout()],
                         width=15,
                         height=2).pack(side='left', padx=10)
                
            tk.Button(button_frame,
                         text="取消",
                         command=self.preview_window.destroy,
                         width=15,
                         height=2).pack(side='left', padx=10)
            
            # 生成预览图
            background = Image.open(self.background_path)
            if background.size != (4800, 3200):
                background = background.resize((4800, 3200))
            
            # 绘制参考线
            draw = ImageDraw.Draw(background)
            
            # 获取边距值
            top_margin = int(self.top_margin_var.get())
            bottom_margin = int(self.bottom_margin_var.get())
            side_margin = int(self.side_margin_var.get())  # 统一的左右边距
            
            # 绘制照片区域边界线（红色）
            draw.line([(side_margin, top_margin), (4800-side_margin, top_margin)], fill='red', width=2)
            draw.line([(side_margin, 3200-bottom_margin), (4800-side_margin, 3200-bottom_margin)], fill='red', width=2)
            draw.line([(side_margin, top_margin), (side_margin, 3200-bottom_margin)], fill='red', width=2)
            draw.line([(4800-side_margin, top_margin), (4800-side_margin, 3200-bottom_margin)], fill='red', width=2)
            
            # 绘制标题边界线（蓝色）
            title_bottom = int(self.title_bottom_margin_var.get())
            title_height = int(self.class_size_var.get()) + 20  # 字体大小加上一些间距
            
            # 绘制标题参考框
            y_bottom = 3200 - title_bottom
            y_top = y_bottom - title_height
            draw.line([(side_margin, y_bottom), (4800-side_margin, y_bottom)], fill='blue', width=2)
            draw.line([(side_margin, y_top), (side_margin, y_bottom)], fill='blue', width=2)
            draw.line([(4800-side_margin, y_top), (4800-side_margin, y_bottom)], fill='blue', width=2)
            draw.line([(side_margin, y_top), (4800-side_margin, y_top)], fill='blue', width=2)
            
            # 绘制避让区域（绿色）
            if self.avoid_area_var.get() != "无":
                avoid_range = self.get_avoid_x_range()
                if avoid_range:
                    if self.avoid_area_var.get() == "中部":
                        # 中部避让：在第二行中间
                        row_height = (3200 - top_margin - bottom_margin) / 3
                        y_center = top_margin + row_height * 1.5
                        avoid_height = 400
                        y1 = int(y_center - avoid_height/2)
                        y2 = int(y_center + avoid_height/2)
                    else:
                        # 下部避让：在第三行中间
                        row_height = (3200 - top_margin - bottom_margin) / 3
                        y_center = top_margin + row_height * 2.5
                        avoid_height = 400
                        y1 = int(y_center - avoid_height/2)
                        y2 = int(y_center + avoid_height/2)
                    
                    draw.rectangle([avoid_range[0], y1, avoid_range[1], y2], outline='green', width=2)
            
            # 处理排版
            self.process_class(background, class_path, first_folder)
            
            # 缩放预览图
            preview_width = 800
            preview_height = int(3200 * preview_width / 4800)
            preview_img = background.resize((preview_width, preview_height))
            
            # 更新预览图
            photo = ImageTk.PhotoImage(preview_img)
            self.preview_label.configure(image=photo)
            self.preview_label.image = photo
            
            # 将预览窗口提到前台
            self.preview_window.lift()
            
        except Exception as e:
            logging.error(f"生成预览时出错: {e}")
            messagebox.showerror("错误", f"生成预览时出错：\n{str(e)}")

    def wrap_text(self, draw, text, font, max_width):
        """处理文本换行，保留空格"""
        # 如果文本宽度小于最大宽度，直接返回
        if draw.textlength(text, font=font) <= max_width:
            return [text]
        
        # 保留原始空格，将文本分割成单词
        words = text.split(' ')
        lines = []
        current_line = []
        current_width = 0
        
        for word in words:
            # 计算当前单词的宽度（包括一个空格）
            word_width = draw.textlength(word + ' ', font=font)
            
            # 如果加上这个单词会超出宽度限制
            if current_width + word_width > max_width:
                if current_line:  # 如果当前行有内容
                    # 将当前行的单词用空格连接并添加到结果中
                    lines.append(' '.join(current_line))
                    # 开始新的一行
                    current_line = [word]
                    current_width = draw.textlength(word + ' ', font=font)
                else:
                    # 如果单个单词就超过最大宽度，强制添加
                    lines.append(word)
                    current_line = []
                    current_width = 0
            else:
                # 如果没有超出宽度限制，添加到当前行
                current_line.append(word)
                current_width += word_width
        
        # 添加最后一行
        if current_line:
            lines.append(' '.join(current_line))
        
        return lines

    def calculate_left_aligned_layout(self, total_avatars, max_per_row):
        """计算左对齐布局的每行照片数，考虑最大限制"""
        # 基础分配
        base = min(total_avatars // 3, max_per_row)
        remainder = total_avatars - base * 3
        
        row_counts = [base, base, base]
        
        # 分配剩余的照片
        if remainder == 1:
            # 余1时，第三行少两个位置
            if row_counts[0] < max_per_row:
                row_counts[0] += 1
                row_counts[1] += 1
                row_counts[2] -= 1
        elif remainder == 2:
            # 余2时给前两行各加1
            if row_counts[0] < max_per_row and row_counts[1] < max_per_row:
                row_counts[0] += 1
                row_counts[1] += 1
        
        return row_counts

    def calculate_centered_layout(self, total_avatars, max_per_row):
        """计算居中布局的每行照片数，考虑最大限制"""
        # 基础分配
        base = min(total_avatars // 3, max_per_row)
        remainder = total_avatars - base * 3
        
        row_counts = [base, base, base]
        
        # 修改余数分配逻辑
        if remainder == 1:
            # 余1时给中间行
            if row_counts[1] < max_per_row:
                row_counts[1] += 1
        elif remainder == 2:
            # 余2时给第一行和第三行
            if row_counts[0] < max_per_row:
                row_counts[0] += 1
            if row_counts[2] < max_per_row:
                row_counts[2] += 1
        
        return row_counts

    def create_font_combobox(self, parent, var, width=20):
        """创建可搜索的字体下拉框"""
        frame = tk.Frame(parent, bg=self.COLORS['bg_main'])
        
        # 创建输入框
        entry = tk.Entry(frame, width=width, textvariable=var, **self.STYLES['entry'])
        entry.pack(side='left', fill='x', expand=True)
        
        # 创建下拉列表窗口
        listbox_window = tk.Toplevel(frame)
        listbox_window.withdraw()
        listbox_window.overrideredirect(True)
        listbox_window.transient(self.window)
        
        # 创建列表框和滚动条
        listbox = tk.Listbox(listbox_window, width=width, height=10, **self.STYLES['entry'])
        scrollbar = ttk.Scrollbar(listbox_window, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)
        
        # 初始填充字体列表
        for font in self.available_fonts:
            listbox.insert(tk.END, font)
        
        def show_listbox():
            x = entry.winfo_rootx()
            y = entry.winfo_rooty() + entry.winfo_height()
            listbox_window.geometry(f"+{x}+{y}")
            listbox_window.deiconify()
            listbox_window.lift()
        
        def hide_listbox():
            listbox_window.withdraw()
        
        def update_list(event=None):
            search_text = entry.get().lower()
            listbox.delete(0, tk.END)
            for font in self.available_fonts:
                if search_text in font.lower():
                    listbox.insert(tk.END, font)
            show_listbox()
        
        def on_select(event):
            if listbox.curselection():
                selected = listbox.get(listbox.curselection())
                var.set(selected)
                hide_listbox()
                self.window.focus_set()  # 选择后将焦点转移到主窗口
        
        # 绑定事件
        entry.bind('<KeyRelease>', update_list)
        entry.bind('<FocusIn>', lambda e: show_listbox())
        listbox.bind('<<ListboxSelect>>', on_select)
        
        # 点击其他地方时隐藏列表
        def on_click_outside(event):
            if event.widget not in (entry, listbox):
                hide_listbox()
        
        # 只在主窗口上绑定点击事件
        self.window.bind('<Button-1>', on_click_outside, add='+')
        
        # 确保窗口关闭时清理
        def on_destroy(event):
            self.window.unbind('<Button-1>')
            listbox_window.destroy()
        
        frame.bind('<Destroy>', on_destroy)
        
        return frame, entry

    def process_avatar(self, avatar_path, size, border_color=None, border_size=0):
        """处理头像图片"""
        try:
            # 打开并裁剪照片
            with Image.open(avatar_path) as img:
                # 转换为RGB模式
                img = img.convert('RGB')
                
                # 计算裁剪尺寸以保持比例
                target_ratio = 4/5 if self.ratio_var.get() == "4:5" else 1
                current_ratio = img.width / img.height
                
                if current_ratio > target_ratio:
                    # 图片太宽，需要裁剪宽度
                    new_width = int(img.height * target_ratio)
                    left = (img.width - new_width) // 2
                    img = img.crop((left, 0, left + new_width, img.height))
                else:
                    # 图片太高，需要裁剪高度
                    new_height = int(img.width / target_ratio)
                    top = (img.height - new_height) // 2
                    img = img.crop((0, top, img.width, top + new_height))
                
                # 调整大小
                img = img.resize(size)
                
                # 转换为RGBA模式以支持透明度
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                
                # 获取圆角系数
                corner_radius = float(self.corner_radius_var.get()) if self.corner_radius_var.get() else 0
                
                # 如果需要圆角，创建圆角遮罩
                if corner_radius > 0:
                    mask = Image.new('L', size, 0)
                    draw = ImageDraw.Draw(mask)
                    
                    # 计算圆角半径
                    r = min(size[0], size[1]) * corner_radius
                    
                    # 使用 rounded_rectangle 创建真正的圆角
                    draw.rounded_rectangle(
                        [0, 0, size[0]-1, size[1]-1],  # 矩形区域
                        radius=r,                       # 圆角半径
                        fill=255                        # 填充颜色
                    )
                    
                    # 应用遮罩
                    img.putalpha(mask)
                
                # 如果需要添加边框
                if border_color and border_size > 0:
                    result = Image.new('RGBA', size, border_color)
                    
                    if corner_radius > 0:
                        result.putalpha(mask)
                    
                    # 计算内部图像尺寸
                    inner_size = (size[0] - 2*border_size, size[1] - 2*border_size)
                    inner_img = img.resize(inner_size)
                    
                    # 粘贴内部图像
                    x = (size[0] - inner_size[0]) // 2
                    y = (size[1] - inner_size[1]) // 2
                    result.paste(inner_img, (x, y), inner_img if corner_radius > 0 else None)
                    
                    return result
                else:
                    return img
        
        except Exception as e:
            logging.error(f"处理头像时出错: {e}")
            logging.error("错误详情:", exc_info=True)
            return Image.new('RGBA', size, (200, 200, 200, 255))

    def save_result(self, image, class_path):
        """保存处理结果"""
        try:
            # 记录输入参数
            logging.info(f"\n保存图片信息:")
            logging.info(f"class_path: {class_path}")
            logging.info(f"avatars_folder: {self.avatars_folder}")
            
            # 获取头像文件夹的上一层目录（即素材文件夹的上一层）
            parent_dir = os.path.dirname(os.path.dirname(self.avatars_folder))
            logging.info(f"计算的保存目录: {parent_dir}")
            
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"montage_{timestamp}.png"
            
            # 构建完整的输出路径
            output_path = os.path.join(parent_dir, output_filename)
            logging.info(f"完整的输出路径: {output_path}")
            
            # 保存图片
            image.save(output_path, "PNG")
            logging.info(f"已保存结果到: {output_path}")
            
            # 打开输出目录
            os.startfile(parent_dir)
            
        except Exception as e:
            logging.error(f"保存结果时出错: {e}")
            logging.error(f"错误详情:", exc_info=True)  # 添加完整的错误堆栈
            messagebox.showerror("错误", f"保存结果时出错: {e}")

    def show_preview(self):
        """显示预览窗口"""
        try:
            # 如果预览窗口已存在，就更新它
            if self.preview_window and self.preview_window.winfo_exists():
                # 更新预览图片
                self.update_preview()
                return
                
            # 如果预览窗口不存在，创建新窗口
            self.preview_window = tk.Toplevel(self.window)
            self.preview_window.title("预览效果")
            self.preview_window.configure(bg=self.COLORS['bg_light'])
            
            # 计算预览图片的尺寸
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            width_scale = (screen_width * 0.5) / 4800
            height_scale = (screen_height * 0.5) / 3200
            self.preview_scale = min(width_scale, height_scale)  # 保存缩放比例
            
            preview_width = int(4800 * self.preview_scale)
            preview_height = int(3200 * self.preview_scale)
            
            # 设置窗口大小
            window_width = preview_width + 200
            window_height = preview_height + 200
            
            # 设置窗口位置（居中）
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            self.preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            # 创建滚动区域
            canvas_frame = tk.Frame(self.preview_window, bg=self.COLORS['bg_light'])
            canvas_frame.pack(fill='both', expand=True)
            
            # 创建画布和滚动条
            self.preview_canvas = tk.Canvas(canvas_frame, bg=self.COLORS['bg_light'])
            h_scrollbar = ttk.Scrollbar(canvas_frame, orient='horizontal', command=self.preview_canvas.xview)
            v_scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.preview_canvas.yview)
            
            # 配置画布的滚动
            self.preview_canvas.configure(
                xscrollcommand=h_scrollbar.set,
                yscrollcommand=v_scrollbar.set
            )
            
            # 放置组件
            h_scrollbar.pack(side='bottom', fill='x')
            v_scrollbar.pack(side='right', fill='y')
            self.preview_canvas.pack(side='left', fill='both', expand=True)
            
            # 添加按钮
            button_frame = tk.Frame(self.preview_window, bg=self.COLORS['bg_light'])
            button_frame.pack(side='bottom', fill='x', padx=10, pady=5)
            
            ttk.Button(button_frame, text="关闭", command=self.preview_window.destroy).pack(side='right', padx=5)
            ttk.Button(button_frame, text="生成", command=self.generate_from_preview).pack(side='right', padx=5)
            
            # 初次显示预览图片
            self.update_preview()
            
        except Exception as e:
            logging.error(f"创建预览窗口时出错: {e}")
            logging.error("错误详情:", exc_info=True)
            messagebox.showerror("错误", f"创建预览窗口时出错: {e}")

    def update_preview(self):
        """更新预览图片"""
        try:
            # 创建预览图片
            self.create_preview(self.preview_canvas, self.preview_scale)
            
            # 更新画布上的图片
            self.preview_canvas.delete("all")  # 清除旧图片
            self.preview_canvas.create_image(0, 0, image=self.preview_image, anchor='nw')
            
        except Exception as e:
            logging.error(f"更新预览图片时出错: {e}")

    def create_preview(self, canvas, scale):
        """在画布上创建预览内容"""
        preview_width = int(4800 * scale)
        preview_height = int(3200 * scale)
        
        # 获取布局设置中的边距
        layout_type = self.layout_type_var.get()
        layout_margin = int(self.side_margin_var.get())  # 使用布局设置中的边距值
        top_margin = int(self.top_margin_var.get()) * scale  # 使用设置的上边距
        bottom_margin_red = int(self.bottom_margin_var.get()) * scale  # 使用设置的下边距
        
        # 根据布局类型设置左右边距
        if layout_type == "左对齐布局":
            left_margin = layout_margin * scale
            right_margin = preview_width - (layout_margin * scale)
        else:  # 居中布局
            left_margin = layout_margin * scale
            right_margin = preview_width - (layout_margin * scale)
        
        # 添加头像区域参考线（红色）
        # 上边界水平线
        canvas.create_line(
            left_margin, top_margin, 
            right_margin, top_margin, 
            fill='red', width=2, tags='reference_line'
        )
        # 下边界水平线
        canvas.create_line(
            left_margin, preview_height - bottom_margin_red, 
            right_margin, preview_height - bottom_margin_red, 
            fill='red', width=2, tags='reference_line'
        )
        # 左边界垂直线
        canvas.create_line(
            left_margin, top_margin,
            left_margin, preview_height - bottom_margin_red,
            fill='red', width=2, tags='reference_line'
        )
        # 右边界垂直线
        canvas.create_line(
            right_margin, top_margin,
            right_margin, preview_height - bottom_margin_red,
            fill='red', width=2, tags='reference_line'
        )
        
        # 添加标题参考线（蓝色）
        class_font_size = int(self.class_size_var.get()) * scale
        bottom_margin = int(self.title_bottom_margin_var.get()) * scale
        
        # 标题区域使用相同的左右边距
        # 底部水平线
        bottom_line = canvas.create_line(
            left_margin, preview_height - bottom_margin,
            right_margin, preview_height - bottom_margin,
            fill='blue', width=2, tags='reference_line'
        )
        
        # 顶部水平线
        top_line = canvas.create_line(
            left_margin, preview_height - bottom_margin - class_font_size,
            right_margin, preview_height - bottom_margin - class_font_size,
            fill='blue', width=2, tags='reference_line'
        )
        
        # 左侧垂直线
        canvas.create_line(
            left_margin, preview_height - bottom_margin - class_font_size,
            left_margin, preview_height - bottom_margin,
            fill='blue', width=2, tags='reference_line'
        )
        
        # 右侧垂直线
        canvas.create_line(
            right_margin, preview_height - bottom_margin - class_font_size,
            right_margin, preview_height - bottom_margin,
            fill='blue', width=2, tags='reference_line'
        )

    def generate_from_preview(self):
        """从预览状态生成最终图片"""
        try:
            # 关闭预览窗口
            self.preview_window.destroy()
            
            # 调用生成布局方法
            self.generate_layout()
            
        except Exception as e:
            logging.error(f"从预览生成时出错: {e}")
            messagebox.showerror("错误", f"生成时出错: {e}")

    def export_settings(self, background_path):
        """导出当前布局设置"""
        settings = {
            'layout': {
                'ratio': self.ratio_var.get(),
                'layout_type': self.layout_type_var.get(),
                'avoid_area': self.avoid_area_var.get(),
                'avoid_count': self.avoid_count_var.get(),
                'top_margin': self.top_margin_var.get(),
                'bottom_margin': self.bottom_margin_var.get(),
                'side_margin': self.side_margin_var.get()
            },
            'avatar': {
                'name_font': self.name_font_var.get(),
                'name_size': self.name_size_var.get(),
                'name_color': self.name_color,
                'border_enabled': self.border_enabled.get(),
                'border_color': self.border_color,
                'border_width': self.border_width_var.get(),
                'corner_radius': self.corner_radius_var.get()  # 添加圆角系数
            },
            'title': {
                'font': self.class_font_var.get(),
                'size': self.class_size_var.get(),
                'color': self.title_color,
                'align': self.title_align_var.get(),
                'bottom_margin': self.title_bottom_margin_var.get(),
                'side_margin': self.title_side_margin_var.get()
            }
        }
        
        # 生成配置文件路径（与背景图片同名，但后缀为 .layout）
        config_path = os.path.splitext(background_path)[0] + '.layout'
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            logging.info(f"布局设置已保存到: {config_path}")
        except Exception as e:
            logging.error(f"保存布局设置时出错: {e}")
            messagebox.showerror("错误", f"保存布局设置时出错: {e}")

    def import_settings(self, background_path):
        """从背景图片所在目录导入设置"""
        config_path = os.path.splitext(background_path)[0] + '.layout'
        
        if not os.path.exists(config_path):
            logging.info("未找到布局配置文件，使用默认设置")
            return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            # 应用布局设置
            layout = settings.get('layout', {})
            self.ratio_var.set(layout.get('ratio', '4:5'))
            self.layout_type_var.set(layout.get('layout_type', '左对齐布局'))
            self.avoid_area_var.set(layout.get('avoid_area', '无'))
            self.avoid_count_var.set(layout.get('avoid_count', '2'))
            self.top_margin_var.set(layout.get('top_margin', '270'))
            self.bottom_margin_var.set(layout.get('bottom_margin', '650'))
            self.side_margin_var.set(layout.get('side_margin', '130'))
            
            # 应用头像设置
            avatar = settings.get('avatar', {})
            self.name_font_var.set(avatar.get('name_font', '微软雅黑'))
            self.name_size_var.set(avatar.get('name_size', '40'))
            self.name_color = avatar.get('name_color', '#000000')
            self.border_enabled.set(avatar.get('border_enabled', True))
            self.border_color = avatar.get('border_color', '#000000')
            self.border_width_var.set(avatar.get('border_width', '2'))
            self.corner_radius_var.set(avatar.get('corner_radius', '0'))  # 添加圆角系数
            
            # 应用标题设置
            title = settings.get('title', {})
            self.class_font_var.set(title.get('font', '微软雅黑'))
            self.class_size_var.set(title.get('size', '120'))
            self.title_color = title.get('color', '#000000')
            self.title_align_var.set(title.get('align', '居中'))
            self.title_bottom_margin_var.set(title.get('bottom_margin', '200'))
            self.title_side_margin_var.set(title.get('side_margin', '0'))
            
            logging.info(f"已从 {config_path} 导入布局设置")
        except Exception as e:
            logging.error(f"导入布局设置时出错: {e}")
            messagebox.showerror("错误", f"导入布局设置时出错: {e}")

    def choose_name_color(self):
        """选择姓名文字颜色"""
        color = colorchooser.askcolor(title="选择姓名文字颜色")
        if color[1]:
            self.name_color = color[1]

    def add_hover_effect(self, widget, hover_color, normal_color):
        """添加鼠标悬浮效果"""
        def on_enter(e):
            widget['bg'] = hover_color
        def on_leave(e):
            widget['bg'] = normal_color
        
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)

    def create_modern_button(self, parent, text, command, is_primary=True):
        """创建现代风格按钮"""
        style = self.STYLES['button'] if is_primary else self.STYLES['button_secondary']
        btn = tk.Button(parent, text=text, command=command, **style)
        
        # 添加悬浮效果
        if is_primary:
            self.add_hover_effect(btn, self.COLORS['primary_light'], self.COLORS['primary'])
        else:
            self.add_hover_effect(btn, self.COLORS['bg_dark'], self.COLORS['bg_main'])
        
        return btn

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, width=120, height=30, corner_radius=8, **kwargs):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bg=parent['bg'])
        
        # 保存参数
        self.command = command
        self.text = text
        self.corner_radius = corner_radius
        self.colors = kwargs.get('colors', {
            'normal': '#367bf0',      # 正常状态颜色
            'hover': '#5294ff',       # 悬浮状态颜色
            'pressed': '#2860E1',     # 按下状态颜色
            'text': 'white'           # 文字颜色
        })
        
        # 创建圆角矩形背景
        self.states = {}
        for state in ['normal', 'hover', 'pressed']:
            self.states[state] = self._create_rounded_rect(self.colors[state])
            self.itemconfig(self.states[state], state='hidden')
        
        # 绘制文字
        text_x = width // 2
        text_y = height // 2
        self.text_item = self.create_text(
            text_x, text_y,
            text=text,
            fill=self.colors['text'],
            font=('微软雅黑', 10),
            anchor='center'
        )
        
        # 绑定事件
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        
        # 初始显示正常状态
        self._show_state('normal')
    
    def _create_rounded_rect(self, color):
        """创建圆角矩形"""
        x1, y1 = 0, 0
        x2, y2 = self.winfo_reqwidth(), self.winfo_reqheight()
        r = self.corner_radius
        
        # 使用贝塞尔曲线创建圆角
        path = [
            x1+r, y1,                                    # 左上横线起点
            x2-r, y1,                                    # 右上横线终点
            x2-r/2, y1, x2, y1+r/2, x2, y1+r,          # 右上角
            x2, y2-r,                                    # 右边线
            x2, y2-r/2, x2-r/2, y2, x2-r, y2,          # 右下角
            x1+r, y2,                                    # 下横线
            x1+r/2, y2, x1, y2-r/2, x1, y2-r,          # 左下角
            x1, y1+r,                                    # 左边线
            x1, y1+r/2, x1+r/2, y1, x1+r, y1           # 左上角
        ]
        
        return self.create_polygon(path, smooth=True, fill=color)
    
    def _show_state(self, state):
        """显示指定状态，隐藏其他状态"""
        for s in self.states:
            self.itemconfig(self.states[s], state='hidden')
        self.itemconfig(self.states[state], state='normal')
    
    def _on_enter(self, e):
        self._show_state('hover')
    
    def _on_leave(self, e):
        self._show_state('normal')
    
    def _on_press(self, e):
        self._show_state('pressed')
    
    def _on_release(self, e):
        self._show_state('hover')
        if self.command:
            self.command()

if __name__ == "__main__":
    try:
        app = ImageArranger()
        app.run()
    except Exception as e:
        print(f"程序运行出错: {e}")
        input("按回车键退出...")