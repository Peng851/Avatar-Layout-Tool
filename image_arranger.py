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
        self.window = tk.Tk()
        self.window.title("头像排版工具")
        self.window.geometry("1000x800")
        self.window.configure(bg='#f0f0f0')
        
        # 获取系统字体列表并处理
        self.available_fonts = []
        self.font_paths = {}
        self.init_fonts()
        
        # 初始化所有变量
        self.init_variables()
        
        # 初始化配置文件
        self.config_file = "arranger_config.json"
        self.load_config()
        
        # 创建UI（包括日志文本框）
        self.create_ui()
        
        # 初始化日志（在创建UI之后）
        self.setup_logging()
        
        # 应用已保存的配置
        if hasattr(self, 'config_to_apply') and self.config_to_apply:
            self.apply_config(self.config_to_apply)
        
        # 初始化预览窗口
        self.preview_window = None
    
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
                                                try:
                                                    for encoding in ('utf-16-be', 'utf-8', 'ascii'):
                                                        try:
                                                            name = record.string.decode(encoding)
                                                            if all(ord(c) < 0x10000 for c in name):
                                                                cursor.execute(
                                                                    "INSERT OR REPLACE INTO fonts VALUES (?, ?, ?, ?)",
                                                                    (name, full_path, i, mtime)
                                                                )
                                                                break
                                                        except:
                                                            continue
                                                except:
                                                    continue
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
        self.border_color = '#000000'
        self.border_width_var = tk.StringVar(value="2")
        
        # 标题设置
        self.title_color = '#000000'
        self.title_align_var = tk.StringVar(value="居中")
        self.title_bottom_margin_var = tk.StringVar(value="200")
        self.title_side_margin_var = tk.StringVar(value="0")
        
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
    
    def create_ui(self):
        # 创建主窗口的左右分栏
        main_container = tk.PanedWindow(self.window, orient=tk.HORIZONTAL, bg='#f0f0f0')
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 左侧控制面板
        left_frame = tk.Frame(main_container, bg='#f0f0f0')
        main_container.add(left_frame)
        
        # 右侧日志窗口
        log_frame = ttk.LabelFrame(main_container, text="运行日志")
        main_container.add(log_frame)
        
        # 创建日志文本框
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, width=50, height=30)
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        # 在左侧创建UI元素
        self.create_file_selection(left_frame)
        self.create_layout_settings(left_frame)  # 改名：布局设置
        self.create_avatar_settings(left_frame)  # 新增：头像设置
        self.create_title_settings(left_frame)   # 新增：标题设置
        
        # 添加按钮区域
        button_frame = tk.Frame(left_frame, bg='#f0f0f0')
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
        status_frame = tk.Frame(left_frame, bg='#f0f0f0')
        status_frame.pack(fill='x', pady=10)
        
        # 状态文本
        self.status_var = tk.StringVar()
        status_label = tk.Label(status_frame, 
                               textvariable=self.status_var,
                               bg='#f0f0f0')
        status_label.pack()
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(status_frame,
                                     variable=self.progress_var,
                                     maximum=100)
        progress_bar.pack(fill='x', pady=5)
    
    def create_file_selection(self, parent):
        file_frame = ttk.LabelFrame(parent, text="文件选择", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))
        
        # 背景图片选择
        bg_frame = tk.Frame(file_frame, bg='#f0f0f0')
        bg_frame.pack(fill='x', pady=5)
        tk.Button(bg_frame, 
                 text="选择背景图片",
                 bg='white',
                 relief='solid',
                 borderwidth=1,
                 command=self.select_background).pack(side='left', padx=(0, 10))
        self.bg_path_var = tk.StringVar()
        tk.Entry(bg_frame, textvariable=self.bg_path_var, width=50).pack(side='left', fill='x', expand=True)
        
        # 头像文件夹选择
        avatar_frame = tk.Frame(file_frame, bg='#f0f0f0')
        avatar_frame.pack(fill='x', pady=5)
        tk.Button(avatar_frame,
                 text="选择头像文件夹",
                 bg='white',
                 relief='solid',
                 borderwidth=1,
                 command=self.select_avatars_folder).pack(side='left', padx=(0, 10))
        self.avatar_path_var = tk.StringVar()
        tk.Entry(avatar_frame, textvariable=self.avatar_path_var, width=50).pack(side='left', fill='x', expand=True)
    
    def create_layout_settings(self, parent):
        """创建布局设置区域"""
        layout_frame = ttk.LabelFrame(parent, text="布局设置", padding=10)
        layout_frame.pack(fill='x', pady=(0, 10))
        
        # 头像比例选择
        ratio_frame = tk.Frame(layout_frame, bg='#f0f0f0')
        ratio_frame.pack(fill='x', pady=5)
        tk.Label(ratio_frame, text="头像比例:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        ttk.Combobox(ratio_frame, 
                    textvariable=self.ratio_var, 
                    values=["4:5", "1:1"],
                    width=10,
                    state="readonly").pack(side='left', padx=(0, 20))
        
        # 布局方式选择
        tk.Label(ratio_frame, text="布局方式:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        ttk.Combobox(ratio_frame,
                    textvariable=self.layout_type_var,
                    values=["左对齐布局", "居中布局"],
                    width=10,
                    state="readonly").pack(side='left')
        
        # 避让区域设置
        avoid_frame = tk.Frame(layout_frame, bg='#f0f0f0')
        avoid_frame.pack(fill='x', pady=5)
        
        tk.Label(avoid_frame, text="避让区域:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        ttk.Combobox(avoid_frame,
                    textvariable=self.avoid_area_var,
                    values=["无", "中部", "下部"],
                    width=10,
                    state="readonly").pack(side='left', padx=(0, 20))
        
        tk.Label(avoid_frame, text="避让数量:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        ttk.Combobox(avoid_frame,
                    textvariable=self.avoid_count_var,
                    values=["2", "4"],
                    width=10,
                    state="readonly").pack(side='left')
        
        # 边距设置
        margin_frame = tk.Frame(layout_frame, bg='#f0f0f0')
        margin_frame.pack(fill='x', pady=5)
        
        # 上边距
        tk.Label(margin_frame, text="上边距:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        self.top_margin_var = tk.StringVar(value="270")
        tk.Entry(margin_frame, textvariable=self.top_margin_var, width=8).pack(side='left', padx=(0, 20))
        
        # 下边距
        tk.Label(margin_frame, text="下边距:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        self.bottom_margin_var = tk.StringVar(value="650")
        tk.Entry(margin_frame, textvariable=self.bottom_margin_var, width=8).pack(side='left', padx=(0, 20))
        
        # 左右边距（统一）
        tk.Label(margin_frame, text="左右边距:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        self.side_margin_var = tk.StringVar(value="130")
        tk.Entry(margin_frame, textvariable=self.side_margin_var, width=8).pack(side='left')
    
    def create_avatar_settings(self, parent):
        """创建头像设置区域"""
        avatar_frame = ttk.LabelFrame(parent, text="头像设置", padding=10)
        avatar_frame.pack(fill='x', pady=(0, 10))
        
        # 姓名字体设置
        name_font_frame = tk.Frame(avatar_frame, bg='#f0f0f0')
        name_font_frame.pack(fill='x', pady=5)
        
        tk.Label(name_font_frame, text="姓名字体:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        font_frame, self.name_font_combo = self.create_font_combobox(name_font_frame, self.name_font_var)
        font_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(name_font_frame, text="字号:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        tk.Entry(name_font_frame, textvariable=self.name_size_var, width=8).pack(side='left')
        
        # 边框设置
        border_frame = tk.Frame(avatar_frame, bg='#f0f0f0')
        border_frame.pack(fill='x', pady=5)
        
        tk.Checkbutton(border_frame, 
                       text="启用边框",
                       variable=self.border_enabled,
                       bg='#f0f0f0').pack(side='left', padx=(0, 20))
        
        tk.Button(border_frame,
                  text="边框颜色",
                  bg='white',
                  relief='solid',
                  borderwidth=1,
                  command=self.choose_border_color).pack(side='left', padx=(0, 20))
        
        tk.Label(border_frame, text="边框宽度:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        tk.Entry(border_frame, textvariable=self.border_width_var, width=8).pack(side='left')
    
    def create_title_settings(self, parent):
        """创建标题设置区域"""
        title_frame = ttk.LabelFrame(parent, text="标题设置", padding=10)
        title_frame.pack(fill='x', pady=(0, 10))
        
        # 标题字体设置
        font_frame = tk.Frame(title_frame, bg='#f0f0f0')
        font_frame.pack(fill='x', pady=5)
        
        tk.Label(font_frame, text="标题字体:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        font_frame, self.title_font_combo = self.create_font_combobox(font_frame, self.class_font_var)
        font_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(font_frame, text="字号:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        tk.Entry(font_frame, textvariable=self.class_size_var, width=8).pack(side='left')
        
        # 标题颜色和对齐
        align_frame = tk.Frame(title_frame, bg='#f0f0f0')
        align_frame.pack(fill='x', pady=5)
        
        tk.Button(align_frame,
                 text="标题颜色",
                 bg='white',
                 relief='solid',
                 borderwidth=1,
                 command=self.choose_title_color).pack(side='left', padx=(0, 20))
        
        tk.Label(align_frame, text="对齐方式:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        ttk.Combobox(align_frame,
                    textvariable=self.title_align_var,
                    values=["左对齐", "居中", "右对齐"],
                    width=10,
                     state="readonly").pack(side='left')
        
        # 标题边距
        margin_frame = tk.Frame(title_frame, bg='#f0f0f0')
        margin_frame.pack(fill='x', pady=5)
        
        tk.Label(margin_frame, text="下边距:", bg='#f0f0f0').pack(side='left', padx=(0, 10))
        tk.Entry(margin_frame, textvariable=self.title_bottom_margin_var, width=8).pack(side='left')
    
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
    
    def select_background(self):
        path = filedialog.askopenfilename(
            title="选择背景图片",
            filetypes=[("Image files", "*.jpg *.png")]
        )
        if path:
            self.background_path = path
            self.bg_path_var.set(path)
    
    def select_avatars_folder(self):
        path = filedialog.askdirectory(title="选择头像文件夹")
        if path:
            self.avatars_folder = path
            self.avatar_path_var.set(path)
    
    def generate_layout(self):
        """生成排版"""
        if not self.background_path or not self.avatars_folder:
            messagebox.showerror("错误", "请先选择背景图片和头像文件夹")
            return
        
        try:
            logging.info("开始生成排版...")
            
            # 2. 获取所有标题文件夹
            class_folders = [f for f in os.listdir(self.avatars_folder) 
                            if os.path.isdir(os.path.join(self.avatars_folder, f))
                            and f != "排版结果"]
            
            if not class_folders:
                logging.error("未找到任何标题文件夹")
                messagebox.showerror("错误", "未找到任何标题文件夹")
                return
            
            logging.info(f"找到 {len(class_folders)} 个标题文件夹")
            
            # 创建输出文件夹（在上一级目录）
            parent_dir = os.path.dirname(self.avatars_folder)  # 只获取上一级目录
            output_folder = os.path.join(parent_dir, "排版完成")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                logging.info(f"创建输出文件夹: {output_folder}")
            
            # 3. 处理每个标题
            total_classes = len(class_folders)
            for i, class_folder in enumerate(class_folders, 1):
                try:
                    logging.info(f"开始处理第 {i}/{total_classes} 个标题: {class_folder}")
                    
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
                    logging.info(f"正在保存: {output_path}")
                    background.save(output_path, quality=95)
                    logging.info(f"已保存: {output_path}")
                    
                except Exception as e:
                    logging.error(f"处理标题 {class_folder} 时出错: {e}")
                    messagebox.showerror("错误", f"处理标题 {class_folder} 时出错：\n{str(e)}")
            
            # 4. 完成处理
            self.progress_var.set(100)
            self.status_var.set("处理完成！")
            logging.info("排版生成完成！")
            
            # 5. 显示完成消息并打开输出文件夹
            messagebox.showinfo("完成", f"排版已完成！\n输出文件夹：{output_folder}")
            os.startfile(output_folder)
            
        except Exception as e:
            logging.error(f"生成排版时出错: {e}")
            messagebox.showerror("错误", f"生成排版时出错：\n{str(e)}")

    def process_class(self, background, class_path, class_name):
        """处理单个标题的排版"""
        try:
            logging.info(f"开始处理标题: {class_name}")
            logging.info(f"标题路径: {class_path}")
            
            # 1. 获取并排序头像文件
            avatars = []
            for file in os.listdir(class_path):
                if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    avatars.append((file, os.path.join(class_path, file)))
            avatars.sort(key=lambda x: x[0])
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
                        # 打开照片并保持比例裁剪
                        avatar = Image.open(filepath)
                        avatar = avatar.convert('RGB')
                        
                        # 计算裁剪尺寸以保持比例
                        target_ratio = 4/5 if self.ratio_var.get() == "4:5" else 1
                        current_ratio = avatar.width / avatar.height
                        
                        if current_ratio > target_ratio:
                            # 图片太宽，需要裁剪宽度
                            new_width = int(avatar.height * target_ratio)
                            left = (avatar.width - new_width) // 2
                            avatar = avatar.crop((left, 0, left + new_width, avatar.height))
                        else:
                            # 图片太高，需要裁剪高度
                            new_height = int(avatar.width / target_ratio)
                            top = (avatar.height - new_height) // 2
                            avatar = avatar.crop((0, top, avatar.width, top + new_height))
                        
                        # 调整大小
                        avatar = avatar.resize((photo_width, photo_height))
                        
                        # 添加边框
                        if self.border_enabled.get():
                            border_width = int(self.border_width_var.get())
                            draw.rectangle(
                                [x-border_width, y-border_width,
                                 x+photo_width+border_width, y+photo_height+border_width],
                                outline=self.border_color,
                                width=border_width
                            )
                        
                        # 粘贴照片
                        background.paste(avatar, (x, y))
                        
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
                            draw.text((text_x, text_y), line, font=name_font, fill='black')
                        
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
            'avoid_count': self.avoid_count_var.get()
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
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
        if not self.background_path:
            messagebox.showerror("错误", "请先选择背景图片")
            return
        
        try:
            # 创建预览窗口
            self.show_preview()
            
            # 等待窗口渲染完成
            self.preview_window.update()
            self.preview_window.after(100)
            
            # 打开背景图片并创建副本
            background = Image.open(self.background_path).copy()
            if background.size != (4800, 3200):
                background = background.resize((4800, 3200))
            
            # 获取第一个文件夹的信息
            class_folders = [f for f in os.listdir(self.avatars_folder) 
                            if os.path.isdir(os.path.join(self.avatars_folder, f))]
            if class_folders:
                first_class = class_folders[0]
                class_path = os.path.join(self.avatars_folder, first_class)
                
                # 处理第一个文件夹的头像和标题
                self.process_class(background, class_path, first_class)
            
            # 计算缩放比例（调小到60%）
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            width_scale = (screen_width * 0.6) / 4800
            height_scale = (screen_height * 0.6) / 3200
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
        """将文字按照最大宽度换行，保持英文单词完整"""
        if not text:
            return []
        
        lines = []
        current_line = ""
        current_word = ""
        
        def is_chinese(char):
            """判断是否为中文字符"""
            return '\u4e00' <= char <= '\u9fff'
        
        def add_word_to_line(word, line):
            """尝试将单词添加到当前行"""
            test_line = line + word if line else word
            return draw.textlength(test_line, font=font) <= max_width, test_line
        
        for char in text:
            if is_chinese(char):
                # 处理中文字符
                if current_word:
                    # 先处理之前的英文单词
                    fits, new_line = add_word_to_line(current_word, current_line)
                    if fits:
                        current_line = new_line
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = current_word
                current_word = ""
            
            # 处理中文字符
            fits, new_line = add_word_to_line(char, current_line)
            if fits:
                current_line = new_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = char
        else:
            if char.isspace():
                # 处理空格，将当前单词添加到行
                if current_word:
                    fits, new_line = add_word_to_line(current_word, current_line)
                    if fits:
                        current_line = new_line
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = current_word
                current_word = ""
            else:
                # 累积英文字符
                current_word += char
    
        # 处理最后剩余的单词
        if current_word:
            fits, new_line = add_word_to_line(current_word, current_line)
            if fits:
                current_line = new_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = current_word
        
        # 添加最后一行
        if current_line:
            lines.append(current_line)
        
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
        frame = tk.Frame(parent, bg='#f0f0f0')
        
        # 创建输入框
        entry = tk.Entry(frame, width=width, textvariable=var)
        entry.pack(side='left')
        
        # 创建下拉列表窗口
        listbox_window = tk.Toplevel(frame)
        listbox_window.withdraw()
        listbox_window.overrideredirect(True)
        listbox_window.transient(self.window)
        
        # 创建列表框和滚动条
        listbox = tk.Listbox(listbox_window, width=width, height=10)
        scrollbar = tk.Scrollbar(listbox_window, orient="vertical", command=listbox.yview)
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
            # 打开并调整头像大小
            with Image.open(avatar_path) as img:
                # 确保图像是 RGBA 模式
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                
                # 如果需要添加边框
                if border_color and border_size > 0:
                    # 先将图像调整到目标大小（减去边框的两倍宽度）
                    inner_size = size - (border_size * 2)
                    resized_img = img.resize((inner_size, inner_size), Image.Resampling.LANCZOS)
                    
                    # 获取调整后图像的实际尺寸
                    actual_width, actual_height = resized_img.size
                    
                    # 计算边框的外部尺寸
                    total_width = actual_width + (border_size * 2)
                    total_height = actual_height + (border_size * 2)
                    
                    # 创建最终图像（使用计算出的实际尺寸）
                    result = Image.new('RGBA', (total_width, total_height), (0, 0, 0, 0))
                    
                    # 使用 ImageDraw 绘制边框
                    draw = ImageDraw.Draw(result)
                    # 外框
                    draw.rectangle([
                        (0, 0), 
                        (total_width - 1, total_height - 1)
                    ], fill=border_color)
                    # 内框（清空内部区域）
                    draw.rectangle([
                        (border_size, border_size),
                        (border_size + actual_width - 1, border_size + actual_height - 1)
                    ], fill=(0, 0, 0, 0))
                    
                    # 将调整大小后的图像精确粘贴到边框内
                    result.paste(resized_img, (border_size, border_size), resized_img)
                    
                    # 如果最终尺寸与目标尺寸不同，进行最后的调整
                    if result.size != (size, size):
                        result = result.resize((size, size), Image.Resampling.LANCZOS)
                    
                    return result
                else:
                    # 如果不需要边框，直接调整大小
                    return img.resize((size, size), Image.Resampling.LANCZOS)
                
        except Exception as e:
            logging.error(f"处理头像时出错: {e}")
            # 创建一个空白的替代图像
            return Image.new('RGBA', (size, size), (200, 200, 200, 255))

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
        if self.preview_window:
            self.preview_window.destroy()
        
        # 创建新窗口
        self.preview_window = tk.Toplevel(self.window)
        self.preview_window.title("预览")
        
        # 计算合适的窗口大小
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # 计算适合的缩放比例（使窗口大小为屏幕的50%）
        width_scale = (screen_width * 0.5) / 4800
        height_scale = (screen_height * 0.5) / 3200
        scale = min(width_scale, height_scale)
        
        preview_width = int(4800 * scale)
        preview_height = int(3200 * scale)
        
        # 确保最小高度包含按钮
        min_height = preview_height + 100  # 增加额外空间给按钮
        
        # 创建主框架
        main_frame = ttk.Frame(self.preview_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 创建画布框架，设置固定高度
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill='both', expand=True)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient='horizontal')
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical')
        canvas = tk.Canvas(canvas_frame, 
                          width=preview_width,
                          height=preview_height,
                          xscrollcommand=h_scrollbar.set,
                          yscrollcommand=v_scrollbar.set)
        
        # 配置滚动条
        h_scrollbar.config(command=canvas.xview)
        v_scrollbar.config(command=canvas.yview)
        
        # 放置组件
        h_scrollbar.pack(side='bottom', fill='x')
        v_scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)  # 添加 expand=True
        
        # 创建按钮框架，不允许压缩
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        # 添加按钮
        ttk.Button(button_frame, text="取消", command=self.preview_window.destroy).pack(side='right', padx=5)
        ttk.Button(button_frame, text="生成", command=self.generate_from_preview).pack(side='right', padx=5)
        
        # 添加缩放功能
        def on_mousewheel(event):
            if event.state == 4:  # Ctrl键被按下
                # 获取当前鼠标位置
                x = canvas.canvasx(event.x)
                y = canvas.canvasy(event.y)
                
                # 根据滚轮方向确定缩放因子
                scale_factor = 1.1 if event.delta > 0 else 0.9
                
                # 执行缩放
                canvas.scale('all', x, y, scale_factor, scale_factor)
                
                # 更新图片
                canvas.delete('background_image')
                canvas.create_image(0, 0, image=self.preview_image, anchor='nw', tags='background_image')
                
                # 重新创建参考线
                canvas.delete('reference_line')
                self.create_preview(canvas, scale * scale_factor)
                
                # 更新滚动区域
                canvas.configure(scrollregion=canvas.bbox('all'))
        
        # 添加拖动功能
        def start_drag(event):
            canvas.scan_mark(event.x, event.y)
        
        def do_drag(event):
            canvas.scan_dragto(event.x, event.y, gain=1)
        
        canvas.bind('<Control-MouseWheel>', on_mousewheel)
        canvas.bind('<ButtonPress-1>', start_drag)
        canvas.bind('<B1-Motion>', do_drag)
        
        # 保存canvas引用
        self.preview_canvas = canvas
        
        # 设置窗口大小和位置
        window_width = preview_width + v_scrollbar.winfo_reqwidth()
        window_height = preview_height + h_scrollbar.winfo_reqheight() + button_frame.winfo_reqheight() + 40
        
        # 居中显示窗口
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.preview_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口最小大小
        self.preview_window.minsize(
            preview_width + v_scrollbar.winfo_reqwidth() + 20,
            min_height
        )

    def create_preview(self, canvas, scale):
        """在画布上创建预览内容"""
        preview_width = int(4800 * scale)
        preview_height = int(3200 * scale)
        
        # 获取布局设置中的边距
        layout_type = self.layout_type_var.get()
        layout_margin = int(self.side_margin_var.get())  # 使用布局设置中的边距值
        
        # 根据布局类型设置左右边距
        if layout_type == "左对齐布局":
            # 左对齐时，左边距固定为布局边距，右边距为图片宽度减去布局边距
            left_margin = layout_margin * scale
            right_margin = preview_width - (layout_margin * scale)
        elif layout_type == "右对齐布局":
            # 右对齐时，右边距固定为布局边距，左边距为图片宽度减去布局边距
            left_margin = layout_margin * scale
            right_margin = preview_width - (layout_margin * scale)
        else:  # 居中布局
            # 居中布局时，左右边距相等
            left_margin = layout_margin * scale
            right_margin = preview_width - (layout_margin * scale)
        
        # 添加头像区域参考线（红色）和标题区域参考线（蓝色）都使用相同的边距
        top_margin = 250 * scale
        bottom_margin_red = 600 * scale
        
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

if __name__ == "__main__":
    try:
        app = ImageArranger()
        app.run()
    except Exception as e:
        print(f"程序运行出错: {e}")
        input("按回车键退出...") 