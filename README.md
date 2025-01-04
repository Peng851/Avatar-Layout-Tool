<<<<<<< HEAD
# 头像排版工具 (Avatar Layout Tool)

一个用于快速生成班级/团队头像排版的工具，支持自定义背景、标题和多种布局方式。专为大型团队照片排版设计，支持批量处理和实时预览。

## 项目概览

本工具旨在解决团队照片排版的痛点，特别适用于：
- 班级毕业照排版
- 团队合影制作
- 活动纪念照排版
- 其他需要统一布局的多人头像排版

## 主要功能

### 1. 文件处理
- 支持选择背景图片（建议尺寸：4800×3200像素）
- 支持选择头像文件夹（支持多个子文件夹批量处理）
- 自动创建输出文件夹

### 2. 布局系统
- 头像比例：支持 4:5 和 1:1 两种比例
- 布局方式：支持左对齐和居中两种布局
- 边距控制：可自定义上下左右边距
- 避让区域：支持中部和下部避让，可设置避让数量

### 3. 样式定制
- 头像边框：支持开关、颜色、粗细设置
- 标题字体：支持系统所有可用字体
- 标题样式：支持字号、颜色、对齐方式调整
- 边距调整：支持标题的底部边距和左右边距

### 4. 预览系统
- 实时预览：所见即所得的效果展示
- 智能参考线：
  - 红色参考线标识头像区域
  - 蓝色参考线标识标题区域
- 交互控制：支持缩放和拖动

## 使用方法

1. 准备工作
   - 准备背景图片（4800×3200像素）
   - 准备头像文件夹，按以下结构组织：
     ```
     主文件夹/
     ├── 班级1/
     │   ├── 学生1.jpg
     │   ├── 学生2.jpg
     │   └── ...
     ├── 班级2/
     │   ├── 学生1.jpg
     │   └── ...
     └── ...
     ```

2. 基本操作流程
   - 启动软件
   - 选择背景图片
   - 选择头像主文件夹
   - 调整布局设置
   - 预览效果
   - 生成最终图片

## 安装依赖

1. Python环境要求
   ```bash
   Python 3.6+
   ```

2. 必要的Python库
   ```bash
   pip install pillow
   pip install tkinter
   pip install fonttools
   ```

## 运行开发

1. 克隆项目
   ```bash
   git clone [项目地址]
   cd [项目目录]
   ```

2. 安装依赖
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序
   ```bash
   python run.py
   ```

## 技术栈

- Python 3.6+
- Tkinter (GUI框架)
- Pillow (图像处理)
- FontTools (字体处理)
- SQLite (字体缓存)

## 项目结构

```
avatar-layout-tool/
├── image_arranger.py    # 主程序文件
├── run.py              # 启动文件
├── requirements.txt    # 项目依赖
├── LICENSE            # MIT 许可证
├── .gitignore         # Git 忽略文件
├── README.md          # 项目文档
└── .vscode/           # VS Code 配置
    └── settings.json  # VS Code 设置
```

## 贡献指南

1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的改动 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启一个 Pull Request

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情 
=======
# Avatar-Layout-Tool
一个用于快速生成班级/团队头像排版的工具，支持自定义背景、标题和多种布局方式。专为大型团队照片排版设计，支持批量处理和实时预览。
>>>>>>> e957ab2916df86f0874751d24fe340d75a11e8ad
