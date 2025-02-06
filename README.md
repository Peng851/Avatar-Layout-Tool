# 头像排版工具 (Avatar Layout Tool)

## 项目简介
一个专业的团队照片排版工具，支持自定义背景、标题和多种布局方式。专为大型团队照片排版设计，具备批量处理和实时预览功能。

### 适用场景
- 班级毕业照排版
- 团队合影制作
- 活动纪念照排版
- 其他需要统一布局的多人头像排版

## 功能特性

### 1. 基础功能
- 背景图片处理（支持 4800×3200 像素）
- 批量头像处理
- 实时预览与参考线
- 自动保存布局配置

### 2. 布局系统
- **头像比例**：支持 4:5 和 1:1
- **布局方式**：支持左对齐和居中两种布局
- **避让功能**：
  - 支持中部和下部避让
  - 可设置避让数量（2/4个位置）
- **边距控制**：
  - 上边距（默认：270px）
  - 下边距（默认：650px）
  - 左右边距（默认：130px）

### 3. 样式定制
- **头像设置**：
  - 边框开关
  - 边框颜色和粗细（默认：2px）
  - 姓名字体（默认：微软雅黑）
  - 姓名大小（默认：40px）
  - 姓名颜色
- **标题设置**：
  - 字体选择（支持系统所有字体）
  - 字号调整（默认：120px）
  - 颜色选择
  - 对齐方式（左对齐/居中/右对齐）
  - 底部边距（默认：200px）

### 4. 预览系统
- **实时预览**：所见即所得
- **智能参考线**：
  - 红色：标识头像区域边界
  - 蓝色：标识标题区域边界
- **交互控制**：
  - 支持缩放预览
  - 支持拖动查看

## 使用指南

### 1. 环境要求
- Python 3.6+
- 操作系统：Windows

### 2. 依赖安装
```bash
pip install -r requirements.txt
```

### 3. 文件准备
1. **背景图片**：
   - 推荐尺寸：4800×3200 像素
   - 格式支持：PNG/JPG

2. **头像文件**：
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

### 4. 使用步骤
1. 启动程序：`python run.py`
2. 选择背景图片
3. 选择头像主文件夹
4. 调整布局参数
5. 预览效果
6. 生成最终图片

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
```

## 技术栈
- Python 3.6+
- Tkinter (GUI框架)
- Pillow (图像处理)
- FontTools (字体处理)
- SQLite (字体缓存)

## 常见问题 (FAQ)
1. **为什么需要固定的背景尺寸？**
   - 为确保排版效果的一致性，系统会自动将背景调整为 4800×3200 像素

2. **如何调整避让区域？**
   - 在布局设置中选择"中部"或"下部"避让
   - 可选择2个或4个位置的避让数量

3. **字体显示异常怎么办？**
   - 检查系统字体是否正确安装
   - 程序会自动缓存系统字体信息

## 贡献指南
1. Fork 本仓库
2. 创建特性分支：`git checkout -b feature/AmazingFeature`
3. 提交更改：`git commit -m 'Add some AmazingFeature'`
4. 推送分支：`git push origin feature/AmazingFeature`
5. 提交 Pull Request

## 许可证
本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 更新日志
### v1.0.0 (2024-03-xx)
- 初始版本发布
- 支持基础排版功能
- 添加预览系统
