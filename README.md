# Excel Column Extractor

**Excel 列提取合并工具** | A powerful tool for extracting and merging columns from multiple Excel files

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

## 📋 项目简介 / Project Description


一个 Excel 列提取和合并工具，支持从多个 Excel 文件中灵活选择列并进行合并。工具提供了直观的图形界面，支持参照列（主键）对齐、格式保留、异步处理等高级功能。

A powerful tool for extracting and merging columns from multiple Excel files. It provides an intuitive graphical interface with advanced features such as reference column alignment, format preservation, and asynchronous processing.

## ✨ 主要特性 / Key Features

- 🔄 **多文件支持** / Multi-file Support
  - 支持同时选择和处理多个 Excel 文件
  - 支持同时选择和处理多个 Excel files simultaneously

- 📊 **灵活列选择** / Flexible Column Selection
  - 可视化选择需要提取的列
  - 支持全选/全不选快捷操作
  - Visual column selection with select all/deselect all shortcuts

- 🔗 **参照列对齐** / Reference Column Alignment
  - 智能检测相同列名
  - 支持选择参照列（主键）进行数据对齐
  - Intelligent duplicate column detection with reference column (primary key) alignment

- 🎨 **格式保留** / Format Preservation
  - 完整保留原始 Excel 文件的格式（字体、颜色、边框等）
  - 保留列宽设置
  - Complete format preservation (fonts, colors, borders, column widths)

- ⚡ **异步处理** / Asynchronous Processing
  - 后台异步处理，不阻塞界面
  - 实时进度显示和取消功能
  - Background asynchronous processing with real-time progress and cancellation

- 🖥️ **现代化界面** / Modern UI
  - 简洁美观的图形界面
  - 支持高 DPI 显示
  - Clean and modern graphical interface with high DPI support

## 🚀 快速开始 / Quick Start

### 环境要求 / Requirements

- Python 3.7 或更高版本 / Python 3.7 or higher
- Windows 操作系统 / Windows OS

### 安装依赖 / Install Dependencies

```bash
pip install openpyxl
```

### 运行程序 / Run the Application

**方式一：直接运行 Python 脚本 / Method 1: Run Python Script**

```bash
python excel_colomn_extraction.py
```

**方式二：使用打包好的可执行文件 / Method 2: Use Packaged Executable**

1. 下载 `Excel列提取工具.exe` 文件
2. 双击运行即可，无需安装 Python 环境

Download `Excel列提取工具.exe` and double-click to run (no Python installation required).

### 打包程序 / Package the Application

使用 PyInstaller 打包为可执行文件：

```bash
pyinstaller excel_colomn_extraction.spec
```

打包完成后，可执行文件位于 `dist` 目录下。

After packaging, the executable will be in the `dist` directory.

## 📖 使用说明 / Usage Guide

### 基本操作流程 / Basic Workflow

1. **选择文件** / **Select Files**
   - 点击"选择Excel文件（可多选）"按钮
   - 选择一个或多个 Excel 文件（支持 .xlsx, .xlsm, .xls 格式）
   - Click "选择Excel文件（可多选）" button
   - Select one or more Excel files (.xlsx, .xlsm, .xls)

2. **选择列** / **Select Columns**
   - 在列选择区域勾选需要提取的列
   - 不同文件用不同颜色标识，便于区分
   - Check the columns you want to extract
   - Different files are color-coded for easy identification

3. **处理相同列名** / **Handle Duplicate Column Names**
   - 如果多个文件包含相同列名，工具会提示选择参照列（主键）
   - 参照列用于数据对齐，确保数据正确合并
   - If multiple files contain the same column name, you'll be prompted to select a reference column (primary key)
   - Reference columns are used for data alignment

4. **导出结果** / **Export Results**
   - 点击"📄 输出拆分合并Excel"按钮
   - 选择保存位置和文件名
   - 等待处理完成
   - Click "📄 输出拆分合并Excel" button
   - Choose save location and filename
   - Wait for processing to complete

### 高级功能 / Advanced Features

- **自动打开结果文件**：勾选"自动打开结果文件"选项，处理完成后自动打开生成的 Excel 文件
- **打开结果文件夹**：点击"📂 打开结果文件夹"按钮，快速定位到输出文件所在目录
- **取消操作**：处理过程中可以随时点击"取消"按钮中断操作

- **Auto-open Result File**: Check "自动打开结果文件" to automatically open the generated Excel file after processing
- **Open Result Folder**: Click "📂 打开结果文件夹" to quickly locate the output directory
- **Cancel Operation**: Click "取消" button anytime during processing to interrupt the operation

## 🛠️ 技术栈 / Tech Stack

- **Python 3.7+** - 编程语言
- **Tkinter** - GUI 框架
- **openpyxl** - Excel 文件处理
- **PyInstaller** - 程序打包

## 📁 项目结构 / Project Structure

```
excel_colomn_extraction/
├── excel_colomn_extraction.py    # 主程序文件
├── excel_colomn_extraction.spec  # PyInstaller 配置文件
├── README.md                     # 项目说明文档
├── CURSOR编程规范.md            # 编程规范文档
└── dist/                         # 打包输出目录
    └── Excel列提取工具.exe       # 可执行文件
```

## 🎯 适用场景 / Use Cases

- 📊 从多个 Excel 文件中提取特定列并合并
- 🔄 数据整合和清洗
- 📈 报表生成和数据分析
- 🔗 基于主键的数据对齐和合并

- Extract and merge specific columns from multiple Excel files
- Data integration and cleaning
- Report generation and data analysis
- Data alignment and merging based on primary keys

## ⚙️ 配置说明 / Configuration

### PyInstaller 打包配置

项目已包含 `excel_colomn_extraction.spec` 配置文件，包含以下优化：

- 单文件打包模式
- 隐藏导入配置（解决 openpyxl 模块导入问题）
- 无控制台窗口（GUI 应用）
- UPX 压缩支持

The project includes `excel_colomn_extraction.spec` with the following optimizations:

- One-file packaging mode
- Hidden imports configuration (fixes openpyxl module import issues)
- No console window (GUI application)
- UPX compression support





