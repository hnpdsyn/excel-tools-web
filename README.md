# 📊 Excel 工具集 Web 版

> 合并 / 拆分 / 清理 / 格式 / 转换 —— 手机、电脑浏览器直接用，无需安装任何软件

Streamlit 构建的网页版 Excel 工具箱，与桌面版功能对齐。启动后局域网内手机也能访问，适合手机上临时处理表格。

## 功能

- **合并**：多文件按行合并（自动去表头）、按 Sheet 合并，支持 xlsx/xls/csv
- **拆分**：按列值拆分、按行数均分，自动命名输出
- **清理**：去空格、删空行空列、去重复行、清除特殊字符
- **格式**：统一字体对齐、列宽自适应、一键边框、表头加粗
- **转换**：Excel ↔ CSV 互转（utf-8/gbk 编码可选）

## 使用方法

### 方法一：一键启动（推荐）

Windows 双击 `start.bat`，自动安装依赖并打开浏览器。

### 方法二：手动运行

```bash
pip install streamlit pandas openpyxl
streamlit run app.py
```

浏览器访问 `http://localhost:8501` 即可。手机访问：同一 WiFi 下用 `http://电脑IP:8501`。

## 技术栈

Python 3.8+ / Streamlit / pandas / openpyxl

## 🛒 成品获取

本仓库为开源源码。**免配置即用版**（一键启动包+视频教程+答疑）已在闲鱼上架，搜索「Excel 合并拆分工具」即可找到。
