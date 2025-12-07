# Word/PDF 转换器 (Word/PDF Converter)

一个基于 Python Flask 和 pywebview 的本地桌面应用程序，支持 PDF 转 Word 和 Word 转 PDF 功能。

## 功能特点

- **PDF 转 Word**: 将 PDF 文档转换为可编辑的 Word (.docx) 文件。
- **Word 转 PDF**: 将 Word 文档转换为 PDF 文件（需要安装 Microsoft Word）。
- **现代化 UI**: 简洁美观的无边框界面设计。
- **本地运行**: 所有转换在本地完成，保护您的隐私安全。

## 安装与运行

### 前置要求

- Python 3.8+
- Microsoft Word (仅 Word 转 PDF 功能需要)

### 安装步骤

1. 克隆本项目:
   ```bash
   git clone https://github.com/yourusername/WordPdfConvert.git
   cd WordPdfConvert
   ```

2. 安装依赖:
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序:
   ```bash
   python desktop_app.py
   ```

## 打包为 exe

本项目使用 PyInstaller 打包。

```bash
pyinstaller WordPDF转换器.spec
```

打包后的文件位于 `dist` 目录。

## 技术栈

- **后端**: Flask
- **前端**: HTML5, CSS3, JavaScript
- **GUI 框架**: pywebview
- **PDF 转 Word**: pdf2docx
- **Word 转 PDF**: pywin32 (调用 MS Word COM 接口)

## 许可证

MIT License
