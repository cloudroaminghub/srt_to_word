# SRT to Word Converter

## Description
This Flask web application converts SRT (SubRip Text) subtitle files into Microsoft Word documents. It offers a web interface for uploading SRT files, setting conversion parameters, and downloading the converted Word documents.

## Features
- Web-based SRT file upload.
- Configurable text conversion parameters (segment length, punctuation thresholds).
- Downloadable converted Word document.

## Requirements
- Python 3
- Flask
- python-docx

## Installation
Ensure Python 3, Flask, and python-docx are installed on your system. You can install Flask and python-docx using pip:
```bash
pip install Flask python-docx
```

## Detailed Usage Instructions
1. Clone or download the repository from GitHub:
   ```bash
   git clone https://github.com/cloudroaminghub/srt_to_word.git
   ```
2. Navigate to the script's directory:
   ```bash
   cd srt_to_word
   ```
3. Run the Flask application:
   ```bash
   python app.py
   ```
4. Open a web browser and navigate to `http://127.0.0.1:5000`.
5. Use the web interface to upload an SRT file. Adjust conversion settings as needed.
6. Click on 'Upload and Convert' to process the file.
7. Download the converted Word document by clicking on the provided link.

---

# SRT 转 Word 转换器

## 描述
这个 Flask 网络应用程序可以将 SRT（SubRip 文本）字幕文件转换为 Microsoft Word 文档。它提供了一个网页界面，用于上传 SRT 文件，设置转换参数，并下载转换后的 Word 文档。

## 功能
- 基于网页的 SRT 文件上传。
- 可配置的文本转换参数（段落长度，标点阈值）。
- 可下载的转换后的 Word 文档。

## 需求
- Python 3
- Flask
- python-docx

## 安装
确保您的系统已安装 Python 3、Flask 和 python-docx。您可以使用 pip 安装 Flask 和 python-docx：
```bash
pip install Flask python-docx
```

## 详细使用说明
1. 从 GitHub 克隆或下载仓库：
   ```bash
   git clone https://github.com/cloudroaminghub/srt_to_word.git
   ```
2. 导航到脚本目录：
   ```bash
   cd srt_to_word
   ```
3. 运行 Flask 应用程序：
   ```bash
   python app.py
   ```
4. 打开网页浏览器并导航至 `http://127.0.0.1:5000`。
5. 使用网页界面上传 SRT 文件。根据需要调整转换设置。
6. 点击“上传并转换”处理文件。
7. 通过点击提供的链接下载转换后的 Word 文档。
