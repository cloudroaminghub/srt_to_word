from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def parse_time(time_str):
    """解析时间字符串为毫秒"""
    hours, minutes, seconds, milliseconds = map(int, re.split('[:,]', time_str))
    return ((hours * 60 + minutes) * 60 + seconds) * 1000 + milliseconds

def read_and_process_srt(file_path, max_length=200, comma_threshold=1000, period_threshold=3000):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    output = []
    current_segment = ""
    prev_end_time = 0
    for line in lines:
        if '-->' in line:
            start_time, end_time = line.strip().split(' --> ')
            start_time = parse_time(start_time)
            end_time = parse_time(end_time)
            time_diff = start_time - prev_end_time

            if current_segment and time_diff < comma_threshold and len(current_segment) + len(line.strip()) <= max_length:
                current_segment += ','
            else:
                if current_segment:
                    # 确保每个段落以句号结束
                    if not current_segment.endswith('。'):
                        current_segment += '。'
                    output.append(current_segment)
                    current_segment = ""

            prev_end_time = end_time

        elif line.strip() and not line.strip().isdigit():
            if len(current_segment) + len(line.strip()) > max_length:
                # 确保每个段落以句号结束
                if not current_segment.endswith('。'):
                    current_segment += '。'
                output.append(current_segment)
                current_segment = line.strip()
            else:
                current_segment += line.strip()

    # 添加最后一个段落
    if current_segment:
        if not current_segment.endswith('。'):
            current_segment += '。'
        output.append(current_segment)

    return '\n\n'.join(output)

def set_default_font(doc):
    """设置文档的默认字体为宋体"""
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def save_to_word(text, file_path):
    doc = Document()
    set_default_font(doc)
    doc.add_paragraph(text)
    doc.save(file_path)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        max_length = request.form.get('max_length', 200, type=int)
        comma_threshold = request.form.get('comma_threshold', 1000, type=int)
        period_threshold = request.form.get('period_threshold', 3000, type=int)

        if file.filename == '':
            return render_template('upload.html', message='No selected file')

        if file and file.filename.endswith('.srt'):
            filename = secure_filename(file.filename)
            srt_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(srt_path)

            processed_text = read_and_process_srt(srt_path, max_length, comma_threshold, period_threshold)
            word_filename = filename.rsplit('.', 1)[0] + '.docx'
            word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
            save_to_word(processed_text, word_path)

            return render_template('upload.html', message=f'File uploaded successfully. Download: <a href="/downloads/{word_filename}">{word_filename}</a>')

    return render_template('upload.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)
