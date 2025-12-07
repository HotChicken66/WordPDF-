import os
import uuid
import sys
from flask import Flask, render_template, request, send_file, after_this_request, jsonify
from pdf2docx import Converter
from werkzeug.utils import secure_filename
import win32com.client
import pythoncom

import logging

# Configure logging
log_file = os.path.join(os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)), 'app.log')
logging.basicConfig(filename=log_file, level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

logger.info("Application starting...")

try:
    import fitz
    logger.info(f"Successfully imported fitz (PyMuPDF). Version: {fitz.__doc__}")
except ImportError as e:
    logger.error(f"Failed to import fitz: {e}")

try:
    from pdf2docx import Converter
    logger.info("Successfully imported pdf2docx")
except ImportError as e:
    logger.error(f"Failed to import pdf2docx: {e}")

app = Flask(__name__)

# 获取应用程序的基础路径（支持打包后的exe）
def get_base_path():
    """获取应用程序的基础路径，支持PyInstaller打包"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe运行
        return os.path.dirname(sys.executable)
    else:
        # 开发环境运行
        return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()
logger.info(f"Base path: {BASE_PATH}")

# Configure upload and download folders (相对于应用程序路径)
UPLOAD_FOLDER = os.path.join(BASE_PATH, 'uploads')
DOWNLOAD_FOLDER = os.path.join(BASE_PATH, 'downloads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

def cleanup_files(files):
    """Helper to delete files after response is sent."""
    try:
        for f in files:
            if os.path.exists(f):
                os.remove(f)
    except Exception as e:
        logger.error(f"Error cleaning up files: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400
    file = request.files['file']
    if file.filename == '':
        return {'error': 'No selected file'}, 400
    
    if file and file.filename.lower().endswith('.pdf'):
        filename = secure_filename(file.filename)
        # 使用原始文件名（不加UUID前缀），便于用户识别
        output_filename = f"{os.path.splitext(filename)[0]}.docx"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        
        # 如果文件已存在，添加数字后缀
        counter = 1
        base_name = os.path.splitext(output_filename)[0]
        while os.path.exists(output_path):
            output_filename = f"{base_name}_{counter}.docx"
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            counter += 1
        
        file.save(input_path)
        
        try:
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            
            # 清理上传的临时文件
            cleanup_files([input_path])
            
            # 返回成功信息和文件路径
            return jsonify({
                'success': True,
                'filename': output_filename,
                'filepath': os.path.abspath(output_path),
                'folder': os.path.abspath(app.config['DOWNLOAD_FOLDER'])
            })
        except Exception as e:
            logger.error(f"PDF conversion failed: {e}", exc_info=True)
            cleanup_files([input_path])
            return {'error': str(e)}, 500
    
    return {'error': 'Invalid file type'}, 400

@app.route('/convert/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400
    file = request.files['file']
    if file.filename == '':
        return {'error': 'No selected file'}, 400
    
    if file and file.filename.lower().endswith('.docx'):
        filename = secure_filename(file.filename)
        # 使用原始文件名（不加UUID前缀），便于用户识别
        output_filename = f"{os.path.splitext(filename)[0]}.pdf"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        
        # 如果文件已存在，添加数字后缀
        counter = 1
        base_name = os.path.splitext(output_filename)[0]
        while os.path.exists(output_path):
            output_filename = f"{base_name}_{counter}.pdf"
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            counter += 1
        
        file.save(input_path)
        word_app = None
        
        try:
            # 使用 win32com 直接调用 Word
            abs_input_path = os.path.abspath(input_path)
            abs_output_path = os.path.abspath(output_path)
            
            logger.info(f"Converting Word to PDF: {abs_input_path} -> {abs_output_path}")
            
            # 初始化 COM
            pythoncom.CoInitialize()
            
            # 使用 dynamic.Dispatch 避免 gen_py 缓存问题
            import win32com.client.dynamic
            word_app = win32com.client.dynamic.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            
            logger.info("Word application started successfully")
            
            # 打开文档
            doc = word_app.Documents.Open(abs_input_path)
            logger.info("Document opened")
            
            # 保存为 PDF (17 = wdFormatPDF)
            doc.SaveAs(abs_output_path, FileFormat=17)
            logger.info("Document saved as PDF")
            
            # 关闭文档
            doc.Close()
            
            # 检查输出文件
            if not os.path.exists(output_path):
                raise Exception("转换失败：输出文件未生成")
            
            cleanup_files([input_path])
            
            return jsonify({
                'success': True,
                'filename': output_filename,
                'filepath': os.path.abspath(output_path),
                'folder': os.path.abspath(app.config['DOWNLOAD_FOLDER'])
            })
            
        except Exception as e:
            logger.error(f"Word conversion failed: {e}", exc_info=True)
            cleanup_files([input_path])
            return {'error': f'Word转换失败: {str(e)}'}, 500
            
        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
            
    return {'error': 'Invalid file type'}, 400

# 打开文件所在文件夹的API
@app.route('/open-folder', methods=['POST'])
def open_folder():
    """打开指定文件所在的文件夹并选中该文件"""
    try:
        data = request.get_json()
        filepath = data.get('filepath', '')
        
        if filepath and os.path.exists(filepath):
            # 使用Windows资源管理器打开并选中文件
            import subprocess
            subprocess.Popen(f'explorer /select,"{filepath}"')
            return jsonify({'success': True})
        else:
            return jsonify({'error': '文件不存在'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
