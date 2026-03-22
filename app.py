import os
import sys
import json
import time
import uuid
import hashlib
import fitz
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from flask import Flask, render_template, request, url_for, jsonify
from werkzeug.utils import secure_filename

# --- 动态获取模板文件夹路径 ---
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
else:
    template_folder = 'templates'

app = Flask(__name__, template_folder=template_folder)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['STATIC_FOLDER'] = 'static'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['STATIC_FOLDER'], exist_ok=True)

# 初始化 OCR 引擎 (全局单例，Gunicorn多进程模式下每个进程会有一个实例)
ocr_engine = RapidOCR()

def get_file_hash(file_stream):
    """计算文件的 MD5 哈希值，用于多用户环境下的精准缓存"""
    md5_hash = hashlib.md5()
    # 分块读取，防止大文件吃光内存
    for chunk in iter(lambda: file_stream.read(4096), b""):
        md5_hash.update(chunk)
    file_stream.seek(0)  # 读完后必须把文件指针拨回头部
    return md5_hash.hexdigest()

@app.route('/', methods=['GET', 'POST'])
def index():
    image_url = None
    ocr_data_json = "[]"
    orig_w, orig_h = 0, 0
    safe_base_name = ""  
    
    # ==============================
    # 🌟 处理从“历史下拉菜单”加载图纸的请求 (仅限 GET)
    # ==============================
    if request.method == 'GET':
        load_file = request.args.get('load')
        if load_file:
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(pdf_path)
            
            pdf_doc = fitz.open(pdf_path)
            page = pdf_doc.load_page(0) 
            # 🚀 内存与超时终极优化：Render 免费版只有 512MB 内存！
            # 将 4x 渲染降为 2x，内存占用暴降 75%，防止服务器崩溃报 502
            zoom = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=zoom)
            pix.save(image_path)
            image_url = url_for('static', filename=image_filename)
            
            img = Image.open(image_path)
                orig_w, orig_h = img.size
                image_url = url_for('static', filename=image_filename)
                
                return render_template('index.html', 
                                       image_url=image_url, ocr_data=ocr_data_json,
                                       orig_w=orig_w, orig_h=orig_h,
                                       filename=safe_base_name)
    
    # ==============================
    # 正常的 POST 上传与初次识别逻辑
    # ==============================
    if request.method == 'POST':
        if 'pdf_file' not in request.files:
            return "没有选择文件", 400
        file = request.files['pdf_file']
        
        if file.filename != '':
            # 🚀 核心升级：多用户并发防御与极速缓存
            # 原来的代码如果张三和李四都上传 1.pdf 会互相覆盖。
            # 现在我们通过文件内容的 MD5 来命名。同一张图纸永远只识别一次，不同图纸哪怕同名也不会冲突！
            file_hash = get_file_hash(file)
            safe_base_name = file_hash  # 用哈希值代替文件名作为系统内部ID
            
            image_filename = f"{safe_base_name}.png"
            json_filename = f"{safe_base_name}.json"
            image_path = os.path.join(app.config['STATIC_FOLDER'], image_filename)
            json_path = os.path.join(app.config['STATIC_FOLDER'], json_filename)
            
            # 命中缓存（无论是谁上传过这张图）
            if os.path.exists(image_path) and os.path.exists(json_path):
                print(f"⚡ 极速缓存！检测到曾识别过此图纸哈希 [{safe_base_name}]")
                with open(json_path, 'r', encoding='utf-8') as f:
                    ocr_data_json = f.read()
                img = Image.open(image_path)
                orig_w, orig_h = img.size
                image_url = url_for('static', filename=image_filename)
                
                return render_template('index.html', 
                                       image_url=image_url, ocr_data=ocr_data_json,
                                       orig_w=orig_w, orig_h=orig_h,
                                       filename=safe_base_name) # 返回哈希给前端用于历史记录
            
            # 首次识别
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{safe_base_name}.pdf")
            file.save(pdf_path)
            
            try:
                pdf_doc = fitz.open(pdf_path)
                page = pdf_doc.load_page(0) 
                zoom = fitz.Matrix(4, 4)
                pix = page.get_pixmap(matrix=zoom)
                pix.save(image_path)
                image_url = url_for('static', filename=image_filename)
                
                img = Image.open(image_path)
                orig_w, orig_h = img.size
                
                start_t = time.time()
                ocr_result, _ = ocr_engine(image_path)
                
                ocr_data = []
                if ocr_result:
                    for item in ocr_result:
                        box = item[0]
                        text = item[1]
                        xs = [p[0] for p in box]
                        ys = [p[1] for p in box]
                        x, y = min(xs), min(ys)
                        w, h = max(xs) - x, max(ys) - y
                        ocr_data.append({'x': float(x), 'y': float(y), 'w': float(w), 'h': float(h), 'text': str(text)})
                
                ocr_data_json = json.dumps(ocr_data, ensure_ascii=False)
                with open(json_path, 'w', encoding='utf-8') as f:
                    f.write(ocr_data_json)
                    
            except Exception as e:
                print(f"识别错误: {e}")
                return "图纸处理失败，请确保上传的是有效PDF文件", 500
            finally:
                # 阅后即焚PDF原件，节约服务器空间
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                
            return render_template('index.html', 
                                   image_url=image_url, ocr_data=ocr_data_json,
                                   orig_w=orig_w, orig_h=orig_h,
                                   filename=safe_base_name)
            
    # GET 请求（刷新空网页时）
    return render_template('index.html', image_url=image_url, ocr_data=ocr_data_json, orig_w=orig_w, orig_h=orig_h, filename="")

# ==============================
# 📊 高级全息导出 Excel 引擎 (多工作表支持)
# ==============================
@app.route('/api/export', methods=['POST'])
def api_export():
    try:
        payload = request.json
        sheets_data = payload.get('sheets', [])
        
        # 🚀 核心升级：为每次导出生成唯一文件名，防止多用户导出互相覆盖
        unique_export_id = uuid.uuid4().hex[:8]
        export_filename = f'红星号管_工程汇总表_{unique_export_id}.xlsx'
        export_path = os.path.join(app.config['STATIC_FOLDER'], export_filename)

        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"]) 

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        thick_bottom = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='medium'))

        for sheet_info in sheets_data:
            sheet_name = sheet_info.get('sheet_name', '未命名图纸').strip()
            sheet_name = "".join([c for c in sheet_name if c not in r'[]:*?/\ '])[:31]
            if not sheet_name:
                sheet_name = "未命名"

            table_data = sheet_info.get('data', [])

            base_name = sheet_name
            counter = 1
            while sheet_name in wb.sheetnames:
                sheet_name = f"{base_name}_{counter}"
                counter += 1

            ws = wb.create_sheet(title=sheet_name)

            for r_idx, row in enumerate(table_data, 1):
                for c_idx, cell_info in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    val = cell_info.get('val', '')
                    color = cell_info.get('color', '')
                    is_header = cell_info.get('is_header', False)
                    is_group_end = cell_info.get('is_group_end', False)

                    cell.value = val
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                    if is_header:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                        cell.border = thin_border
                    else:
                        cell.font = Font(bold=True) if cell_info.get('bold') else Font()
                        if color and color != 'transparent' and color.startswith('#'):
                            hex_color = color.replace('#', '')
                            if len(hex_color) == 6:
                                cell.fill = PatternFill(start_color="FF"+hex_color, end_color="FF"+hex_color, fill_type="solid")
                        cell.border = thick_bottom if is_group_end else thin_border

            col_widths = [15, 14, 8, 10, 12, 20, 35, 18]
            for i, width in enumerate(col_widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

        if not wb.sheetnames: 
            wb.create_sheet("Empty")

        wb.save(export_path)
        # 前端无需修改，直接接收这个动态 URL
        return jsonify({'url': url_for('static', filename=export_filename)})
    except Exception as e:
        print("❌ 导出发生错误:", e)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    # 删除了单机版的 open_browser
    # 将 host 改为 0.0.0.0，允许公网访问。端口使用云服务商提供的端口 (环境变量)
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
