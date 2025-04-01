from flask import Flask, request, send_file, render_template, Response
import os
import sys
import time
import subprocess
import logging
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import tempfile
import PyPDF2
import shutil
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt  # Thêm Pt cho điều chỉnh font size
from io import BytesIO
from PIL import Image
from docx import Document  # Thêm để đọc file Word

app = Flask(__name__, template_folder='templates')

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Thư mục upload
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# Health check endpoint
@app.route('/health')
def health_check():
    return 'OK', 200


def find_libreoffice():
    possible_paths = [
        'soffice',
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/usr/lib/libreoffice/program/soffice',
    ]

    for path in possible_paths:
        try:
            if os.path.isfile(path) or shutil.which(path):
                result = subprocess.run([path, '--version'],
                                        capture_output=True, text=True, check=False)
                if result.returncode == 0:
                    logger.info(f"Found LibreOffice at: {path}")
                    return path
        except Exception as e:
            logger.warning(f"Error checking LibreOffice at {path}: {e}")

    logger.warning("Using default 'soffice' path")
    return 'soffice'


SOFFICE_PATH = find_libreoffice()
logger.info(f"Using LibreOffice path: {SOFFICE_PATH}")


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=5, delay=1):
    for i in range(retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
        except Exception as e:
            logger.warning(f"Failed to remove {file_path} (attempt {i + 1}): {e}")
            time.sleep(delay)
    return False


def convert_pdf_to_pptx_python(input_path, output_path):
    try:
        # Chuyển PDF sang hình ảnh chất lượng cao
        images = convert_from_path(input_path, dpi=300, fmt='jpeg')

        if not images:
            raise ValueError("No pages found in PDF")

        # Tạo presentation với tỷ lệ phù hợp
        prs = Presentation()

        # Xác định tỷ lệ trang từ trang đầu tiên
        first_page = images[0]
        page_ratio = first_page.width / first_page.height

        # Đặt kích thước slide dựa trên tỷ lệ trang (16:9 hoặc 4:3)
        if abs(page_ratio - 16 / 9) < abs(page_ratio - 4 / 3):
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(5.625)  # Tỷ lệ 16:9
        else:
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)  # Tỷ lệ 4:3

        blank_layout = prs.slide_layouts[6]

        for image in images:
            # Sử dụng bộ nhớ đệm thay vì file tạm
            img_bytes = BytesIO()
            image.save(img_bytes, format='JPEG', quality=95)
            img_bytes.seek(0)

            slide = prs.slides.add_slide(blank_layout)

            # Tính toán kích thước ảnh để giữ tỷ lệ
            img_ratio = image.width / image.height
            slide_ratio = prs.slide_width / prs.slide_height

            if img_ratio > slide_ratio:
                # Ảnh rộng hơn slide - fit theo chiều rộng
                width = prs.slide_width
                height = width / img_ratio
            else:
                # Ảnh cao hơn slide - fit theo chiều cao
                height = prs.slide_height
                width = height * img_ratio

            # Căn giữa ảnh
            left = (prs.slide_width - width) / 2
            top = (prs.slide_height - height) / 2

            slide.shapes.add_picture(img_bytes, left, top, width, height)

        prs.save(output_path)
        return True

    except Exception as e:
        logger.error(f"PDF to PPTX conversion error: {e}")
        return False


def pdf_to_pptx_auto_fit(pdf_path, pptx_path):
    """Chuyển PDF sang PPTX bằng cách trích xuất nội dung text"""
    try:
        # Tạo file docx tạm thời
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            docx_path = temp_docx.name

        # Bước 1: Chuyển PDF sang Word để lấy nội dung
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()

        # Bước 2: Đọc nội dung từ file Word
        doc = Document(docx_path)
        prs = Presentation()

        # Thiết lập kích thước slide (có thể tuỳ chỉnh)
        prs.slide_width = Inches(10)  # Rộng 10 inches
        prs.slide_height = Inches(7.5)  # Cao 7.5 inches

        for para in doc.paragraphs:
            if para.text.strip():  # Bỏ qua đoạn trống
                slide = prs.slides.add_slide(prs.slide_layouts[1])  # Layout "Tiêu đề và nội dung"
                text_frame = slide.shapes[1].text_frame  # Khung nội dung
                text_frame.text = para.text

                # Điều chỉnh cỡ chữ tự động
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(12)  # Cỡ chữ mặc định

        # Xử lý bảng (nếu có)
        for table in doc.tables:
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Layout trống
            cols, rows = len(table.columns), len(table.rows)
            left, top, width, height = Inches(1), Inches(1), Inches(8), Inches(5)  # Vị trí và kích thước
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height).table

            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    table_shape.cell(i, j).text = cell.text

        prs.save(pptx_path)
        return True
    except Exception as e:
        logger.error(f"PDF to PPTX (auto-fit) conversion error: {e}")
        return False
    finally:
        # Dọn dẹp file tạm
        if 'docx_path' in locals() and os.path.exists(docx_path):
            safe_remove(docx_path)


def convert_jpg_to_pdf(input_path, output_path):
    try:
        image = Image.open(input_path)
        # Chuyển sang RGB nếu ảnh ở chế độ CMYK
        if image.mode == 'CMYK':
            image = image.convert('RGB')

        # Tạo PDF mới từ ảnh
        image.save(output_path, "PDF", resolution=100.0)
        return True
    except Exception as e:
        logger.error(f"JPG to PDF conversion error: {e}")
        return False


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    input_path = output_path = None
    try:
        # Kiểm tra input
        if 'file' not in request.files:
            return "No file uploaded", 400

        file = request.files['file']
        if not file or file.filename == '':
            return "No file selected", 400

        if not allowed_file(file.filename):
            return "Invalid file type", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            return "No conversion type selected", 400

        # Chuẩn bị thư mục upload
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Lưu file upload
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved: {input_path}")

        # Xác định loại chuyển đổi
        ext = filename.rsplit('.', 1)[1].lower()
        conversions = {
            'pdf_to_docx': ('pdf', 'docx'),
            'docx_to_pdf': ('docx', 'pdf'),
            'pdf_to_ppt': ('pdf', 'pptx'),
            'ppt_to_pdf': (['ppt', 'pptx'], 'pdf'),
            'jpg_to_pdf': (['jpg', 'jpeg'], 'pdf'),
            'pdf_docx': ('pdf', 'docx') if ext == 'pdf' else ('docx', 'pdf'),
            'pdf_ppt': ('pdf', 'pptx') if ext == 'pdf' else (['ppt', 'pptx'], 'pdf'),
            'image_pdf': (['jpg', 'jpeg'], 'pdf') if ext in ['jpg', 'jpeg'] else ('pdf', 'jpg')
        }

        if conversion_type not in conversions:
            return "Invalid conversion type", 400

        valid_exts, out_ext = conversions[conversion_type]
        if isinstance(valid_exts, list):
            if ext not in valid_exts:
                return "File type mismatch", 400
        elif ext != valid_exts:
            return "File type mismatch", 400

        # Tạo tên file output
        base_name = filename.rsplit('.', 1)[0]
        output_filename = f"converted_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # Thực hiện chuyển đổi
        if conversion_type in ['pdf_to_docx', 'pdf_docx'] and ext == 'pdf':
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        elif conversion_type in ['docx_to_pdf', 'pdf_docx'] and ext == 'docx':
            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], check=True, timeout=60)

            temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
            if os.path.exists(temp_output):
                os.rename(temp_output, output_path)
            else:
                raise Exception("Conversion failed - no output file")

        elif conversion_type in ['pdf_to_ppt', 'pdf_ppt'] and ext == 'pdf':
            # Thử LibreOffice trước
            try:
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--infilter="draw_pdf_import"',
                    '--convert-to', 'pptx',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], check=True, timeout=120)

                temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pptx")
                if os.path.exists(temp_output):
                    os.rename(temp_output, output_path)
                else:
                    raise Exception("LibreOffice conversion failed")

            except Exception as e:
                logger.warning(f"LibreOffice failed, trying python-pptx image method: {e}")
                if not convert_pdf_to_pptx_python(input_path, output_path):
                    logger.warning("Image method failed, trying text extraction method")
                    if not pdf_to_pptx_auto_fit(input_path, output_path):
                        raise Exception("All conversion methods failed")

        elif conversion_type in ['ppt_to_pdf', 'pdf_ppt'] and ext in ['ppt', 'pptx']:
            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], check=True, timeout=60)

            temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
            if os.path.exists(temp_output):
                os.rename(temp_output, output_path)
            else:
                raise Exception("Conversion failed - no output file")

        elif conversion_type in ['jpg_to_pdf', 'image_pdf'] and ext in ['jpg', 'jpeg']:
            if not convert_jpg_to_pdf(input_path, output_path):
                raise Exception("JPG to PDF conversion failed")

        else:
            return "Unsupported conversion", 400

        # Trả về file đã chuyển đổi
        with open(output_path, 'rb') as f:
            file_data = f.read()

        mimetypes = {
            'pdf': 'application/pdf',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg'
        }

        return Response(
            file_data,
            mimetype=mimetypes.get(out_ext, 'application/octet-stream'),
            headers={'Content-Disposition': f'attachment; filename={output_filename}'}
        )

    except Exception as e:
        logger.error(f"Conversion error: {e}")
        return f"Conversion failed: {str(e)}", 500

    finally:
        # Dọn dẹp file tạm
        for path in [input_path, output_path]:
            if path and os.path.exists(path):
                safe_remove(path)


@app.teardown_appcontext
def cleanup(exception=None):
    if not os.path.exists(UPLOAD_FOLDER):
        return

    try:
        now = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                if os.path.isfile(path) and os.path.getmtime(path) < now - 3600:
                    safe_remove(path)
            except Exception as e:
                logger.error(f"Cleanup error for {path}: {e}")
    except Exception as e:
        logger.error(f"Cleanup error: {e}")


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logger.info(f"Starting server on port {port}")
    app.run(host='0.0.0.0', port=port)