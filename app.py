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
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image
from docx import Document

app = Flask(__name__, template_folder='templates')

# Cấu hình giới hạn upload file 100MB
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Cấu hình thư mục upload
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# Endpoint kiểm tra tình trạng server
@app.route('/health')
def health_check():
    return 'OK', 200


def find_libreoffice():
    """Tìm đường dẫn đến LibreOffice trên hệ thống"""
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
                    logger.info(f"Tìm thấy LibreOffice tại: {path}")
                    return path
        except Exception as e:
            logger.warning(f"Lỗi khi kiểm tra LibreOffice tại {path}: {e}")

    logger.warning("Sử dụng đường dẫn mặc định 'soffice'")
    return 'soffice'


SOFFICE_PATH = find_libreoffice()
logger.info(f"Sử dụng đường dẫn LibreOffice: {SOFFICE_PATH}")


def allowed_file(filename):
    """Kiểm tra phần mở rộng file có hợp lệ không"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=5, delay=1):
    """Xóa file an toàn với nhiều lần thử"""
    for i in range(retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
        except Exception as e:
            logger.warning(f"Không thể xóa {file_path} (lần thử {i + 1}): {e}")
            time.sleep(delay)
    return False


def get_pdf_page_size(pdf_path):
    """Lấy kích thước trang PDF (đơn vị points)"""
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        page = reader.pages[0]
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        return width, height


def setup_slide_size(prs, pdf_path):
    """Thiết lập kích thước slide dựa trên PDF"""
    try:
        # Lấy kích thước trang PDF (đơn vị points)
        pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)

        # Chuyển đổi từ points sang inches (1 inch = 72 points)
        pdf_width_in = pdf_width_pt / 72
        pdf_height_in = pdf_height_pt / 72

        # Giới hạn kích thước tối đa của PowerPoint (13.33 inches)
        max_size = 13.33

        # Tính tỷ lệ giữa chiều rộng và cao
        ratio = pdf_width_in / pdf_height_in

        # Điều chỉnh kích thước để phù hợp với giới hạn PowerPoint
        if pdf_width_in > max_size or pdf_height_in > max_size:
            if ratio > 1:  # Ngang
                prs.slide_width = Inches(max_size)
                prs.slide_height = Inches(max_size / ratio)
            else:  # Dọc
                prs.slide_height = Inches(max_size)
                prs.slide_width = Inches(max_size * ratio)
        else:
            prs.slide_width = Inches(pdf_width_in)
            prs.slide_height = Inches(pdf_height_in)

        logger.info(f"Thiết lập kích thước slide: {pdf_width_in:.2f} x {pdf_height_in:.2f} inches")
        return prs
    except Exception as e:
        logger.warning(f"Không thể đọc kích thước PDF, sử dụng kích thước mặc định: {e}")
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        return prs


def pdf_to_pptx_auto_fit(pdf_path, pptx_path):
    """Chuyển PDF sang PPTX bằng cách trích xuất nội dung văn bản"""
    try:
        # Tạo presentation
        prs = Presentation()

        # Thiết lập kích thước slide dựa trên PDF
        prs = setup_slide_size(prs, pdf_path)

        # Chuyển PDF sang Word để lấy nội dung
        docx_path = os.path.join(UPLOAD_FOLDER, "temp.docx")
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()

        # Đọc nội dung từ file Word
        doc = Document(docx_path)

        # Thêm nội dung văn bản vào slide
        for para in doc.paragraphs:
            if para.text.strip():  # Bỏ qua đoạn trống
                slide = prs.slides.add_slide(prs.slide_layouts[1])  # Layout "Tiêu đề và nội dung"
                text_frame = slide.shapes[1].text_frame
                text_frame.text = para.text

                # Tự động điều chỉnh kích thước font
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(12)

        # Xử lý bảng biểu nếu có
        for table in doc.tables:
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Layout trống
            cols, rows = len(table.columns), len(table.rows)
            left, top, width, height = Inches(1), Inches(1), Inches(8), Inches(5)
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height).table

            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    table_shape.cell(i, j).text = cell.text

        prs.save(pptx_path)
        safe_remove(docx_path)
        return True
    except Exception as e:
        logger.warning(f"Lỗi chuyển đổi PDF sang PPTX (phương pháp văn bản): {e}")
        return False


def _convert_pdf_to_pptx_images(input_path, output_path):
    """Chuyển PDF sang PPTX bằng cách chuyển từng trang thành hình ảnh"""
    try:
        # Chuyển PDF thành các hình ảnh chất lượng cao
        images = convert_from_path(input_path, dpi=300, fmt='jpeg')

        if not images:
            raise ValueError("Không tìm thấy trang nào trong PDF")

        # Tạo presentation
        prs = Presentation()

        # Thiết lập kích thước slide dựa trên PDF
        prs = setup_slide_size(prs, input_path)

        blank_layout = prs.slide_layouts[6]  # Layout trống

        for image in images:
            # Sử dụng bộ đệm trong bộ nhớ thay vì file tạm
            img_bytes = BytesIO()
            image.save(img_bytes, format='JPEG', quality=95)
            img_bytes.seek(0)

            slide = prs.slides.add_slide(blank_layout)

            # Tính toán kích thước hình ảnh để giữ nguyên tỷ lệ
            img_ratio = image.width / image.height
            slide_ratio = prs.slide_width / prs.slide_height

            if img_ratio > slide_ratio:
                # Hình ảnh rộng hơn slide - fit theo chiều rộng
                width = prs.slide_width
                height = width / img_ratio
            else:
                # Hình ảnh cao hơn slide - fit theo chiều cao
                height = prs.slide_height
                width = height * img_ratio

            # Căn giữa hình ảnh
            left = (prs.slide_width - width) / 2
            top = (prs.slide_height - height) / 2

            slide.shapes.add_picture(img_bytes, left, top, width, height)

        prs.save(output_path)
        return True

    except Exception as e:
        logger.warning(f"Lỗi chuyển đổi PDF sang PPTX (phương pháp hình ảnh): {e}")
        return False


def convert_pdf_to_pptx_python(input_path, output_path):
    """Chuyển PDF sang PPTX sử dụng cả 2 phương pháp"""
    # Thử phương pháp hình ảnh trước
    if _convert_pdf_to_pptx_images(input_path, output_path):
        return True

    # Nếu không thành công, thử phương pháp văn bản
    if pdf_to_pptx_auto_fit(input_path, output_path):
        return True

    return False


def convert_jpg_to_pdf(input_path, output_path):
    """Chuyển đổi JPG sang PDF"""
    try:
        image = Image.open(input_path)
        # Chuyển sang RGB nếu ảnh ở chế độ CMYK
        if image.mode == 'CMYK':
            image = image.convert('RGB')

        # Tạo PDF mới từ hình ảnh
        image.save(output_path, "PDF", resolution=100.0)
        return True
    except Exception as e:
        logger.error(f"Lỗi chuyển đổi JPG sang PDF: {e}")
        return False


@app.route('/')
def index():
    """Trang chủ hiển thị form upload"""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    """Xử lý chuyển đổi file"""
    input_path = output_path = None
    try:
        # Kiểm tra file upload
        if 'file' not in request.files:
            return "Không có file được tải lên", 400

        file = request.files['file']
        if not file or file.filename == '':
            return "Không có file được chọn", 400

        if not allowed_file(file.filename):
            return "Loại file không hợp lệ", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            return "Không chọn loại chuyển đổi", 400

        # Chuẩn bị thư mục upload
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Lưu file upload
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File đã lưu: {input_path}")

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
            return "Loại chuyển đổi không hợp lệ", 400

        valid_exts, out_ext = conversions[conversion_type]
        if isinstance(valid_exts, list):
            if ext not in valid_exts:
                return "Loại file không phù hợp", 400
        elif ext != valid_exts:
            return "Loại file không phù hợp", 400

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
                raise Exception("Chuyển đổi thất bại - không có file output")

        elif conversion_type in ['pdf_to_ppt', 'pdf_ppt'] and ext == 'pdf':
            # Thử dùng LibreOffice trước
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
                    raise Exception("LibreOffice chuyển đổi thất bại")

            except Exception as e:
                logger.warning(f"LibreOffice thất bại, chuyển sang python-pptx: {e}")
                if not convert_pdf_to_pptx_python(input_path, output_path):
                    raise Exception("Tất cả phương pháp chuyển đổi đều thất bại")

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
                raise Exception("Chuyển đổi thất bại - không có file output")

        elif conversion_type in ['jpg_to_pdf', 'image_pdf'] and ext in ['jpg', 'jpeg']:
            if not convert_jpg_to_pdf(input_path, output_path):
                raise Exception("Chuyển đổi JPG sang PDF thất bại")

        else:
            return "Chuyển đổi không được hỗ trợ", 400

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
        logger.error(f"Lỗi chuyển đổi: {e}")
        return f"Chuyển đổi thất bại: {str(e)}", 500

    finally:
        # Dọn dẹp file tạm
        for path in [input_path, output_path]:
            if path and os.path.exists(path):
                safe_remove(path)


@app.teardown_appcontext
def cleanup(exception=None):
    """Dọn dẹp các file cũ trong thư mục upload"""
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
                logger.error(f"Lỗi khi dọn dẹp {path}: {e}")
    except Exception as e:
        logger.error(f"Lỗi khi dọn dẹp: {e}")


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logger.info(f"Khởi động server trên cổng {port}")
    app.run(host='0.0.0.0', port=port)