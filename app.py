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
from pptx.util import Inches

app = Flask(__name__, template_folder='templates')

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Cấu hình thư mục tạm
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# Health check endpoint
@app.route('/health')
def health_check():
    """Endpoint kiểm tra tình trạng ứng dụng"""
    return 'OK', 200


# Tìm đường dẫn LibreOffice
def find_libreoffice():
    """Tìm đường dẫn thực thi LibreOffice"""
    possible_paths = [
        'soffice',
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/usr/lib/libreoffice/program/soffice',
    ]

    # Trước tiên thử gọi 'soffice' thông qua PATH
    try:
        result = subprocess.run(['soffice', '--version'],
                                capture_output=True, text=True, check=False)
        if result.returncode == 0:
            logger.info(f"Tìm thấy LibreOffice trong PATH")
            return 'soffice'
    except Exception as e:
        logger.warning(f"Không tìm thấy LibreOffice trong PATH: {e}")

    # Sau đó thử các đường dẫn cụ thể
    for path in possible_paths:
        try:
            if os.path.isfile(path):
                # Kiểm tra xem nó có thực sự hoạt động không trước khi trả về
                result = subprocess.run([path, '--version'],
                                        capture_output=True, text=True, check=False)
                if result.returncode == 0:
                    logger.info(f"Tìm thấy LibreOffice hoạt động tại: {path}")
                    return path
            elif shutil.which(path):
                full_path = shutil.which(path)
                # Kiểm tra xem nó có hoạt động không
                result = subprocess.run([full_path, '--version'],
                                        capture_output=True, text=True, check=False)
                if result.returncode == 0:
                    logger.info(f"Tìm thấy LibreOffice hoạt động tại: {full_path}")
                    return full_path
        except Exception as e:
            logger.warning(f"Đã thử LibreOffice tại {path} nhưng gặp lỗi: {e}")
            continue

    # Trở về giá trị mặc định
    logger.warning("Không tìm thấy LibreOffice trong các vị trí dự kiến, sử dụng mặc định 'soffice'")
    return 'soffice'


# Lấy đường dẫn LibreOffice
SOFFICE_PATH = find_libreoffice()
logger.info(f"Using LibreOffice path: {SOFFICE_PATH}")

# Kiểm tra LibreOffice khi khởi động
try:
    subprocess.run([SOFFICE_PATH, '--version'], check=True)
    logger.info("LibreOffice is working correctly")
except Exception as e:
    logger.error(f"LibreOffice check failed: {e}")


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=5, delay=1):
    """Xóa tệp an toàn với số lần thử lại và trì hoãn."""
    if not os.path.exists(file_path):
        logger.debug(f"Tệp không tồn tại, không cần xóa: {file_path}")
        return True

    for i in range(retries):
        try:
            os.remove(file_path)
            logger.debug(f"Đã xóa tệp thành công: {file_path}")
            return True
        except Exception as e:
            logger.warning(f"Không thể xóa tệp {file_path} (lần thử {i + 1}/{retries}): {e}")
            time.sleep(delay)

    logger.error(f"Không thể xóa tệp sau {retries} lần thử: {file_path}")
    return False


def convert_pdf_to_pptx_python(input_path, output_path):
    """Chuyển đổi PDF sang PPTX sử dụng pdf2image và python-pptx"""
    try:
        # Chuyển PDF thành hình ảnh
        images = convert_from_path(input_path, dpi=200)

        # Tạo presentation mới
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]  # Layout trống

        for i, image in enumerate(images):
            # Lưu ảnh tạm
            temp_img_path = os.path.join(UPLOAD_FOLDER, f"temp_{i}.jpg")
            image.save(temp_img_path, 'JPEG')

            # Thêm slide mới
            slide = prs.slides.add_slide(blank_slide_layout)

            # Thêm ảnh vào slide
            left = top = Inches(0)
            slide.shapes.add_picture(temp_img_path, left, top,
                                     width=prs.slide_width,
                                     height=prs.slide_height)

            # Xóa ảnh tạm
            safe_remove(temp_img_path)

        # Lưu file PPTX
        prs.save(output_path)
        return True
    except Exception as e:
        logger.error(f"Lỗi khi chuyển đổi PDF sang PPTX bằng python-pptx: {e}")
        return False


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    input_path = None
    output_path = None
    try:
        if 'file' not in request.files:
            logger.error("No file uploaded")
            return "No file uploaded", 400

        file = request.files['file']
        if not file or file.filename == '':
            logger.error("No file selected")
            return "No file selected", 400

        if not allowed_file(file.filename):
            logger.error("Invalid file type")
            return "Only PDF, DOCX, PPT, PPTX, JPG files are supported", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            logger.error("No conversion type selected")
            return "Please select a conversion type", 400

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        try:
            os.chmod(UPLOAD_FOLDER, 0o755)
        except Exception as e:
            logger.warning(f"Không thể đặt quyền cho thư mục uploads: {e}")

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved at: {input_path}")

        file_extension = filename.rsplit('.', 1)[1].lower()

        # Xác định loại chuyển đổi
        if conversion_type == 'pdf_to_docx' and file_extension == 'pdf':
            actual_conversion = 'pdf_to_docx'
        elif conversion_type == 'docx_to_pdf' and file_extension == 'docx':
            actual_conversion = 'docx_to_pdf'
        elif conversion_type == 'pdf_to_ppt' and file_extension == 'pdf':
            actual_conversion = 'pdf_to_ppt'
        elif conversion_type == 'ppt_to_pdf' and file_extension in ['ppt', 'pptx']:
            actual_conversion = 'ppt_to_pdf'
        elif conversion_type == 'pdf_docx':
            actual_conversion = 'pdf_to_docx' if file_extension == 'pdf' else 'docx_to_pdf'
        elif conversion_type == 'pdf_ppt':
            actual_conversion = 'pdf_to_ppt' if file_extension == 'pdf' else 'ppt_to_pdf'
        else:
            logger.error("Unknown conversion type or file type mismatch")
            return "Unknown conversion type or file type mismatch", 400

        logger.info(f"Detected conversion type: {actual_conversion}")

        # Thực hiện chuyển đổi
        if actual_conversion == 'pdf_to_docx':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.docx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        elif actual_conversion == 'docx_to_pdf':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')

            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], capture_output=True, text=True, check=True, timeout=60)

            if not os.path.exists(actual_output_path):
                logger.error(f"Output file not created: {actual_output_path}")
                return "Error converting DOCX to PDF", 500

            if actual_output_path != output_path:
                os.rename(actual_output_path, output_path)

        elif actual_conversion == 'pdf_to_ppt':
            base_filename = filename.rsplit('.', 1)[0]
            output_filename = f"converted_{base_filename}.pptx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)

            # Thử phương pháp LibreOffice trước
            libreoffice_success = False
            try:
                # Thử chuyển đổi trực tiếp
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--infilter="draw_pdf_import"',
                    '--convert-to', 'pptx',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], capture_output=True, text=True, check=True, timeout=120)

                actual_output_path = os.path.join(UPLOAD_FOLDER, f"{base_filename}.pptx")
                if os.path.exists(actual_output_path):
                    if actual_output_path != output_path:
                        os.rename(actual_output_path, output_path)
                    libreoffice_success = True
                    logger.info("Chuyển đổi thành công bằng LibreOffice")
            except Exception as e:
                logger.warning(f"LibreOffice conversion failed, trying alternative method: {e}")

            # Nếu LibreOffice thất bại, thử phương pháp python-pptx
            if not libreoffice_success:
                logger.info("Trying python-pptx method...")
                if not convert_pdf_to_pptx_python(input_path, output_path):
                    return "Failed to convert PDF to PPTX using all methods", 500

        elif actual_conversion == 'ppt_to_pdf':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')

            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], capture_output=True, text=True, check=True, timeout=60)

            if not os.path.exists(actual_output_path):
                logger.error(f"Output file not created: {actual_output_path}")
                return "Error converting PPT to PDF", 500

            if actual_output_path != output_path:
                os.rename(actual_output_path, output_path)

        else:
            logger.error("Unsupported conversion type")
            return "Unsupported conversion type", 400

        with open(output_path, 'rb') as f:
            file_data = f.read()

        # Xác định MIME type
        if output_filename.endswith('.pdf'):
            mimetype = 'application/pdf'
        elif output_filename.endswith('.docx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif output_filename.endswith('.pptx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        else:
            mimetype = 'application/octet-stream'

        return Response(
            file_data,
            mimetype=mimetype,
            headers={'Content-Disposition': f'attachment; filename={output_filename}'}
        )

    except Exception as e:
        logger.error(f"Lỗi khi chuyển đổi: {str(e)}")
        return f"Lỗi khi chuyển đổi: {str(e)}", 500
    finally:
        if input_path and os.path.exists(input_path):
            safe_remove(input_path)
        if output_path and os.path.exists(output_path):
            safe_remove(output_path)


@app.teardown_appcontext
def cleanup(exception=None):
    """Dọn dẹp các tệp tạm khi context kết thúc."""
    if not os.path.exists(UPLOAD_FOLDER):
        return

    try:
        # Xóa các tệp cũ hơn 1 giờ
        current_time = time.time()
        one_hour_ago = current_time - 3600

        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                if os.path.isfile(file_path):
                    # Kiểm tra thời gian sửa đổi
                    file_mod_time = os.path.getmtime(file_path)
                    if file_mod_time < one_hour_ago:
                        safe_remove(file_path)
            except Exception as e:
                logger.error(f"Lỗi khi dọn dẹp tệp {file_path}: {e}")
    except Exception as e:
        logger.error(f"Lỗi khi dọn dẹp thư mục tạm: {e}")


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logger.info(f"Starting server on port {port}")
    app.run(host='0.0.0.0', port=port)