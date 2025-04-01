from flask import Flask, request, send_file, render_template, Response, flash
import os
import time
import subprocess
import logging
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import img2pdf  # Để chuyển JPG sang PDF
import fitz  # PyMuPDF để chuyển PDF sang JPG

app = Flask(__name__, template_folder='templates')
app.secret_key = os.urandom(24)  # Cần thiết cho flash messages

# Cấu hình logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Cấu hình thư mục tạm
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'xls', 'xlsx', 'jpg', 'jpeg'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB giới hạn

# Cấu hình giới hạn kích thước file
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=5, delay=1):
    for _ in range(retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            return True
        except Exception as e:
            logger.warning(f"Không thể xóa file {file_path}: {e}")
            time.sleep(delay)
    return False


@app.route('/')
def index():
    return render_template('index.html')


# Thêm hàm để dọn dẹp file tạm định kỳ
def cleanup_temp_files(max_age_seconds=3600):  # Mặc định 1 giờ
    """Xóa các file tạm cũ hơn max_age_seconds"""
    if not os.path.exists(UPLOAD_FOLDER):
        return

    current_time = time.time()
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.isfile(file_path):
            file_age = current_time - os.path.getmtime(file_path)
            if file_age > max_age_seconds:
                safe_remove(file_path)
                logger.info(f"Đã xóa file tạm cũ: {file_path}")


@app.route('/convert', methods=['POST'])
def convert_file():
    input_path = None
    output_path = None
    try:
        # Dọn dẹp file tạm cũ mỗi khi có yêu cầu chuyển đổi
        cleanup_temp_files()

        if 'file' not in request.files:
            logger.error("No file uploaded")
            return "Không có file nào được tải lên", 400

        file = request.files['file']
        if not file or file.filename == '':
            logger.error("No file selected")
            return "Không có file nào được chọn", 400

        if not allowed_file(file.filename):
            logger.error("Invalid file type, only PDF, DOCX, PPT, Excel, JPG supported")
            return "Chỉ hỗ trợ file PDF, DOCX, PPT, Excel, JPG", 400

        # Kiểm tra kích thước file
        if request.content_length and request.content_length > MAX_CONTENT_LENGTH:
            logger.error(f"File too large: {request.content_length} bytes")
            return "File quá lớn. Giới hạn là 16MB.", 413

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            logger.error("No conversion type selected")
            return "Vui lòng chọn loại chuyển đổi", 400

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved at: {input_path}")

        # Phần còn lại của hàm convert_file giữ nguyên...
        # [Giữ nguyên code chuyển đổi hiện tại]

    except subprocess.TimeoutExpired:
        logger.error("Conversion timed out")
        return "Lỗi: Quá trình chuyển đổi mất quá nhiều thời gian", 500
    except Exception as e:
        logger.error(f"Error during conversion: {str(e)}")
        return f"Lỗi trong quá trình chuyển đổi: {str(e)}", 500
    finally:
        if input_path and os.path.exists(input_path):
            safe_remove(input_path)
        if output_path and os.path.exists(output_path):
            safe_remove(output_path)


# Thêm một route để dọn dẹp thủ công
@app.route('/cleanup', methods=['GET'])
def manual_cleanup():
    if request.remote_addr == '127.0.0.1':  # Chỉ cho phép từ localhost
        cleanup_temp_files(0)  # Xóa tất cả file
        return "Đã dọn dẹp thành công"
    return "Không được phép", 403


@app.teardown_appcontext
def cleanup(exception=None):
    if os.path.exists(UPLOAD_FOLDER):
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                logger.error(f"Error cleaning up file {file_path}: {str(e)}")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5003)))