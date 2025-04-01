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
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
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

    for path in possible_paths:
        try:
            if os.path.isfile(path):
                logger.info(f"Found LibreOffice at: {path}")
                return path
            elif shutil.which(path):
                logger.info(f"Found LibreOffice in PATH: {shutil.which(path)}")
                return shutil.which(path)
        except Exception:
            continue

    logger.warning("LibreOffice not found in expected locations")
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

            try:
                # Thử nhiều phương pháp chuyển đổi
                conversion_methods = [
                    # Phương pháp 1: Chuyển trực tiếp PDF sang PPTX
                    lambda: subprocess.run([
                        SOFFICE_PATH,
                        '--headless',
                        '--convert-to', 'pptx',
                        '--outdir', UPLOAD_FOLDER,
                        input_path
                    ], capture_output=True, text=True, check=True, timeout=60),

                    # Phương pháp 2: Chuyển PDF sang HTML, rồi từ HTML sang PPTX
                    lambda: subprocess.run([
                        SOFFICE_PATH,
                        '--headless',
                        '--convert-to', 'html',
                        '--outdir', UPLOAD_FOLDER,
                        input_path
                    ], capture_output=True, text=True, check=True, timeout=60) or
                            subprocess.run([
                                SOFFICE_PATH,
                                '--headless',
                                '--convert-to', 'pptx',
                                '--outdir', UPLOAD_FOLDER,
                                os.path.join(UPLOAD_FOLDER, f"{base_filename}.html")
                            ], capture_output=True, text=True, check=True, timeout=60)
                ]

                # Biến để theo dõi kết quả chuyển đổi
                conversion_successful = False

                # Thử từng phương pháp chuyển đổi
                for method in conversion_methods:
                    try:
                        # Tìm tất cả các file có thể được tạo ra
                        before_files = set(os.listdir(UPLOAD_FOLDER))

                        # Thực hiện chuyển đổi
                        result = method()
                        logger.info(f"Conversion method result stdout: {result.stdout}")
                        logger.info(f"Conversion method result stderr: {result.stderr}")

                        # Tìm file mới được tạo ra
                        after_files = set(os.listdir(UPLOAD_FOLDER))
                        new_files = after_files - before_files

                        # Tìm file PPTX
                        pptx_files = [f for f in new_files if f.endswith('.pptx')]

                        if pptx_files:
                            # Đổi tên file PPTX đầu tiên tìm được
                            first_pptx = pptx_files[0]
                            full_pptx_path = os.path.join(UPLOAD_FOLDER, first_pptx)

                            if full_pptx_path != output_path:
                                os.rename(full_pptx_path, output_path)

                            conversion_successful = True
                            break
                    except Exception as conversion_error:
                        logger.warning(f"Conversion method failed: {conversion_error}")
                        continue

                # Kiểm tra kết quả cuối cùng
                if not conversion_successful:
                    logger.error("Không thể chuyển đổi PDF sang PPTX bằng bất kỳ phương pháp nào")
                    return "Lỗi: Không thể chuyển đổi PDF sang PPTX", 500

                # Xóa các file HTML tạm
                for file in os.listdir(UPLOAD_FOLDER):
                    if file.endswith('.html'):
                        os.remove(os.path.join(UPLOAD_FOLDER, file))

            except Exception as e:
                logger.error(f"Lỗi chuyển đổi PDF sang PPTX: {str(e)}")
                return f"Lỗi khi chuyển đổi: {str(e)}", 500

        elif actual_conversion == 'ppt_to_pdf':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')

            subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], capture_output=True, text=True, check=True, timeout=60)

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
    if os.path.exists(UPLOAD_FOLDER):
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception:
                pass


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logger.info(f"Starting server on port {port}")
    app.run(host='0.0.0.0', port=port)