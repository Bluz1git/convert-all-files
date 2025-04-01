from flask import Flask, request, send_file, render_template, Response
import os
import time
import subprocess
import logging
# Xóa import pandas
# import pandas as pd
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import tempfile
import PyPDF2
import shutil
# Nâng cấp logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # Log ra console
        logging.FileHandler('app.log', encoding='utf-8')  # Log ra file
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__, template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Giới hạn file 16MB
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')

# Cấu hình logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Cấu hình thư mục tạm
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# Tìm đường dẫn LibreOffice
def find_libreoffice():
    """Find the LibreOffice executable path"""
    # Thứ tự ưu tiên tìm kiếm LibreOffice
    possible_paths = [
        'soffice',  # If in PATH
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/usr/lib/libreoffice/program/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]

    # Kiểm tra từng đường dẫn
    for path in possible_paths:
        try:
            # Kiểm tra nếu đường dẫn tuyệt đối tồn tại
            if os.path.isfile(path):
                logger.info(f"Found LibreOffice at: {path}")
                return path
            # Kiểm tra nếu lệnh có trong PATH
            elif shutil.which(path):
                logger.info(f"Found LibreOffice in PATH: {shutil.which(path)}")
                return shutil.which(path)
        except Exception as e:
            continue

    # Thông báo nếu không tìm thấy
    logger.warning("LibreOffice not found in expected locations")
    return 'soffice'  # Sử dụng giá trị mặc định nếu không tìm thấy


# Lấy đường dẫn LibreOffice
SOFFICE_PATH = find_libreoffice()
logger.info(f"Using LibreOffice path: {SOFFICE_PATH}")


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
            logger.error("Invalid file type, only PDF, DOCX, PPT, PPTX, JPG supported")
            return "Only PDF, DOCX, PPT, PPTX, JPG files are supported", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            logger.error("No conversion type selected")
            return "Please select a conversion type", 400

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        # Đảm bảo thư mục có đủ quyền
        try:
            os.chmod(UPLOAD_FOLDER, 0o755)
        except Exception as e:
            logger.warning(f"Không thể đặt quyền cho thư mục uploads: {e}")

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved at: {input_path}")

        # Tự động nhận diện định dạng file và xác định loại chuyển đổi
        file_extension = filename.rsplit('.', 1)[1].lower()

        # Xác định loại chuyển đổi dựa trên định dạng đầu vào và loại chuyển đổi được chọn
        if conversion_type == 'pdf_to_docx' and file_extension == 'pdf':
            actual_conversion = 'pdf_to_docx'
        elif conversion_type == 'docx_to_pdf' and file_extension == 'docx':
            actual_conversion = 'docx_to_pdf'
        elif conversion_type == 'pdf_to_ppt' and file_extension == 'pdf':
            actual_conversion = 'pdf_to_ppt'
        elif conversion_type == 'ppt_to_pdf' and file_extension in ['ppt', 'pptx']:
            actual_conversion = 'ppt_to_pdf'
        elif conversion_type == 'pdf_docx':
            if file_extension == 'pdf':
                actual_conversion = 'pdf_to_docx'
            elif file_extension == 'docx':
                actual_conversion = 'docx_to_pdf'
            else:
                return "File type does not match conversion type PDF ↔ DOCX", 400
        elif conversion_type == 'pdf_ppt':
            if file_extension == 'pdf':
                actual_conversion = 'pdf_to_ppt'
            elif file_extension in ['ppt', 'pptx']:
                actual_conversion = 'ppt_to_pdf'
            else:
                return "File type does not match conversion type PDF ↔ PPT", 400
        else:
            logger.error("Unknown conversion type or file type mismatch")
            return "Unknown conversion type or file type mismatch", 400

        logger.info(f"Detected conversion type: {actual_conversion}")

        # PDF to DOCX conversion
        if actual_conversion == 'pdf_to_docx':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.docx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            logger.info(f"Starting PDF to DOCX conversion: {input_path} -> {output_path}")
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
            logger.info("PDF to DOCX conversion completed")

        # DOCX to PDF conversion
        elif actual_conversion == 'docx_to_pdf':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
            logger.info(f"Starting DOCX to PDF conversion: {input_path} -> {output_path}")

            # Kiểm tra LibreOffice
            try:
                # Kiểm tra phiên bản LibreOffice
                soffice_check = subprocess.run([SOFFICE_PATH, '--version'],
                                               capture_output=True, text=True,
                                               check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            except FileNotFoundError:
                logger.error(f"LibreOffice executable not found at {SOFFICE_PATH}")
                return "Error: LibreOffice executable not found", 500

            # Thực hiện chuyển đổi
            try:
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], capture_output=True, text=True, check=True, timeout=60)

                logger.info(f"soffice stdout: {result.stdout}")
                if result.stderr:
                    logger.warning(f"soffice stderr: {result.stderr}")

                if not os.path.exists(actual_output_path):
                    logger.error(f"Output file not created: {actual_output_path}")
                    return "Error converting DOCX to PDF", 500

                if actual_output_path != output_path:
                    os.rename(actual_output_path, output_path)
                    logger.info(f"Renamed file from {actual_output_path} to {output_path}")

                logger.info("DOCX to PDF conversion completed")
            except Exception as e:
                logger.error(f"Error in DOCX to PDF conversion: {str(e)}")
                return f"Error converting DOCX to PDF: {str(e)}", 500

        # PDF to PPT conversion
        elif actual_conversion == 'pdf_to_ppt':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pptx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            logger.info(f"Starting PDF to PPT conversion: {input_path} -> {output_path}")

            try:
                # Use LibreOffice for PDF to PPT conversion
                # First, convert PDF to an intermediate format that LibreOffice handles well
                temp_path = os.path.join(UPLOAD_FOLDER, f"temp_{filename.rsplit('.', 1)[0]}.html")

                # Use LibreOffice to convert PDF to HTML first (better preservation of layout)
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'html',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], capture_output=True, text=True, check=True, timeout=60)

                # Check if intermediate HTML file was created
                if not os.path.exists(temp_path):
                    logger.error(f"Intermediate file not created: {temp_path}")
                    return "Error converting PDF to PPT: Intermediate conversion failed", 500

                # Now convert HTML to PPTX
                actual_output_path = os.path.join(UPLOAD_FOLDER, f"temp_{filename.rsplit('.', 1)[0]}.pptx")
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pptx',
                    '--outdir', UPLOAD_FOLDER,
                    temp_path
                ], capture_output=True, text=True, check=True, timeout=60)

                # Check if final PPTX file was created
                if not os.path.exists(actual_output_path):
                    logger.error(f"Output file not created: {actual_output_path}")
                    return "Error converting PDF to PPT", 500

                # Rename to final output path
                if actual_output_path != output_path:
                    os.rename(actual_output_path, output_path)
                    logger.info(f"Renamed file from {actual_output_path} to {output_path}")

                # Clean up temp files
                if os.path.exists(temp_path):
                    os.remove(temp_path)

                logger.info("PDF to PPT conversion completed")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice conversion error: {str(e)}")
                return f"Error converting PDF to PPT: {str(e)}", 500
            except Exception as e:
                logger.error(f"Error in PDF to PPT conversion: {str(e)}")
                return f"Error converting PDF to PPT: {str(e)}", 500

        # PPT to PDF conversion
        elif actual_conversion == 'ppt_to_pdf':
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
            logger.info(f"Starting PPT to PDF conversion: {input_path} -> {output_path}")

            try:
                # Use LibreOffice for PPT to PDF conversion
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], capture_output=True, text=True, check=True, timeout=60)

                logger.info(f"soffice stdout: {result.stdout}")
                if result.stderr:
                    logger.warning(f"soffice stderr: {result.stderr}")

                if not os.path.exists(actual_output_path):
                    logger.error(f"Output file not created: {actual_output_path}")
                    return "Error converting PPT to PDF", 500

                if actual_output_path != output_path:
                    os.rename(actual_output_path, output_path)
                    logger.info(f"Renamed file from {actual_output_path} to {output_path}")

                logger.info("PPT to PDF conversion completed")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice conversion error: {str(e)}")
                return f"Error converting PPT to PDF: {str(e)}", 500

        else:
            logger.error("Unsupported conversion type")
            return "Unsupported conversion type", 400

        with open(output_path, 'rb') as f:
            file_data = f.read()
        logger.info(f"Output file read: {output_path}")

        safe_remove(input_path)
        safe_remove(output_path)

        # Xác định MIME type dựa vào định dạng output
        if output_filename.endswith('.pdf'):
            mimetype = 'application/pdf'
        elif output_filename.endswith('.docx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif output_filename.endswith('.pptx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        elif output_filename.endswith('.ppt'):
            mimetype = 'application/vnd.ms-powerpoint'
        else:
            mimetype = 'application/octet-stream'

        return Response(
            file_data,
            mimetype=mimetype,
            headers={'Content-Disposition': f'attachment; filename={output_filename}'}
        )

    except subprocess.TimeoutExpired:
        logger.error("Conversion timed out")
        return "Error: Conversion took too long", 500
    except Exception as e:
        logger.error(f"Error during conversion: {str(e)}")
        return f"Error during conversion: {str(e)}", 500
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
@app.before_request
def log_request_info():
    logger.info('Request headers: %s', request.headers)
    logger.info('Request method: %s', request.method)

@app.after_request
def add_header(response):
    # Thêm các headers để tăng performance và bảo mật
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    return response

# Xử lý lỗi tập trung
@app.errorhandler(413)
def request_entity_too_large(error):
    logger.error('File quá lớn: %s', error)
    return 'File quá lớn. Giới hạn tối đa 16MB.', 413


@app.before_request
def log_request_info():
    logger.info('Request headers: %s', request.headers)
    logger.info('Request method: %s', request.method)

@app.after_request
def add_header(response):
    # Thêm các headers để tăng performance và bảo mật
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    return response

# Xử lý lỗi tập trung
@app.errorhandler(413)
def request_entity_too_large(error):
    logger.error('File quá lớn: %s', error)
    return 'File quá lớn. Giới hạn tối đa 16MB.', 413