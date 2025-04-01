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

        # Thực hiện chuyển đổi dựa trên loại
        output_filename = f"converted_{os.path.splitext(filename)[0]}"

        if conversion_type == 'pdf_to_docx':
            if not filename.lower().endswith('.pdf'):
                return "Cần file PDF để chuyển sang DOCX", 400
            output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename}.docx")
            # Chuyển đổi PDF sang DOCX
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
            content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

        elif conversion_type == 'docx_to_pdf':
            if not filename.lower().endswith('.docx'):
                return "Cần file DOCX để chuyển sang PDF", 400
            output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename}.pdf")
            # Chuyển đổi DOCX sang PDF sử dụng libreoffice
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER, input_path
            ], timeout=60, check=True)
            # Sửa lại output_path vì libreoffice lưu tên file khác
            output_path = os.path.join(UPLOAD_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
            content_type = 'application/pdf'

        elif conversion_type == 'pdf_to_jpg':
            if not filename.lower().endswith('.pdf'):
                return "Cần file PDF để chuyển sang JPG", 400
            # Tạo thư mục tạm để chứa các file JPG
            temp_img_dir = os.path.join(UPLOAD_FOLDER, f"{os.path.splitext(filename)[0]}_imgs")
            os.makedirs(temp_img_dir, exist_ok=True)

            # Mở file PDF
            pdf_doc = fitz.open(input_path)
            # Dùng trang đầu tiên làm ví dụ (có thể mở rộng để xử lý nhiều trang)
            page = pdf_doc[0]
            pix = page.get_pixmap(dpi=300)  # Độ phân giải cao

            output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename}.jpg")
            pix.save(output_path)
            content_type = 'image/jpeg'

        elif conversion_type == 'jpg_to_pdf':
            if not (filename.lower().endswith('.jpg') or filename.lower().endswith('.jpeg')):
                return "Cần file JPG để chuyển sang PDF", 400
            output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename}.pdf")

            # Chuyển JPG sang PDF
            with open(output_path, "wb") as f:
                f.write(img2pdf.convert(input_path))
            content_type = 'application/pdf'

        elif conversion_type in ['ppt_to_pdf', 'excel_to_pdf']:
            ext = '.pptx' if conversion_type == 'ppt_to_pdf' else '.xlsx'
            if not (filename.lower().endswith('.ppt') or filename.lower().endswith('.pptx') or
                    filename.lower().endswith('.xls') or filename.lower().endswith('.xlsx')):
                return f"Cần file {ext[1:].upper()} để chuyển sang PDF", 400

            output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename}.pdf")
            # Chuyển đổi PPT/Excel sang PDF sử dụng libreoffice
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER, input_path
            ], timeout=60, check=True)
            # Sửa lại output_path vì libreoffice lưu tên file khác
            output_path = os.path.join(UPLOAD_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
            content_type = 'application/pdf'

        else:
            return f"Loại chuyển đổi {conversion_type} không được hỗ trợ", 400

        if not os.path.exists(output_path):
            return "Chuyển đổi thất bại, không tạo được file đầu ra", 500

        # Trả về file đã chuyển đổi
        return send_file(
            output_path,
            as_attachment=True,
            download_name=os.path.basename(output_path),
            mimetype=content_type
        )

    except subprocess.TimeoutExpired:
        logger.error("Conversion timed out")
        return "Lỗi: Quá trình chuyển đổi mất quá nhiều thời gian", 500
    except Exception as e:
        logger.error(f"Error during conversion: {str(e)}")
        return f"Lỗi trong quá trình chuyển đổi: {str(e)}", 500
    finally:
        # Đảm bảo xóa các file tạm sau khi xử lý xong
        try:
            if input_path and os.path.exists(input_path):
                safe_remove(input_path)
        except Exception as e:
            logger.error(f"Error cleaning up input file: {str(e)}")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5003)))