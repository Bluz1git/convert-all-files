# --- START OF FILE app.py ---

from flask import Flask, request, send_file, render_template, Response, jsonify, url_for
import os
import sys
import time
import subprocess
import logging
from werkzeug.utils import secure_filename
# pdf2docx is no longer the primary method for PDF to DOCX, but keep import if needed as fallback (optional)
# from pdf2docx import Converter
import tempfile
import PyPDF2  # Used for getting PDF page size
import shutil
from pdf2image import convert_from_path  # Used for PDF -> PPTX image method
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image  # Used for JPG -> PDF and getting image dimensions

app = Flask(__name__, template_folder='templates', static_folder='static')  # Thêm static_folder

# Cấu hình giới hạn upload file 100MB
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] %(message)s',  # Added funcName
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Cấu hình thư mục upload
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# --- Helper Functions ---

def find_libreoffice():
    """Tìm đường dẫn đến LibreOffice trên hệ thống"""
    possible_paths = [
        'soffice',  # Check PATH first
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # macOS path
        '/usr/lib/libreoffice/program/soffice',
        # Add paths for Windows if needed, though soffice in PATH should work
        # 'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        # 'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]

    for path in possible_paths:
        try:
            # Use shutil.which to find executable in PATH first
            resolved_path = shutil.which(path)
            if resolved_path and os.path.isfile(resolved_path):
                # Test if it's actually LibreOffice by running --version
                # Use a short timeout to avoid hanging
                result = subprocess.run([resolved_path, '--version'],
                                        capture_output=True, text=True, check=False, timeout=10)
                if result.returncode == 0 and 'LibreOffice' in result.stdout:
                    logger.info(f"Tìm thấy LibreOffice tại (qua which): {resolved_path}")
                    return resolved_path
            # If not found via which, check the direct path (useful if not in PATH)
            elif os.path.isfile(path):
                result = subprocess.run([path, '--version'],
                                        capture_output=True, text=True, check=False, timeout=10)
                if result.returncode == 0 and 'LibreOffice' in result.stdout:
                    logger.info(f"Tìm thấy LibreOffice tại (đường dẫn trực tiếp): {path}")
                    return path
        except FileNotFoundError:
            logger.debug(f"Không tìm thấy LibreOffice tại {path} (hoặc không trong PATH)")
        except subprocess.TimeoutExpired:
            logger.warning(f"Kiểm tra LibreOffice tại {path} bị timeout (10s).")
        except Exception as e:
            logger.warning(f"Lỗi khi kiểm tra LibreOffice tại {path}: {e}")

    logger.error("KHÔNG TÌM THẤY LibreOffice thực thi hợp lệ. Chuyển đổi dựa trên LibreOffice sẽ thất bại.")
    return None  # Return None if not found


SOFFICE_PATH = find_libreoffice()
if SOFFICE_PATH:
    logger.info(f"Sử dụng đường dẫn LibreOffice: {SOFFICE_PATH}")
else:
    # Log error here, but the check within the route will handle user feedback
    logger.error("Không thể tìm thấy LibreOffice khi khởi động. Các chức năng chuyển đổi DOCX/PPT sẽ bị ảnh hưởng.")


def allowed_file(filename):
    """Kiểm tra phần mở rộng file có hợp lệ không"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=3, delay=0.5):
    """Xóa file hoặc thư mục an toàn với nhiều lần thử"""
    if not file_path:
        return True

    is_dir = os.path.isdir(file_path)

    if not os.path.exists(file_path):
        # logger.debug(f"{'Thư mục' if is_dir else 'File'} không tồn tại, không cần xóa: {file_path}")
        return True  # Already gone

    for i in range(retries):
        try:
            if is_dir:
                shutil.rmtree(file_path)
                logger.info(f"Đã xóa thư mục tạm: {file_path}")
            else:
                os.remove(file_path)
                logger.info(f"Đã xóa file tạm: {file_path}")
            return True
        except PermissionError as pe:
            logger.warning(
                f"Không có quyền xóa {'thư mục' if is_dir else 'file'} {file_path} (lần thử {i + 1}/{retries}): {pe}. Đang đợi {delay}s...")
            time.sleep(delay)
            delay *= 1.5  # Slightly increase delay
        except FileNotFoundError:
            logger.info(f"{'Thư mục' if is_dir else 'File'} {file_path} đã được xóa (lần thử {i + 1}/{retries}).")
            return True  # Removed by another process?
        except OSError as oe:  # Catch OSError which might include "Directory not empty" on Windows
            logger.warning(
                f"Lỗi OS khi xóa {'thư mục' if is_dir else 'file'} {file_path} (lần thử {i + 1}/{retries}): {oe}. Đang đợi {delay}s...")
            time.sleep(delay)
            delay *= 1.5
        except Exception as e:
            logger.warning(
                f"Không thể xóa {'thư mục' if is_dir else 'file'} {file_path} (lần thử {i + 1}/{retries}): {e}")
            time.sleep(delay)

    logger.error(f"Xóa {'thư mục' if is_dir else 'file'} {file_path} thất bại sau {retries} lần thử.")
    return False


def get_pdf_page_size(pdf_path):
    """Lấy kích thước trang PDF đầu tiên (đơn vị points)"""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if not reader.pages:
                logger.warning(f"PDF không có trang nào: {pdf_path}")
                return None, None
            # Lấy trang đầu tiên
            page = reader.pages[0]
            # mediabox là kích thước vật lý của trang
            mediabox = page.mediabox
            if mediabox:
                width = float(mediabox.width)
                height = float(mediabox.height)
                logger.debug(f"Kích thước PDF (trang 1): {width:.2f} x {height:.2f} points")
                return width, height
            else:
                logger.warning(f"Không tìm thấy mediabox cho trang đầu tiên trong {pdf_path}")
                return None, None
    except FileNotFoundError:
        logger.error(f"File PDF không tồn tại khi lấy kích thước: {pdf_path}")
        return None, None
    except Exception as e:
        # Log full traceback for PDF parsing errors
        logger.error(f"Lỗi khi đọc kích thước PDF {pdf_path}: {e}", exc_info=True)
        return None, None


def setup_slide_size(prs, pdf_path):
    """Thiết lập kích thước slide dựa trên PDF"""
    pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)

    # Kích thước slide mặc định (tương đương 16:9)
    default_width_in = 13.333  # Inches for 16:9 widescreen
    default_height_in = 7.5  # Inches for 16:9 widescreen

    if pdf_width_pt is None or pdf_height_pt is None or pdf_width_pt <= 0 or pdf_height_pt <= 0:
        logger.warning(
            "Không thể đọc kích thước PDF hoặc kích thước không hợp lệ, sử dụng kích thước slide mặc định (16:9).")
        prs.slide_width = Inches(default_width_in)
        prs.slide_height = Inches(default_height_in)
        return prs

    try:
        pdf_width_in = pdf_width_pt / 72.0
        pdf_height_in = pdf_height_pt / 72.0

        # Giới hạn kích thước tối đa cho slide trong PowerPoint (thường là 56 inches)
        max_slide_dim = 56.0

        # Kiểm tra và điều chỉnh nếu kích thước PDF quá lớn
        if pdf_width_in > max_slide_dim or pdf_height_in > max_slide_dim:
            logger.warning(
                f"Kích thước gốc PDF ({pdf_width_in:.2f}x{pdf_height_in:.2f} in) vượt giới hạn {max_slide_dim} in.")
            ratio = pdf_width_in / pdf_height_in
            if pdf_width_in >= pdf_height_in:  # Landscape or square
                final_width = max_slide_dim
                final_height = max_slide_dim / ratio
            else:  # Portrait
                final_height = max_slide_dim
                final_width = max_slide_dim * ratio

            # Ensure dimensions are not zero after scaling
            if final_width <= 0 or final_height <= 0:
                raise ValueError("Kích thước slide sau khi điều chỉnh không hợp lệ.")

            prs.slide_width = Inches(final_width)
            prs.slide_height = Inches(final_height)
            logger.info(f"Điều chỉnh kích thước slide thành: {final_width:.2f}x{final_height:.2f} in")
        else:
            # Kích thước PDF nằm trong giới hạn, sử dụng trực tiếp
            prs.slide_width = Inches(pdf_width_in)
            prs.slide_height = Inches(pdf_height_in)
            logger.info(f"Thiết lập kích thước slide theo PDF: {pdf_width_in:.2f} x {pdf_height_in:.2f} inches")

        return prs

    except ValueError as ve:  # Catch specific error from calculation
        logger.error(f"Lỗi giá trị khi tính toán kích thước slide: {ve}. Sử dụng mặc định.")
        prs.slide_width = Inches(default_width_in)
        prs.slide_height = Inches(default_height_in)
        return prs
    except Exception as e:
        logger.warning(f"Lỗi không xác định khi tính toán/thiết lập kích thước slide từ PDF, sử dụng mặc định: {e}",
                       exc_info=True)
        prs.slide_width = Inches(default_width_in)
        prs.slide_height = Inches(default_height_in)
        return prs


def _convert_pdf_to_pptx_images(input_path, output_path):
    """Chuyển PDF sang PPTX bằng cách chuyển từng trang thành hình ảnh"""
    temp_dir = None
    try:
        # Tạo thư mục tạm duy nhất cho các ảnh của request này trong UPLOAD_FOLDER
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_", dir=UPLOAD_FOLDER)
        logger.info(f"Tạo thư mục tạm cho ảnh: {temp_dir}")

        # Chuyển đổi PDF sang ảnh (JPEG chất lượng cao)
        images = convert_from_path(
            input_path,
            dpi=300,  # Độ phân giải cao
            fmt='jpeg',  # Định dạng ảnh
            output_folder=temp_dir,
            thread_count=4,  # Tăng tốc độ nếu CPU cho phép
            jpegopt={'quality': 95, 'progressive': True},  # Chất lượng JPEG tốt
            poppler_path=None
            # Để trống nếu poppler nằm trong PATH, nếu không thì chỉ định đường dẫn tới thư mục bin của poppler
        )

        if not images:
            # convert_from_path trả về danh sách các đối tượng PIL Image
            # Nếu danh sách trống tức là có lỗi hoặc PDF trống
            raise ValueError("pdf2image không trả về ảnh nào. PDF có thể trống hoặc có lỗi đọc.")

        prs = Presentation()
        # Thiết lập kích thước slide dựa trên PDF gốc
        prs = setup_slide_size(prs, input_path)

        # Sử dụng layout trống (thường là index 6)
        try:
            blank_layout = prs.slide_layouts[6]
        except IndexError:
            logger.warning("Không tìm thấy slide layout 6 (blank), sử dụng layout 0.")
            blank_layout = prs.slide_layouts[0]

        # Lấy danh sách file ảnh đã tạo và sắp xếp đúng thứ tự
        # pdf2image tạo file với tên chứa số trang, sort tự nhiên thường là đủ
        image_files = sorted(
            [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))],
            key=lambda f: int(os.path.splitext(f)[0].split('-')[-1])
            # Sort based on page number in filename like '...-01.jpg'
        )

        if not image_files:
            # This case shouldn't happen if `images` list was not empty, but check anyway
            raise ValueError("Không tìm thấy file ảnh JPEG nào trong thư mục tạm sau khi convert_from_path.")

        logger.info(f"Tìm thấy {len(image_files)} ảnh trang để thêm vào PPTX.")

        # Thêm từng ảnh vào một slide mới
        for i, image_path in enumerate(image_files):
            try:
                slide = prs.slides.add_slide(blank_layout)

                # Lấy kích thước ảnh để tính toán tỷ lệ
                with Image.open(image_path) as img:
                    img_width_px, img_height_px = img.size

                if img_width_px <= 0 or img_height_px <= 0:
                    logger.warning(f"Kích thước ảnh không hợp lệ cho {os.path.basename(image_path)}. Bỏ qua.")
                    continue

                img_ratio = img_width_px / img_height_px

                slide_width_emu = prs.slide_width
                slide_height_emu = prs.slide_height
                slide_ratio = slide_width_emu / slide_height_emu

                # Tính toán kích thước và vị trí để ảnh vừa slide mà không bị méo
                if img_ratio > slide_ratio:  # Ảnh rộng hơn slide -> Chiều rộng ảnh bằng chiều rộng slide
                    pic_width = slide_width_emu
                    pic_height = int(pic_width / img_ratio)
                    pic_left = 0
                    pic_top = int((slide_height_emu - pic_height) / 2)
                else:  # Ảnh cao hơn hoặc bằng slide -> Chiều cao ảnh bằng chiều cao slide
                    pic_height = slide_height_emu
                    pic_width = int(pic_height * img_ratio)
                    pic_left = int((slide_width_emu - pic_width) / 2)
                    pic_top = 0

                # Đảm bảo kích thước không âm hoặc quá nhỏ
                if pic_width > 0 and pic_height > 0:
                    slide.shapes.add_picture(image_path, pic_left, pic_top, width=pic_width, height=pic_height)
                    logger.debug(f"Đã thêm ảnh {os.path.basename(image_path)} vào slide {i + 1}.")
                else:
                    logger.warning(
                        f"Kích thước tính toán cho ảnh {os.path.basename(image_path)} không hợp lệ ({pic_width}x{pic_height}). Bỏ qua.")

            except FileNotFoundError:
                logger.warning(f"File ảnh {image_path} không tìm thấy khi thêm vào slide. Bỏ qua.")
            except Exception as page_err:
                logger.warning(
                    f"Lỗi khi thêm ảnh {os.path.basename(image_path)} vào slide {i + 1}: {page_err}. Bỏ qua trang này.",
                    exc_info=True)  # Log traceback for page errors

        if not prs.slides:
            raise ValueError("Không có slide nào được thêm vào bản trình bày. Có thể tất cả ảnh đều bị lỗi.")

        prs.save(output_path)
        logger.info(f"Đã lưu PPTX thành công tại: {output_path}")
        return True

    except ValueError as ve:
        logger.error(f"Lỗi giá trị khi chuyển đổi PDF sang PPTX (ảnh): {ve}")
        return False
    except ImportError:
        logger.error("Lỗi: Thư viện pdf2image hoặc Pillow chưa được cài đặt đúng cách.")
        return False
    except Exception as e:
        # Catch potential errors from poppler/pdf2image itself
        logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang PPTX (phương pháp hình ảnh): {e}", exc_info=True)
        return False
    finally:
        # Luôn dọn dẹp thư mục tạm chứa ảnh, ngay cả khi có lỗi
        if temp_dir:  # Check if temp_dir was assigned
            safe_remove(temp_dir)  # Use safe_remove for directories


def convert_pdf_to_pptx_python(input_path, output_path):
    """Chuyển PDF sang PPTX chỉ sử dụng phương pháp hình ảnh (ổn định hơn)"""
    logger.info("Thử chuyển đổi PDF -> PPTX bằng phương pháp hình ảnh (Python/pdf2image)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)


def convert_jpg_to_pdf(input_path, output_path):
    """Chuyển đổi JPG/JPEG sang PDF"""
    try:
        with Image.open(input_path) as image:
            # Chuyển sang RGB nếu là ảnh có palette (P) hoặc RGBA để tránh lỗi Pillow
            if image.mode in ['P', 'RGBA']:
                logger.info(f"Ảnh '{os.path.basename(input_path)}' có mode {image.mode}, chuyển sang RGB.")
                image = image.convert('RGB')
            elif image.mode not in ['RGB', 'L', 'CMYK']:  # L là grayscale, CMYK cũng thường được hỗ trợ
                logger.warning(
                    f"Ảnh '{os.path.basename(input_path)}' có mode không điển hình: {image.mode}. Cố gắng chuyển sang RGB.")
                image = image.convert('RGB')

            # Lưu ảnh dưới dạng PDF
            image.save(output_path, "PDF", resolution=100.0, save_all=False)  # save_all=False cho single image PDF
        logger.info(f"Đã chuyển đổi {input_path} sang PDF thành công: {output_path}")
        return True
    except FileNotFoundError:
        logger.error(f"File ảnh không tìm thấy: {input_path}")
        return False
    except Exception as e:
        logger.error(f"Lỗi chuyển đổi JPG/JPEG sang PDF cho file {input_path}: {e}", exc_info=True)
        return False


# --- Flask Routes ---

@app.route('/health')
def health_check():
    """Endpoint kiểm tra tình trạng server"""
    status = {'status': 'OK', 'libreoffice_found': bool(SOFFICE_PATH)}
    http_status = 200
    if not SOFFICE_PATH:
        status['status'] = 'WARN'
        status['message'] = 'LibreOffice executable not found, conversions depending on it will fail.'
        # http_status = 503 # Service Unavailable might be too strong, maybe just 200 with warning?
    return jsonify(status), http_status


@app.route('/get_translations')
def get_translations():
    """Trả về các bản dịch ngôn ngữ được yêu cầu"""
    # (Giữ nguyên như trong code gốc của bạn)
    translations = {
        'en': {
            'lang-title': 'PDF Tools',
            'lang-subtitle': 'Simple, powerful PDF tools for everyone',
            'lang-error-title': 'Error!',
            'lang-convert-title': 'Convert PDF',
            'lang-convert-desc': 'Transform PDFs to other formats or vice versa',
            'lang-compress-title': 'Compress PDF',
            'lang-compress-desc': 'Reduce file size while maintaining quality',
            'lang-merge-title': 'Merge PDF',
            'lang-merge-desc': 'Combine multiple PDFs into one file',
            'lang-split-title': 'Split PDF',
            'lang-split-desc': 'Extract pages from your PDF',
            'lang-rotate-title': 'Rotate PDF',
            'lang-rotate-desc': 'Change page orientation',
            'lang-edit-title': 'Edit PDF',
            'lang-edit-desc': 'Modify text and images in your PDF',
            'lang-size-limit': 'Size limit: 100MB',
            'lang-select-conversion': 'Select conversion type',
            'lang-converting': 'Converting...',
            'lang-convert-btn': 'Convert Now',
            'lang-file-input-label': 'Select file',
            'file-no-selected': 'No file selected',
            'err-select-file': 'Please select a file to convert.',
            'err-file-too-large': 'File is too large. Limit is 100MB.',
            'err-select-conversion': 'Please select a conversion type.',
            'err-format-docx': 'File format not compatible with PDF ↔ DOCX conversion.',
            'err-format-ppt': 'File format not compatible with PDF ↔ PPT conversion.',
            'err-format-jpg': 'File format not compatible with PDF ↔ JPG conversion.',
            'err-conversion': 'An error occurred during conversion.',
            'err-fetch-translations': 'Could not load language data.',
            'lang-select-btn-text': 'Browse',
            'lang-select-conversion-label': 'Conversion Type'
        },
        'vi': {
            'lang-title': 'Công Cụ PDF',
            'lang-subtitle': 'Công cụ PDF đơn giản, mạnh mẽ cho mọi người',
            'lang-error-title': 'Lỗi!',
            'lang-convert-title': 'Chuyển đổi PDF',
            'lang-convert-desc': 'Chuyển đổi PDF sang các định dạng khác hoặc ngược lại',
            'lang-compress-title': 'Nén PDF',
            'lang-compress-desc': 'Giảm kích thước tệp trong khi duy trì chất lượng',
            'lang-merge-title': 'Gộp PDF',
            'lang-merge-desc': 'Kết hợp nhiều tệp PDF thành một tệp',
            'lang-split-title': 'Tách PDF',
            'lang-split-desc': 'Trích xuất các trang từ tệp PDF của bạn',
            'lang-rotate-title': 'Xoay PDF',
            'lang-rotate-desc': 'Thay đổi hướng trang',
            'lang-edit-title': 'Chỉnh sửa PDF',
            'lang-edit-desc': 'Sửa đổi văn bản và hình ảnh trong tệp PDF của bạn',
            'lang-size-limit': 'Giới hạn kích thước: 100MB',
            'lang-select-conversion': 'Chọn kiểu chuyển đổi',
            'lang-converting': 'Đang chuyển đổi...',
            'lang-convert-btn': 'Chuyển đổi ngay',
            'lang-file-input-label': 'Chọn tệp',
            'file-no-selected': 'Không có tệp nào được chọn',
            'err-select-file': 'Vui lòng chọn một tệp để chuyển đổi.',
            'err-file-too-large': 'Tệp quá lớn. Giới hạn là 100MB.',
            'err-select-conversion': 'Vui lòng chọn kiểu chuyển đổi.',
            'err-format-docx': 'Định dạng tệp không phù hợp với kiểu chuyển đổi PDF ↔ DOCX.',
            'err-format-ppt': 'Định dạng tệp không phù hợp với kiểu chuyển đổi PDF ↔ PPT.',
            'err-format-jpg': 'Định dạng tệp không phù hợp với kiểu chuyển đổi PDF ↔ JPG.',
            'err-conversion': 'Đã xảy ra lỗi trong quá trình chuyển đổi.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.',
            'lang-select-btn-text': 'Duyệt...',
            'lang-select-conversion-label': 'Kiểu chuyển đổi'
        }
    }
    lang = request.args.get('lang', 'en')
    return jsonify(translations.get(lang, translations['en']))


@app.route('/')
def index():
    """Trang chủ hiển thị form upload"""
    # Đảm bảo thư mục upload tồn tại khi truy cập trang chủ
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    translations_url = url_for('get_translations')
    return render_template('index.html', translations_url=translations_url)


@app.route('/convert', methods=['POST'])
def convert_file():
    """Xử lý chuyển đổi file"""
    input_path = None
    output_path = None
    # Biến lưu tên file tạm do LibreOffice tạo ra (tên thường giống input nhưng khác extension)
    temp_libreoffice_output_path = None

    # --- Initial Checks ---
    # Check if LibreOffice is needed and available early
    conversion_type = request.form.get('conversion_type')
    libreoffice_needed = conversion_type in ['pdf_to_docx', 'docx_to_pdf', 'pdf_to_ppt', 'ppt_to_pdf']

    if libreoffice_needed and SOFFICE_PATH is None:
        logger.error(f"LibreOffice is required for conversion '{conversion_type}' but was not found.")
        # Return 503 Service Unavailable as the core tool is missing
        return "Lỗi máy chủ: Công cụ chuyển đổi cần thiết (LibreOffice) không khả dụng. Vui lòng liên hệ quản trị viên.", 503

    try:
        # 1. Validate request - File present?
        if 'file' not in request.files:
            logger.warning("Request thiếu 'file' part.")
            return jsonify({'error': 'err-select-file'}), 400  # Use key for translation
        file = request.files['file']
        if not file or file.filename == '':
            logger.warning("Request có 'file' part nhưng không có tên file.")
            return jsonify({'error': 'file-no-selected'}), 400  # Use key for translation

        # 2. Secure filename and check extension
        filename = secure_filename(file.filename)
        if not allowed_file(filename):
            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'không có'
            allowed_str = ", ".join(ALLOWED_EXTENSIONS)
            logger.warning(f"Loại file không hợp lệ: '{ext}'. Input filename: '{filename}'")
            # Determine specific error key based on conversion type
            error_key = 'err-conversion'  # Generic fallback
            if conversion_type == 'pdf_to_docx' or conversion_type == 'docx_to_pdf':
                error_key = 'err-format-docx'
            elif conversion_type == 'pdf_to_ppt' or conversion_type == 'ppt_to_pdf':
                error_key = 'err-format-ppt'
            elif conversion_type == 'jpg_to_pdf':
                error_key = 'err-format-jpg'
            return jsonify({'error': error_key, 'details': f"Allowed: {allowed_str}"}), 400

        # 3. Get conversion type (already fetched)
        if not conversion_type:
            logger.warning("Request thiếu 'conversion_type'.")
            return jsonify({'error': 'err-select-conversion'}), 400

        logger.info(f"Yêu cầu chuyển đổi nhận được: file='{filename}', type='{conversion_type}'")

        # 4. Prepare paths
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        input_filename_base, input_ext = os.path.splitext(filename)
        input_ext = input_ext.lower()  # Ensure extension is lowercase
        # Create a more unique prefix
        unique_prefix = f"{int(time.time())}_{os.urandom(4).hex()}"
        input_path = os.path.join(UPLOAD_FOLDER, f"input_{unique_prefix}{input_ext}")

        # 5. Save uploaded file
        try:
            file.save(input_path)
            # Check file size after saving
            if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
                raise OSError("File lưu vào bị trống hoặc không tồn tại.")
            logger.info(f"File đã lưu vào: {input_path} (Size: {os.path.getsize(input_path)} bytes)")
        except Exception as save_err:
            logger.error(f"Lỗi khi lưu file upload vào {input_path}: {save_err}", exc_info=True)
            # Don't leave partial/empty files around
            safe_remove(input_path)
            return jsonify({'error': 'err-conversion', 'details': 'Failed to save uploaded file.'}), 500

        # 6. Determine output extension and final output path
        out_ext = ''
        if conversion_type == 'pdf_to_docx':
            out_ext = '.docx'
        elif conversion_type == 'docx_to_pdf':
            out_ext = '.pdf'
        elif conversion_type == 'pdf_to_ppt':
            out_ext = '.pptx'  # Always create modern pptx
        elif conversion_type == 'ppt_to_pdf':
            out_ext = '.pdf'
        elif conversion_type == 'jpg_to_pdf':
            out_ext = '.pdf'
        else:
            # This case should ideally not be reached due to earlier checks, but as a safeguard:
            safe_remove(input_path)  # Clean up saved input file
            logger.error(f"Loại chuyển đổi không được hỗ trợ trong logic chính: {conversion_type}")
            return jsonify({'error': 'err-select-conversion', 'details': 'Unsupported conversion type.'}), 400

        # Construct user-friendly output filename (original name + suffix + new ext)
        output_filename_base = f"{input_filename_base}_converted"
        output_filename = f"{output_filename_base}{out_ext}"
        # Path for storing the file on server uses the unique prefix
        output_path = os.path.join(UPLOAD_FOLDER, f"{output_filename_base}_{unique_prefix}{out_ext}")
        logger.info(f"File output dự kiến (server path): {output_path}")
        logger.info(f"File output dự kiến (download name): {output_filename}")

        # --- 7. Perform Conversion ---
        conversion_success = False
        # More specific error message to potentially show user (keep it simple)
        user_friendly_error = "Đã xảy ra lỗi không xác định."

        # --- PDF to DOCX (Using LibreOffice) ---
        if conversion_type == 'pdf_to_docx':
            # SOFFICE_PATH checked at the start
            try:
                # LibreOffice usually creates output with same basename as input
                expected_lo_output_name = os.path.splitext(os.path.basename(input_path))[0] + '.docx'
                temp_libreoffice_output_path = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)

                # Clean up potential leftovers before running
                safe_remove(temp_libreoffice_output_path)
                safe_remove(output_path)  # Also remove final target path if it exists

                logger.info(
                    f"Chạy LibreOffice (PDF->DOCX): {SOFFICE_PATH} --headless --infilter=\"writer_pdf_import\" --convert-to docx:\"MS Word 2007 XML\" --outdir {UPLOAD_FOLDER} {input_path}")
                cmd = [
                    SOFFICE_PATH,
                    '--headless',
                    '--invisible',  # Use invisible instead of just headless sometimes helps
                    '--infilter="writer_pdf_import"',  # Explicit import filter
                    '--convert-to', 'docx:"MS Word 2007 XML"',  # Explicit format
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ]
                # Use a longer timeout for potentially complex PDFs
                result = subprocess.run(cmd, check=True, timeout=240, capture_output=True, text=True, encoding='utf-8',
                                        errors='replace')

                # Log output even if successful
                logger.info(f"LibreOffice (PDF->DOCX) stdout: {result.stdout}")
                if result.stderr:
                    # Stderr might contain warnings, not necessarily errors
                    logger.warning(f"LibreOffice (PDF->DOCX) stderr: {result.stderr}")

                # Check if the expected temporary file was created
                if os.path.exists(temp_libreoffice_output_path):
                    # Rename the temp file to our final unique output path
                    try:
                        os.rename(temp_libreoffice_output_path, output_path)
                        # Verify the final file exists and is not empty
                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                            conversion_success = True
                            logger.info(f"Chuyển đổi PDF -> DOCX bằng LibreOffice thành công. Output: {output_path}")
                        else:
                            user_friendly_error = "Chuyển đổi hoàn tất nhưng file kết quả bị lỗi hoặc trống."
                            logger.error(user_friendly_error + f" File đích sau đổi tên: {output_path}")
                            safe_remove(temp_libreoffice_output_path)  # Clean up original LO output if rename failed
                    except OSError as rename_err:
                        user_friendly_error = "Lỗi hệ thống khi lưu file kết quả."
                        logger.error(f"Lỗi OS khi đổi tên file DOCX output từ LibreOffice: {rename_err}", exc_info=True)
                        safe_remove(temp_libreoffice_output_path)  # Clean up original LO output

                else:
                    user_friendly_error = "Công cụ chuyển đổi không tạo ra file kết quả như mong đợi."
                    logger.error(user_friendly_error)
                    logger.error(
                        f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    # Log directory listing for debugging if file not found
                    try:
                        logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")
                    except OSError as list_err:
                        logger.error(f"Không thể liệt kê thư mục upload: {list_err}")


            except subprocess.CalledProcessError as e:
                user_friendly_error = "Lỗi trong quá trình xử lý file PDF."
                logger.error(
                    f"Lỗi LibreOffice (PDF->DOCX - CalledProcessError): Return code {e.returncode}. Stderr: {e.stderr}",
                    exc_info=False)  # Log stderr without full traceback for this specific error
            except subprocess.TimeoutExpired:
                user_friendly_error = "Quá trình chuyển đổi mất quá nhiều thời gian."
                logger.error(f"Lỗi LibreOffice (PDF->DOCX): Quá thời gian chuyển đổi (240s). File: {input_path}")
            except Exception as e:
                user_friendly_error = "Lỗi không xác định trong quá trình chuyển đổi PDF."
                logger.error(f"Lỗi không xác định khi chạy LibreOffice (PDF->DOCX): {e}", exc_info=True)


        # --- DOCX to PDF (Using LibreOffice) ---
        elif conversion_type == 'docx_to_pdf':
            # SOFFICE_PATH checked at the start
            try:
                expected_lo_output_name = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
                temp_libreoffice_output_path = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                safe_remove(temp_libreoffice_output_path)
                safe_remove(output_path)

                logger.info(
                    f"Chạy LibreOffice (DOCX->PDF): {SOFFICE_PATH} --headless --invisible --convert-to pdf:writer_pdf_Export --outdir {UPLOAD_FOLDER} {input_path}")
                cmd = [
                    SOFFICE_PATH,
                    '--headless',
                    '--invisible',
                    '--convert-to', 'pdf:writer_pdf_Export',  # Explicit PDF export filter
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ]
                result = subprocess.run(cmd, check=True, timeout=180, capture_output=True, text=True, encoding='utf-8',
                                        errors='replace')
                logger.info(f"LibreOffice (DOCX->PDF) stdout: {result.stdout}")
                if result.stderr:
                    logger.warning(f"LibreOffice (DOCX->PDF) stderr: {result.stderr}")

                if os.path.exists(temp_libreoffice_output_path):
                    try:
                        os.rename(temp_libreoffice_output_path, output_path)
                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                            conversion_success = True
                            logger.info(f"Chuyển đổi DOCX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                        else:
                            user_friendly_error = "Chuyển đổi hoàn tất nhưng file kết quả bị lỗi hoặc trống."
                            logger.error(user_friendly_error + f" File đích sau đổi tên: {output_path}")
                            safe_remove(temp_libreoffice_output_path)
                    except OSError as rename_err:
                        user_friendly_error = "Lỗi hệ thống khi lưu file kết quả."
                        logger.error(f"Lỗi OS khi đổi tên file PDF output (từ DOCX) từ LibreOffice: {rename_err}",
                                     exc_info=True)
                        safe_remove(temp_libreoffice_output_path)
                else:
                    user_friendly_error = "Công cụ chuyển đổi không tạo ra file kết quả như mong đợi."
                    logger.error(user_friendly_error)
                    logger.error(
                        f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    try:
                        logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")
                    except OSError as list_err:
                        logger.error(f"Không thể liệt kê thư mục upload: {list_err}")

            except subprocess.CalledProcessError as e:
                user_friendly_error = "Lỗi trong quá trình xử lý file DOCX."
                logger.error(
                    f"Lỗi LibreOffice (DOCX->PDF - CalledProcessError): Return code {e.returncode}. Stderr: {e.stderr}",
                    exc_info=False)
            except subprocess.TimeoutExpired:
                user_friendly_error = "Quá trình chuyển đổi mất quá nhiều thời gian."
                logger.error(f"Lỗi LibreOffice (DOCX->PDF): Quá thời gian chuyển đổi (180s). File: {input_path}")
            except Exception as e:
                user_friendly_error = "Lỗi không xác định trong quá trình chuyển đổi DOCX."
                logger.error(f"Lỗi không xác định khi chạy LibreOffice (DOCX->PDF): {e}", exc_info=True)

        # --- PDF to PPT (Python Image Method First, Fallback LibreOffice) ---
        elif conversion_type == 'pdf_to_ppt':
            logger.info("Bắt đầu chuyển đổi PDF -> PPTX...")
            # Try Python image-based method first
            if convert_pdf_to_pptx_python(input_path, output_path):
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    conversion_success = True
                    logger.info("Chuyển đổi PDF -> PPTX bằng Python (ảnh) thành công.")
                else:
                    logger.warning(
                        "Phương thức Python (ảnh) chạy xong nhưng file PPTX không hợp lệ hoặc trống. Thử fallback...")
                    safe_remove(output_path)  # Remove potentially bad file
                    # Proceed to LibreOffice fallback
            else:
                # Error should have been logged inside convert_pdf_to_pptx_python
                logger.warning("Chuyển PDF->PPTX bằng Python (ảnh) thất bại, thử dùng LibreOffice...")

            # Fallback to LibreOffice if Python method failed or produced invalid file
            if not conversion_success:
                # SOFFICE_PATH checked at the start
                try:
                    expected_lo_output_name = os.path.splitext(os.path.basename(input_path))[0] + '.pptx'
                    temp_libreoffice_output_path = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                    safe_remove(temp_libreoffice_output_path)
                    safe_remove(output_path)  # Ensure clean slate for LO output

                    logger.info(
                        f"Chạy LibreOffice (PDF->PPTX fallback): {SOFFICE_PATH} --headless --invisible --convert-to pptx:impress_pdf_import --outdir {UPLOAD_FOLDER} {input_path}")
                    cmd = [
                        SOFFICE_PATH,
                        '--headless',
                        '--invisible',
                        # Using specific filter might help Impress handle PDF better
                        '--convert-to', 'pptx:impress_pdf_import',
                        '--outdir', UPLOAD_FOLDER,
                        input_path
                    ]
                    result = subprocess.run(cmd, check=True, timeout=300, capture_output=True, text=True,
                                            encoding='utf-8', errors='replace')  # Longer timeout for complex PDFs
                    logger.info(f"LibreOffice (PDF->PPTX fallback) stdout: {result.stdout}")
                    if result.stderr:
                        logger.warning(f"LibreOffice (PDF->PPTX fallback) stderr: {result.stderr}")

                    if os.path.exists(temp_libreoffice_output_path):
                        try:
                            os.rename(temp_libreoffice_output_path, output_path)
                            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                                conversion_success = True
                                logger.info(
                                    f"Chuyển đổi PDF -> PPTX bằng LibreOffice thành công (fallback). Output: {output_path}")
                            else:
                                user_friendly_error = "Chuyển đổi (fallback) hoàn tất nhưng file kết quả bị lỗi hoặc trống."
                                logger.error(user_friendly_error + f" File đích sau đổi tên: {output_path}")
                                safe_remove(temp_libreoffice_output_path)
                        except OSError as rename_err:
                            user_friendly_error = "Lỗi hệ thống khi lưu file kết quả (fallback)."
                            logger.error(f"Lỗi OS khi đổi tên file PPTX output từ LibreOffice (fallback): {rename_err}",
                                         exc_info=True)
                            safe_remove(temp_libreoffice_output_path)
                    else:
                        user_friendly_error = "Công cụ chuyển đổi (fallback) không tạo ra file kết quả như mong đợi."
                        logger.error(user_friendly_error)
                        logger.error(
                            f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                        try:
                            logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")
                        except OSError as list_err:
                            logger.error(f"Không thể liệt kê thư mục upload: {list_err}")

                except subprocess.CalledProcessError as e:
                    user_friendly_error = "Lỗi trong quá trình xử lý file PDF (fallback)."
                    logger.error(
                        f"Lỗi LibreOffice (fallback PDF->PPTX): Return code {e.returncode}. Stderr: {e.stderr}",
                        exc_info=False)
                except subprocess.TimeoutExpired:
                    user_friendly_error = "Quá trình chuyển đổi (fallback) mất quá nhiều thời gian."
                    logger.error(
                        f"Lỗi LibreOffice (fallback PDF->PPTX): Quá thời gian chuyển đổi (300s). File: {input_path}")
                except Exception as e:
                    user_friendly_error = "Lỗi không xác định trong quá trình chuyển đổi PDF (fallback)."
                    logger.error(f"Lỗi không xác định khi chạy LibreOffice (fallback PDF->PPTX): {e}", exc_info=True)

            # Final error message if both methods failed
            if not conversion_success and user_friendly_error == "Đã xảy ra lỗi không xác định.":
                user_friendly_error = "Không thể chuyển đổi PDF sang PPTX bằng cả hai phương pháp."


        # --- PPT/PPTX to PDF (Using LibreOffice) ---
        elif conversion_type == 'ppt_to_pdf':
            # SOFFICE_PATH checked at the start
            try:
                expected_lo_output_name = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
                temp_libreoffice_output_path = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                safe_remove(temp_libreoffice_output_path)
                safe_remove(output_path)

                logger.info(
                    f"Chạy LibreOffice (PPT/PPTX->PDF): {SOFFICE_PATH} --headless --invisible --convert-to pdf:impress_pdf_Export --outdir {UPLOAD_FOLDER} {input_path}")
                cmd = [
                    SOFFICE_PATH,
                    '--headless',
                    '--invisible',
                    '--convert-to', 'pdf:impress_pdf_Export',  # Explicit PDF export filter for Impress
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ]
                result = subprocess.run(cmd, check=True, timeout=240, capture_output=True, text=True, encoding='utf-8',
                                        errors='replace')  # Longer timeout for large PPTs
                logger.info(f"LibreOffice (PPT/PPTX->PDF) stdout: {result.stdout}")
                if result.stderr:
                    logger.warning(f"LibreOffice (PPT/PPTX->PDF) stderr: {result.stderr}")

                if os.path.exists(temp_libreoffice_output_path):
                    try:
                        os.rename(temp_libreoffice_output_path, output_path)
                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                            conversion_success = True
                            logger.info(
                                f"Chuyển đổi PPT/PPTX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                        else:
                            user_friendly_error = "Chuyển đổi hoàn tất nhưng file kết quả bị lỗi hoặc trống."
                            logger.error(user_friendly_error + f" File đích sau đổi tên: {output_path}")
                            safe_remove(temp_libreoffice_output_path)
                    except OSError as rename_err:
                        user_friendly_error = "Lỗi hệ thống khi lưu file kết quả."
                        logger.error(f"Lỗi OS khi đổi tên file PDF output (từ PPT) từ LibreOffice: {rename_err}",
                                     exc_info=True)
                        safe_remove(temp_libreoffice_output_path)
                else:
                    user_friendly_error = "Công cụ chuyển đổi không tạo ra file kết quả như mong đợi."
                    logger.error(user_friendly_error)
                    logger.error(
                        f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    try:
                        logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")
                    except OSError as list_err:
                        logger.error(f"Không thể liệt kê thư mục upload: {list_err}")

            except subprocess.CalledProcessError as e:
                user_friendly_error = "Lỗi trong quá trình xử lý file PowerPoint."
                logger.error(f"Lỗi LibreOffice (PPT/PPTX->PDF): Return code {e.returncode}. Stderr: {e.stderr}",
                             exc_info=False)
            except subprocess.TimeoutExpired:
                user_friendly_error = "Quá trình chuyển đổi mất quá nhiều thời gian."
                logger.error(f"Lỗi LibreOffice (PPT/PPTX->PDF): Quá thời gian chuyển đổi (240s). File: {input_path}")
            except Exception as e:
                user_friendly_error = "Lỗi không xác định trong quá trình chuyển đổi PowerPoint."
                logger.error(f"Lỗi không xác định khi chạy LibreOffice (PPT/PPTX->PDF): {e}", exc_info=True)

        # --- JPG/JPEG to PDF (Using Pillow) ---
        elif conversion_type == 'jpg_to_pdf':
            if convert_jpg_to_pdf(input_path, output_path):
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    conversion_success = True
                    logger.info("Chuyển đổi JPG -> PDF thành công.")
                else:
                    user_friendly_error = "Chuyển đổi ảnh sang PDF hoàn tất nhưng file kết quả bị lỗi hoặc trống."
                    logger.warning(user_friendly_error + f" File output: {output_path}")
                    safe_remove(output_path)  # Remove bad output
            else:
                # Error should have been logged inside convert_jpg_to_pdf
                user_friendly_error = "Không thể chuyển đổi file ảnh sang PDF."
                logger.error(f"Chuyển đổi JPG -> PDF thất bại. Input: {input_path}")

        # --- END OF CONVERSION LOGIC ---

        # 8. Handle result and send file / return error
        # Clean up the input file *after* conversion attempt (success or fail)
        # Input path might be None if saving failed earlier, so check existence
        if input_path and os.path.exists(input_path):
            if not safe_remove(input_path):
                logger.warning(f"Không thể xóa file input sau xử lý: {input_path}")
            else:
                logger.debug(f"Đã xóa file input sau xử lý: {input_path}")

        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            try:
                output_size = os.path.getsize(output_path)
                logger.info(
                    f"Chuyển đổi thành công. Chuẩn bị gửi file output: {output_path}, Tên tải về: {output_filename}, Kích thước: {output_size} bytes")

                # Use a Response object to delete file after sending is complete
                response = send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename  # Use the user-friendly filename
                )

                # Register callback to clean up the specific output file for this request
                # Use a local variable scope for the path to ensure the correct path is captured
                _output_path_to_delete = output_path

                @response.call_on_close
                def cleanup_specific_output_file():
                    logger.info(f"Client đã nhận file (hoặc đóng kết nối), xóa file output: {_output_path_to_delete}")
                    safe_remove(_output_path_to_delete)

                return response

            except Exception as send_err:
                logger.error(f"Lỗi khi gửi file {output_path}: {send_err}", exc_info=True)
                # Attempt to clean up the output file if sending failed
                safe_remove(output_path)
                # Clean up potential LO temp file if it wasn't renamed/deleted
                if temp_libreoffice_output_path and os.path.exists(temp_libreoffice_output_path):
                    safe_remove(temp_libreoffice_output_path)
                # Provide a user-friendly error message
                return jsonify({'error': 'err-conversion', 'details': 'Lỗi khi chuẩn bị file để tải về.'}), 500
        else:
            # Conversion failed or produced an empty/invalid file
            logger.error(
                f"Chuyển đổi thất bại hoặc file output không hợp lệ. Lý do: {user_friendly_error}. Input: {filename}")
            # Ensure potential output files (final and temp) are cleaned up
            if output_path and os.path.exists(output_path):
                safe_remove(output_path)
            if temp_libreoffice_output_path and os.path.exists(temp_libreoffice_output_path):
                safe_remove(temp_libreoffice_output_path)

            # Return the user-friendly error message
            return jsonify({'error': 'err-conversion', 'details': user_friendly_error}), 500

    # --- Catch-all Exception Handler ---
    except Exception as e:
        # Log the detailed error for server-side debugging
        error_id = os.urandom(6).hex()  # Generate a unique ID for this error instance
        logger.error(f"Lỗi không mong muốn trong route /convert (ID: {error_id}): {e}", exc_info=True)

        # Ensure cleanup of any potential temporary files created before the error
        if input_path and os.path.exists(input_path):
            safe_remove(input_path)
        if output_path and os.path.exists(output_path):
            safe_remove(output_path)
        if temp_libreoffice_output_path and os.path.exists(temp_libreoffice_output_path):
            safe_remove(temp_libreoffice_output_path)

        # Return a generic server error message to the client, including the ID
        return jsonify({'error': 'err-conversion',
                        'details': f'Đã xảy ra lỗi máy chủ không mong muốn. Vui lòng thử lại sau. (Mã lỗi: {error_id})'}), 500


# --- After Request Hook (Optional but good for caching headers) ---
@app.after_request
def after_request_func(response):
    # Add headers to prevent caching of downloadable files
    if 'Content-Disposition' in response.headers and response.headers['Content-Disposition'].startswith('attachment'):
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'  # For HTTP/1.0 backwards compatibility
        response.headers['Expires'] = '0'  # Proxies
    # Add security headers
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'  # Or DENY if not using frames
    # response.headers['Content-Security-Policy'] = "default-src 'self'" # Adjust as needed
    return response


# --- Teardown App Context (Scheduled Cleanup) ---
@app.teardown_appcontext
def cleanup_old_files(exception=None):
    """Dọn dẹp các file và thư mục cũ trong thư mục upload"""
    if not os.path.exists(UPLOAD_FOLDER):
        return

    # Thời gian tối đa giữ lại file (ví dụ: 1 giờ) - Điều chỉnh nếu cần
    max_age_seconds = 1 * 60 * 60

    logger.info(f"Chạy dọn dẹp teardown_appcontext (xóa mục cũ hơn {max_age_seconds}s)...")
    try:
        now = time.time()
        cutoff_time = now - max_age_seconds
        cleaned_count = 0
        checked_count = 0

        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            checked_count += 1
            try:
                # Lấy thời gian sửa đổi cuối cùng
                mod_time = os.path.getmtime(path)

                if mod_time < cutoff_time:
                    if os.path.isfile(path):
                        logger.info(f"Teardown: Xóa file cũ (tuổi: {now - mod_time:.0f}s > {max_age_seconds}s): {path}")
                        if safe_remove(path):
                            cleaned_count += 1
                    elif os.path.isdir(path):
                        # Chỉ xóa các thư mục tạm do chúng ta tạo ra (ví dụ: pdfimg_)
                        if filename.startswith("pdfimg_"):
                            logger.info(
                                f"Teardown: Xóa thư mục tạm cũ (tuổi: {now - mod_time:.0f}s > {max_age_seconds}s): {path}")
                            if safe_remove(path):  # safe_remove handles rmtree for dirs
                                cleaned_count += 1
                        else:
                            logger.debug(f"Teardown: Bỏ qua thư mục không xác định {path}")
                # else: # Uncomment for debugging what's NOT being deleted
                #     logger.debug(f"Teardown: Giữ lại mục {path} (tuổi: {now - mod_time:.0f}s <= {max_age_seconds}s)")

            except FileNotFoundError:
                # Item might have been deleted by another process/request just before this check
                continue
            except Exception as e:
                logger.warning(f"Lỗi khi kiểm tra/dọn dẹp mục {path} trong teardown: {e}")

        logger.info(f"Teardown dọn dẹp hoàn tất. Đã kiểm tra {checked_count} mục, xóa {cleaned_count} mục cũ.")

    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng trong quá trình dọn dẹp teardown_appcontext: {e}", exc_info=True)


# --- Main Execution ---
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    # Read DEBUG from environment variable, default to False
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() in ('true', '1', 't')
    host = '0.0.0.0' # Listen on all interfaces

    logger.info(f"--- Khởi động server Flask ---")
    logger.info(f" * Host: {host}")
    logger.info(f" * Cổng: {port}")
    logger.info(f" * Chế độ Debug: {debug_mode}")
    logger.info(f" * Thư mục Upload: {UPLOAD_FOLDER}")
    # Log LibreOffice path status regardless of how the server is run
    logger.info(f" * Đường dẫn LibreOffice: {SOFFICE_PATH if SOFFICE_PATH else 'KHÔNG TÌM THẤY'}")

    # Create upload folder if it doesn't exist at startup
    try:
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        logger.info(f"Đảm bảo thư mục upload tồn tại: {UPLOAD_FOLDER}")
    except OSError as e:
        logger.critical(f"KHÔNG THỂ TẠO THƯ MỤC UPLOAD '{UPLOAD_FOLDER}': {e}. Server không thể hoạt động đúng.")
        sys.exit(1) # Exit if upload folder cannot be created

    # Always use Flask's built-in server
    logger.info(f"Chạy server bằng Flask development server (Debug={debug_mode})...")
    app.run(host=host, port=port, debug=debug_mode)

# --- END OF FILE app.py ---