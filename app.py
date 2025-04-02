# --- START OF FILE app.py ---
# No changes needed in app.py based on the request.
# Keep the existing app.py code.

from flask import Flask, request, send_file, render_template, Response, jsonify, url_for
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


app = Flask(__name__, template_folder='templates', static_folder='static') # Thêm static_folder

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

# Endpoint trả về bản dịch - CẬP NHẬT VỚI ĐỦ KEYS
@app.route('/get_translations')
def get_translations():
    """Trả về các bản dịch ngôn ngữ được yêu cầu"""
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
            'err-format-jpg': 'File format not compatible with PDF ↔ JPG conversion.', # Giữ lại key phòng khi cần sau
            'err-conversion': 'An error occurred during conversion.',
            'err-fetch-translations': 'Could not load language data.', # Thêm key lỗi mới
            'lang-select-btn-text': 'Browse', # Thêm key cho nút Browse
            'lang-select-conversion-label': 'Conversion Type' # Thêm key cho label dropdown
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
            'err-format-jpg': 'Định dạng tệp không phù hợp với kiểu chuyển đổi PDF ↔ JPG.', # Giữ lại key phòng khi cần sau
            'err-conversion': 'Đã xảy ra lỗi trong quá trình chuyển đổi.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.', # Thêm key lỗi mới
            'lang-select-btn-text': 'Duyệt...', # Thêm key cho nút Browse (VI)
            'lang-select-conversion-label': 'Kiểu chuyển đổi' # Thêm key cho label dropdown (VI)
        }
    }
    lang = request.args.get('lang', 'en') # Lấy ngôn ngữ từ query param 'lang'
    return jsonify(translations.get(lang, translations['en'])) # Trả về JSON của ngôn ngữ đó


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
            # Kiểm tra xem path có phải là file thực thi không
            resolved_path = shutil.which(path)
            if resolved_path and os.path.isfile(resolved_path):
                 # Sử dụng resolved_path để chạy lệnh version
                result = subprocess.run([resolved_path, '--version'],
                                        capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0:
                    logger.info(f"Tìm thấy LibreOffice tại: {resolved_path}")
                    return resolved_path
            elif os.path.isfile(path): # Fallback nếu shutil.which không tìm thấy nhưng path là file
                 result = subprocess.run([path, '--version'],
                                        capture_output=True, text=True, check=False, timeout=5)
                 if result.returncode == 0:
                     logger.info(f"Tìm thấy LibreOffice tại: {path}")
                     return path

        except FileNotFoundError:
             logger.debug(f"Không tìm thấy LibreOffice tại {path}")
        except subprocess.TimeoutExpired:
             logger.warning(f"Kiểm tra LibreOffice tại {path} bị timeout.")
        except Exception as e:
            logger.warning(f"Lỗi khi kiểm tra LibreOffice tại {path}: {e}")

    logger.warning("Không tìm thấy LibreOffice thực thi qua các đường dẫn phổ biến hoặc PATH. Sử dụng 'soffice' mặc định.")
    return 'soffice' # Trả về 'soffice' để hệ thống tự tìm trong PATH


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
                logger.info(f"Đã xóa file tạm: {file_path}")
                return True
            else:
                return True # File không tồn tại, coi như thành công
        except PermissionError:
             logger.warning(f"Không có quyền xóa {file_path} (lần thử {i + 1}). Đang đợi...")
             time.sleep(delay*2) # Đợi lâu hơn nếu là lỗi quyền
        except Exception as e:
            logger.warning(f"Không thể xóa {file_path} (lần thử {i + 1}): {e}")
            time.sleep(delay)
    logger.error(f"Xóa file {file_path} thất bại sau {retries} lần thử.")
    return False


def get_pdf_page_size(pdf_path):
    """Lấy kích thước trang PDF (đơn vị points)"""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if not reader.pages:
                logger.warning(f"PDF không có trang nào: {pdf_path}")
                return None, None
            page = reader.pages[0]
            # Đảm bảo mediabox trả về đúng định dạng list/tuple
            mediabox = page.mediabox
            if mediabox:
                width = float(mediabox.width)
                height = float(mediabox.height)
                return width, height
            else:
                 logger.warning(f"Không tìm thấy mediabox cho trang đầu tiên trong {pdf_path}")
                 return None, None
    except Exception as e:
         logger.error(f"Lỗi khi đọc kích thước PDF {pdf_path}: {e}")
         return None, None


def setup_slide_size(prs, pdf_path):
    """Thiết lập kích thước slide dựa trên PDF"""
    pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)

    if pdf_width_pt is None or pdf_height_pt is None:
        logger.warning("Không thể đọc kích thước PDF, sử dụng kích thước mặc định (10x7.5 inches)")
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        return prs

    try:
        # Chuyển đổi từ points sang inches (1 inch = 72 points)
        pdf_width_in = pdf_width_pt / 72
        pdf_height_in = pdf_height_pt / 72

        # Giới hạn kích thước tối đa của PowerPoint (56 inches theo tài liệu mới, nhưng giữ an toàn)
        max_slide_dim = 50.0 # Giảm giới hạn một chút cho an toàn

        # Điều chỉnh kích thước nếu vượt quá giới hạn, giữ tỷ lệ
        if pdf_width_in > max_slide_dim or pdf_height_in > max_slide_dim:
            ratio = pdf_width_in / pdf_height_in
            if pdf_width_in >= pdf_height_in: # Rộng hơn hoặc vuông
                prs.slide_width = Inches(max_slide_dim)
                prs.slide_height = Inches(max_slide_dim / ratio)
            else: # Cao hơn
                prs.slide_height = Inches(max_slide_dim)
                prs.slide_width = Inches(max_slide_dim * ratio)
            logger.info(f"Kích thước gốc ({pdf_width_in:.2f}x{pdf_height_in:.2f} in) vượt giới hạn, điều chỉnh thành: {prs.slide_width.inches:.2f}x{prs.slide_height.inches:.2f} in")
        else:
            prs.slide_width = Inches(pdf_width_in)
            prs.slide_height = Inches(pdf_height_in)
            logger.info(f"Thiết lập kích thước slide theo PDF: {pdf_width_in:.2f} x {pdf_height_in:.2f} inches")

        return prs
    except Exception as e:
        logger.warning(f"Lỗi khi thiết lập kích thước slide từ PDF, sử dụng mặc định: {e}")
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        return prs

# Bỏ hàm này vì không hiệu quả và không được dùng làm fallback nữa
# def pdf_to_pptx_auto_fit(pdf_path, pptx_path): ...

def _convert_pdf_to_pptx_images(input_path, output_path):
    """Chuyển PDF sang PPTX bằng cách chuyển từng trang thành hình ảnh"""
    temp_dir = None # Khởi tạo để dùng trong finally
    try:
        # Tạo thư mục tạm để chứa ảnh (an toàn hơn)
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_")
        logger.info(f"Tạo thư mục tạm cho ảnh: {temp_dir}")

        # Chuyển PDF thành các hình ảnh chất lượng cao vào thư mục tạm
        # poppler_path có thể cần thiết trên Windows
        images = convert_from_path(input_path, dpi=300, fmt='jpeg', output_folder=temp_dir, thread_count=4) # Tăng tốc bằng threads

        if not images:
            raise ValueError("Không tìm thấy trang nào trong PDF hoặc không thể chuyển đổi thành ảnh.")

        # Tạo presentation
        prs = Presentation()

        # Thiết lập kích thước slide dựa trên PDF
        prs = setup_slide_size(prs, input_path)

        blank_layout = prs.slide_layouts[6]  # Layout trống (index 6 thường là blank)

        # Sắp xếp lại các file ảnh theo thứ tự trang (quan trọng!)
        # Tên file ảnh pdf2image tạo ra thường có dạng prefix-<page_number>.<fmt>
        image_files = sorted(
            [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith('.jpg') or f.lower().endswith('.jpeg')],
            key=lambda x: int(os.path.basename(x).split('-')[-1].split('.')[0]) # Sắp xếp theo số trang
        )


        if not image_files:
             raise ValueError("Không tìm thấy file ảnh nào trong thư mục tạm.")

        logger.info(f"Tìm thấy {len(image_files)} ảnh trang để thêm vào PPTX.")

        for image_path in image_files:
            try:
                with Image.open(image_path) as img:
                    img_width, img_height = img.size

                slide = prs.slides.add_slide(blank_layout)

                # Tính toán kích thước và vị trí để ảnh vừa slide và giữ tỷ lệ
                img_ratio = img_width / img_height
                slide_width_emu = prs.slide_width
                slide_height_emu = prs.slide_height
                slide_ratio = slide_width_emu / slide_height_emu

                if img_ratio > slide_ratio:
                    # Ảnh rộng hơn slide -> fit theo chiều rộng slide
                    pic_width = slide_width_emu
                    pic_height = int(pic_width / img_ratio)
                    # Căn giữa theo chiều dọc
                    pic_left = 0
                    pic_top = int((slide_height_emu - pic_height) / 2)
                else:
                    # Ảnh cao hơn slide (hoặc tỷ lệ bằng) -> fit theo chiều cao slide
                    pic_height = slide_height_emu
                    pic_width = int(pic_height * img_ratio)
                    # Căn giữa theo chiều ngang
                    pic_left = int((slide_width_emu - pic_width) / 2)
                    pic_top = 0

                # Thêm ảnh từ file path
                slide.shapes.add_picture(image_path, pic_left, pic_top, width=pic_width, height=pic_height)
            except Exception as page_err:
                 logger.warning(f"Lỗi khi thêm ảnh {os.path.basename(image_path)} vào slide: {page_err}. Bỏ qua trang này.")


        prs.save(output_path)
        logger.info(f"Đã lưu PPTX thành công tại: {output_path}")
        return True

    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang PPTX (phương pháp hình ảnh): {e}", exc_info=True) # Log traceback
        return False
    finally:
         # Dọn dẹp thư mục tạm chứa ảnh
         if temp_dir and os.path.exists(temp_dir):
             try:
                 shutil.rmtree(temp_dir)
                 logger.info(f"Đã xóa thư mục tạm ảnh: {temp_dir}")
             except Exception as cleanup_err:
                 logger.error(f"Lỗi khi xóa thư mục tạm ảnh {temp_dir}: {cleanup_err}")


def convert_pdf_to_pptx_python(input_path, output_path):
    """Chuyển PDF sang PPTX chỉ sử dụng phương pháp hình ảnh (ổn định hơn)"""
    logger.info("Thử chuyển đổi PDF -> PPTX bằng phương pháp hình ảnh (Python)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)


def convert_jpg_to_pdf(input_path, output_path):
    """Chuyển đổi JPG sang PDF"""
    try:
        image = Image.open(input_path)
        # Chuyển sang RGB nếu ảnh ở chế độ khác không được hỗ trợ trực tiếp bởi Pillow PDF save
        # (ví dụ: P, PA, LA, CMYK - mặc dù CMYK thường vẫn ổn)
        if image.mode not in ['RGB', 'L']: # L là Grayscale
            image = image.convert('RGB')
            logger.info(f"Đã chuyển đổi ảnh sang chế độ RGB từ {image.mode}")

        # Tạo PDF mới từ hình ảnh
        # resolution thường không ảnh hưởng nhiều khi chỉ có 1 ảnh
        image.save(output_path, "PDF", save_all=False) # save_all=False để chắc chắn chỉ lưu ảnh này
        return True
    except Exception as e:
        logger.error(f"Lỗi chuyển đổi JPG sang PDF: {e}", exc_info=True)
        return False


@app.route('/')
def index():
    """Trang chủ hiển thị form upload"""
    # Truyền URL của endpoint translations vào template
    translations_url = url_for('get_translations')
    return render_template('index.html', translations_url=translations_url)


@app.route('/convert', methods=['POST'])
def convert_file():
    """Xử lý chuyển đổi file"""
    input_path = None # Khởi tạo để dùng trong finally
    output_path = None # Khởi tạo để dùng trong finally
    temp_libreoffice_output = None # File output tạm của LibreOffice

    try:
        # Kiểm tra file upload
        if 'file' not in request.files:
            return "Không có file được tải lên", 400

        file = request.files['file']
        if not file or file.filename == '':
            return "Không có file được chọn", 400

        filename = secure_filename(file.filename)
        if not allowed_file(filename):
            # Lấy phần mở rộng và báo lỗi cụ thể hơn
            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'không có'
            allowed_str = ", ".join(ALLOWED_EXTENSIONS)
            return f"Loại file '{ext}' không hợp lệ. Chỉ chấp nhận: {allowed_str}", 400

        # Kiểm tra loại chuyển đổi được gửi từ form (đã được xác định rõ ràng ở frontend)
        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            return "Không chọn loại chuyển đổi cụ thể", 400

        logger.info(f"Yêu cầu chuyển đổi: file='{filename}', type='{conversion_type}'")

        # Chuẩn bị thư mục upload
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Lưu file upload
        input_path = os.path.join(UPLOAD_FOLDER, f"input_{time.time()}_{filename}") # Thêm timestamp để tránh trùng
        file.save(input_path)
        logger.info(f"File đã lưu: {input_path}")

        # Xác định tên file output dựa trên loại chuyển đổi
        base_name = filename.rsplit('.', 1)[0]
        if conversion_type == 'pdf_to_docx':
            out_ext = 'docx'
        elif conversion_type == 'docx_to_pdf':
            out_ext = 'pdf'
        elif conversion_type == 'pdf_to_ppt': # Luôn xuất ra pptx
            out_ext = 'pptx'
        elif conversion_type == 'ppt_to_pdf':
            out_ext = 'pdf'
        elif conversion_type == 'jpg_to_pdf':
            out_ext = 'pdf'
        # Bỏ qua pdf_to_jpg vì user yêu cầu không cần
        # elif conversion_type == 'pdf_to_jpg':
        #     out_ext = 'jpg'
        else:
            # Dọn dẹp file input nếu loại chuyển đổi không hợp lệ
            safe_remove(input_path)
            return "Loại chuyển đổi không hợp lệ hoặc không được hỗ trợ", 400

        output_filename = f"converted_{time.time()}_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        logger.info(f"File output dự kiến: {output_path}")

        # --- Thực hiện chuyển đổi ---
        conversion_success = False
        error_message = "Lỗi chuyển đổi không xác định"

        if conversion_type == 'pdf_to_docx':
            try:
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                conversion_success = True
                logger.info("Chuyển đổi PDF -> DOCX bằng pdf2docx thành công.")
            except Exception as e:
                error_message = f"Lỗi pdf2docx: {e}"
                logger.error(f"Lỗi chuyển đổi PDF -> DOCX bằng pdf2docx: {e}", exc_info=True)

        elif conversion_type == 'docx_to_pdf':
            try:
                # LibreOffice sẽ tạo file output với tên gốc + .pdf
                expected_lo_output_name = f"{base_name}.pdf"
                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)

                # Xóa file output có thể tồn tại từ lần chạy trước (quan trọng!)
                if os.path.exists(temp_libreoffice_output):
                     safe_remove(temp_libreoffice_output)

                logger.info(f"Chạy LibreOffice: {SOFFICE_PATH} --headless --convert-to pdf --outdir {UPLOAD_FOLDER} {input_path}")
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], check=True, timeout=120, capture_output=True, text=True) # Tăng timeout, capture output
                logger.info(f"LibreOffice stdout: {result.stdout}")
                logger.warning(f"LibreOffice stderr: {result.stderr}") # Log stderr phòng khi có warning

                # Kiểm tra xem file output có được tạo ra không
                if os.path.exists(temp_libreoffice_output):
                    os.rename(temp_libreoffice_output, output_path) # Đổi tên thành tên output chuẩn của chúng ta
                    conversion_success = True
                    logger.info(f"Chuyển đổi DOCX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                else:
                    error_message = "LibreOffice chạy xong nhưng không tạo file PDF output."
                    logger.error(error_message)
                    # Log thêm thông tin để debug
                    logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")


            except subprocess.CalledProcessError as e:
                error_message = f"Lỗi LibreOffice (CalledProcessError): {e}. Output: {e.stderr}"
                logger.error(error_message, exc_info=True)
            except subprocess.TimeoutExpired:
                 error_message = "Lỗi LibreOffice: Quá thời gian chuyển đổi (120s)."
                 logger.error(error_message)
            except Exception as e:
                error_message = f"Lỗi không xác định khi chạy LibreOffice: {e}"
                logger.error(error_message, exc_info=True)

        elif conversion_type == 'pdf_to_ppt':
             # Ưu tiên dùng phương pháp ảnh Python vì ổn định hơn
            if convert_pdf_to_pptx_python(input_path, output_path):
                 conversion_success = True
                 logger.info("Chuyển đổi PDF -> PPTX bằng Python (ảnh) thành công.")
            else:
                 # Nếu Python lỗi, thử dùng LibreOffice làm phương án cuối
                 logger.warning("Chuyển PDF->PPTX bằng Python thất bại, thử dùng LibreOffice...")
                 try:
                    expected_lo_output_name = f"{base_name}.pptx"
                    temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                    if os.path.exists(temp_libreoffice_output):
                        safe_remove(temp_libreoffice_output)

                    logger.info(f"Chạy LibreOffice: {SOFFICE_PATH} --headless --convert-to pptx --outdir {UPLOAD_FOLDER} {input_path}")
                    # Lưu ý: --infilter="draw_pdf_import" đôi khi gây lỗi, bỏ đi để thử mặc định
                    result = subprocess.run([
                        SOFFICE_PATH,
                        '--headless',
                        # '--infilter="impress_pdf_import"', # Thử filter này hoặc bỏ trống
                        '--convert-to', 'pptx',
                        '--outdir', UPLOAD_FOLDER,
                        input_path
                    ], check=True, timeout=180, capture_output=True, text=True) # Tăng timeout cho PPTX
                    logger.info(f"LibreOffice stdout: {result.stdout}")
                    logger.warning(f"LibreOffice stderr: {result.stderr}")

                    if os.path.exists(temp_libreoffice_output):
                        os.rename(temp_libreoffice_output, output_path)
                        conversion_success = True
                        logger.info(f"Chuyển đổi PDF -> PPTX bằng LibreOffice thành công (fallback). Output: {output_path}")
                    else:
                        error_message = "LibreOffice (fallback) chạy xong nhưng không tạo file PPTX output."
                        logger.error(error_message)
                        logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                        logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")

                 except subprocess.CalledProcessError as e:
                     error_message = f"Lỗi LibreOffice (fallback PDF->PPTX): {e}. Output: {e.stderr}"
                     logger.error(error_message, exc_info=True)
                 except subprocess.TimeoutExpired:
                      error_message = "Lỗi LibreOffice (fallback PDF->PPTX): Quá thời gian chuyển đổi (180s)."
                      logger.error(error_message)
                 except Exception as e:
                     error_message = f"Lỗi không xác định khi chạy LibreOffice (fallback PDF->PPTX): {e}"
                     logger.error(error_message, exc_info=True)

                 # Nếu cả 2 cách đều lỗi
                 if not conversion_success:
                     error_message = "Cả phương pháp Python và LibreOffice đều thất bại khi chuyển đổi PDF -> PPTX."


        elif conversion_type == 'ppt_to_pdf':
            try:
                expected_lo_output_name = f"{base_name}.pdf"
                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                if os.path.exists(temp_libreoffice_output):
                     safe_remove(temp_libreoffice_output)

                logger.info(f"Chạy LibreOffice: {SOFFICE_PATH} --headless --convert-to pdf --outdir {UPLOAD_FOLDER} {input_path}")
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], check=True, timeout=120, capture_output=True, text=True)
                logger.info(f"LibreOffice stdout: {result.stdout}")
                logger.warning(f"LibreOffice stderr: {result.stderr}")

                if os.path.exists(temp_libreoffice_output):
                    os.rename(temp_libreoffice_output, output_path)
                    conversion_success = True
                    logger.info(f"Chuyển đổi PPT/PPTX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                else:
                    error_message = "LibreOffice chạy xong nhưng không tạo file PDF output (từ PPT)."
                    logger.error(error_message)
                    logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")

            except subprocess.CalledProcessError as e:
                error_message = f"Lỗi LibreOffice (PPT->PDF): {e}. Output: {e.stderr}"
                logger.error(error_message, exc_info=True)
            except subprocess.TimeoutExpired:
                 error_message = "Lỗi LibreOffice (PPT->PDF): Quá thời gian chuyển đổi (120s)."
                 logger.error(error_message)
            except Exception as e:
                error_message = f"Lỗi không xác định khi chạy LibreOffice (PPT->PDF): {e}"
                logger.error(error_message, exc_info=True)


        elif conversion_type == 'jpg_to_pdf':
             if convert_jpg_to_pdf(input_path, output_path):
                 conversion_success = True
                 logger.info("Chuyển đổi JPG -> PDF thành công.")
             else:
                 error_message = "Chuyển đổi JPG sang PDF thất bại bằng Pillow."
                 # Lỗi cụ thể đã được log trong hàm convert_jpg_to_pdf


        # --- Xử lý kết quả ---
        if conversion_success and os.path.exists(output_path):
            # Trả về file đã chuyển đổi
            try:
                # Ghi log kích thước file output
                output_size = os.path.getsize(output_path)
                logger.info(f"Chuẩn bị gửi file output: {output_path}, Kích thước: {output_size} bytes")

                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename # Đặt tên file khi tải về
                    # mimetype được send_file tự động xác định hoặc bạn có thể chỉ định nếu cần
                )
            except Exception as send_err:
                 logger.error(f"Lỗi khi gửi file {output_path}: {send_err}", exc_info=True)
                 # Dù lỗi gửi file thì cũng nên báo lỗi server
                 return f"Lỗi khi chuẩn bị file để tải về: {send_err}", 500
            finally:
                # Dọn dẹp file input SAU KHI gửi file thành công hoặc thất bại
                safe_remove(input_path)
                # Không xóa output_path ở đây vì send_file cần nó
                # Nó sẽ được xóa bởi @after_this_request hoặc teardown_appcontext
        else:
            # Nếu chuyển đổi thất bại, báo lỗi 500
            logger.error(f"Chuyển đổi thất bại cuối cùng. Lý do: {error_message}")
            # Dọn dẹp file input nếu chuyển đổi lỗi
            safe_remove(input_path)
            # Cũng nên dọn dẹp output_path nếu nó vô tình được tạo ra nhưng không hợp lệ
            if output_path and os.path.exists(output_path):
                 safe_remove(output_path)
            # Xóa cả file tạm của LO nếu có
            if temp_libreoffice_output and os.path.exists(temp_libreoffice_output):
                 safe_remove(temp_libreoffice_output)

            return f"Chuyển đổi thất bại: {error_message}", 500

    except Exception as e:
        # Bắt các lỗi không mong muốn khác (ví dụ: lỗi đọc form, lưu file ban đầu)
        logger.error(f"Lỗi không mong muốn trong route /convert: {e}", exc_info=True)
        # Dọn dẹp file input/output nếu có thể trong trường hợp lỗi sớm
        if input_path and os.path.exists(input_path):
             safe_remove(input_path)
        if output_path and os.path.exists(output_path):
             safe_remove(output_path)
        if temp_libreoffice_output and os.path.exists(temp_libreoffice_output):
             safe_remove(temp_libreoffice_output)
        return f"Đã xảy ra lỗi máy chủ không mong muốn: {str(e)}", 500

    # Khối finally ở đây không cần thiết nữa vì việc dọn dẹp đã xử lý trong try/except và @after_this_request

@app.after_request
def after_request_func(response):
    """Dọn dẹp file output sau khi request gửi file hoàn tất"""
    # Kiểm tra xem request có phải là gửi file không (dựa vào header hoặc context khác)
    # Cách đơn giản là kiểm tra xem có đường dẫn file output nào được lưu trữ tạm không
    # Tuy nhiên, cách này không an toàn. Dùng teardown_appcontext đáng tin cậy hơn.
    # Thay vào đó, chúng ta sẽ dựa vào teardown_appcontext để dọn dẹp tổng quát.
    return response


@app.teardown_appcontext
def cleanup(exception=None):
    """Dọn dẹp các file cũ trong thư mục upload khi context kết thúc"""
    if not os.path.exists(UPLOAD_FOLDER):
        return

    logger.info("Chạy dọn dẹp teardown_appcontext...")
    try:
        now = time.time()
        # Thời gian giữ file (vd: 1 giờ = 3600 giây)
        # Nên giữ lâu hơn một chút để đảm bảo file vừa tạo không bị xóa ngay
        max_age = 7200 # 2 giờ

        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                # Chỉ xóa file (không xóa thư mục con nếu có)
                if os.path.isfile(path):
                     file_age = now - os.path.getmtime(path)
                     # Xóa file nếu nó cũ hơn max_age VÀ không phải file đang được xử lý (khó xác định chắc chắn)
                     # Cách an toàn là chỉ xóa file cũ
                     if file_age > max_age:
                         logger.info(f"Teardown: Xóa file cũ ({file_age:.0f}s > {max_age}s): {path}")
                         safe_remove(path)
                     # else:
                     #     logger.info(f"Teardown: Giữ lại file gần đây ({file_age:.0f}s <= {max_age}s): {path}")

            except FileNotFoundError:
                 continue # File đã bị xóa bởi tiến trình khác
            except Exception as e:
                logger.error(f"Lỗi khi kiểm tra/dọn dẹp file {path}: {e}")
    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng trong quá trình dọn dẹp teardown_appcontext: {e}")


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    logger.info(f"Khởi động server trên cổng {port} - Chế độ Debug: {debug_mode}")
    # Chạy với reloader nếu debug=True để tự động cập nhật khi code thay đổi
    app.run(host='0.0.0.0', port=port, debug=debug_mode)
# --- END OF FILE app.py ---