# --- START OF FILE app.py ---

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
from pdf2image import convert_from_path, pdfinfo_from_path # Import pdfinfo_from_path
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image
from docx import Document
import zipfile # Import zipfile

app = Flask(__name__, template_folder='templates', static_folder='static')

app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
# Add jpg and jpeg to allowed extensions
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}

@app.route('/health')
def health_check():
    return 'OK', 200

@app.route('/get_translations')
def get_translations():
    translations = {
        'en': {
            'lang-title': 'PDF Tools',
            'lang-subtitle': 'Simple, powerful PDF tools for everyone',
            'lang-error-title': 'Error!',
            'lang-convert-title': 'Convert Files', # Updated title slightly
            'lang-convert-desc': 'Transform PDFs, Office documents, and images', # Updated description
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
            'err-format-jpg': 'File format not compatible with PDF ↔ JPG conversion (Only .pdf, .jpg, .jpeg allowed).', # Updated JPG error
            'err-conversion': 'An error occurred during conversion.',
            'err-fetch-translations': 'Could not load language data.',
            'lang-select-btn-text': 'Browse',
            'lang-select-conversion-label': 'Conversion Type'
        },
        'vi': {
            'lang-title': 'Công Cụ PDF & Văn Phòng',
            'lang-subtitle': 'Công cụ PDF, Office, Hình ảnh đơn giản, mạnh mẽ',
            'lang-error-title': 'Lỗi!',
            'lang-convert-title': 'Chuyển đổi Tệp', # Updated title
            'lang-convert-desc': 'Chuyển đổi PDF, tài liệu Office và hình ảnh', # Updated description
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
            'err-format-jpg': 'Định dạng tệp không phù hợp với kiểu chuyển đổi PDF ↔ JPG (Chỉ chấp nhận .pdf, .jpg, .jpeg).', # Updated JPG error
            'err-conversion': 'Đã xảy ra lỗi trong quá trình chuyển đổi.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.',
            'lang-select-btn-text': 'Duyệt...',
            'lang-select-conversion-label': 'Kiểu chuyển đổi'
        }
    }
    lang = request.args.get('lang', 'en')
    return jsonify(translations.get(lang, translations['en']))

def find_libreoffice():
    possible_paths = [
        'soffice', '/usr/bin/soffice', '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice', '/usr/lib/libreoffice/program/soffice',
        # Windows paths (less common on servers, but good to have)
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
    ]
    for path in possible_paths:
        try:
            resolved_path = shutil.which(path) # Check PATH first
            if resolved_path and os.path.isfile(resolved_path):
                # Use resolved_path for check
                result = subprocess.run([resolved_path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0:
                    logger.info(f"Tìm thấy LibreOffice tại (PATH): {resolved_path}")
                    return resolved_path
            # If not in PATH, check the specific path directly
            elif os.path.isfile(path):
                result = subprocess.run([path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0:
                    logger.info(f"Tìm thấy LibreOffice tại (Direct Path): {path}")
                    return path
        except FileNotFoundError:
            logger.debug(f"Không tìm thấy LibreOffice tại {path} hoặc qua shutil.which")
        except subprocess.TimeoutExpired:
            logger.warning(f"Kiểm tra LibreOffice tại {path} bị timeout.")
        except Exception as e:
            logger.warning(f"Lỗi khi kiểm tra LibreOffice tại {path}: {e}")
    logger.warning("Không tìm thấy LibreOffice thực thi qua các đường dẫn phổ biến hoặc PATH. Sử dụng 'soffice' mặc định.")
    return 'soffice'


SOFFICE_PATH = find_libreoffice()
logger.info(f"Sử dụng đường dẫn LibreOffice: {SOFFICE_PATH}")

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def safe_remove(file_path, retries=5, delay=1):
    """Xóa file hoặc thư mục an toàn với nhiều lần thử"""
    if not file_path or not os.path.exists(file_path):
        # logger.debug(f"File/Dir không tồn tại, không cần xóa: {file_path}")
        return True

    for i in range(retries):
        try:
            if os.path.isdir(file_path):
                shutil.rmtree(file_path)
                logger.info(f"Đã xóa thư mục tạm: {file_path}")
            else:
                os.remove(file_path)
                logger.info(f"Đã xóa file tạm: {file_path}")
            return True
        except PermissionError:
            logger.warning(f"Không có quyền xóa {file_path} (lần thử {i + 1}). Đang đợi...")
            time.sleep(delay* (i + 1)) # Increase delay
        except Exception as e:
            logger.warning(f"Không thể xóa {file_path} (lần thử {i + 1}): {e}")
            time.sleep(delay)
    logger.error(f"Xóa {file_path} thất bại sau {retries} lần thử.")
    return False

def get_pdf_page_size(pdf_path):
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if not reader.pages:
                logger.warning(f"PDF không có trang nào: {pdf_path}")
                return None, None
            page = reader.pages[0]
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
    pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)
    if pdf_width_pt is None or pdf_height_pt is None:
        logger.warning("Không thể đọc kích thước PDF, sử dụng kích thước mặc định (10x7.5 inches)")
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        return prs
    try:
        pdf_width_in = pdf_width_pt / 72
        pdf_height_in = pdf_height_pt / 72
        max_slide_dim = 50.0 # PowerPoint max slide dimension
        if pdf_width_in > max_slide_dim or pdf_height_in > max_slide_dim:
            ratio = pdf_width_in / pdf_height_in
            if pdf_width_in >= pdf_height_in:
                prs.slide_width = Inches(max_slide_dim)
                prs.slide_height = Inches(max_slide_dim / ratio)
            else:
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

def _convert_pdf_to_pptx_images(input_path, output_path):
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_")
        logger.info(f"Tạo thư mục tạm cho ảnh PPTX: {temp_dir}")
        # Check if Poppler is likely available by trying pdfinfo first
        try:
            pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None)
            logger.info(f"PDF info read successfully (Poppler likely present). Pages: {pdfinfo.get('Pages', 'N/A')}")
        except Exception as info_err:
             # Log specific error if pdfinfo fails (likely Poppler path issue)
             logger.error(f"Lỗi khi chạy pdfinfo (kiểm tra Poppler): {info_err}. Kiểm tra xem Poppler đã được cài đặt và trong PATH chưa.", exc_info=True)
             # Re-raise a more user-friendly error or return False directly
             raise ValueError("Could not process PDF, Poppler might be missing or not configured correctly.") from info_err

        images = convert_from_path(input_path, dpi=300, fmt='jpeg', output_folder=temp_dir, thread_count=4)
        if not images:
            raise ValueError("Không tìm thấy trang nào trong PDF hoặc không thể chuyển đổi thành ảnh.")

        prs = Presentation()
        prs = setup_slide_size(prs, input_path)
        blank_layout = prs.slide_layouts[6]

        # Correctly sort image files based on page number if pdf2image naming is consistent
        # Example filename: f'{output_file}-{(page_number+first_page):0{len(str(page_count+first_page))}}.ppm'
        # We expect something like '...-01.jpg', '...-02.jpg'
        image_files = sorted(
            [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))],
             key=lambda x: int(os.path.splitext(os.path.basename(x))[0].split('-')[-1]) # Extract page number reliably
        )

        if not image_files:
            raise ValueError("Không tìm thấy file ảnh nào trong thư mục tạm.")
        logger.info(f"Tìm thấy {len(image_files)} ảnh trang để thêm vào PPTX.")

        for image_path in image_files:
            try:
                with Image.open(image_path) as img:
                    img_width, img_height = img.size
                slide = prs.slides.add_slide(blank_layout)
                img_ratio = img_width / img_height
                slide_width_emu = prs.slide_width
                slide_height_emu = prs.slide_height
                slide_ratio = slide_width_emu / slide_height_emu

                if img_ratio > slide_ratio:
                    pic_width = slide_width_emu
                    pic_height = int(pic_width / img_ratio)
                    pic_left = 0
                    pic_top = int((slide_height_emu - pic_height) / 2)
                else:
                    pic_height = slide_height_emu
                    pic_width = int(pic_height * img_ratio)
                    pic_left = int((slide_width_emu - pic_width) / 2)
                    pic_top = 0
                slide.shapes.add_picture(image_path, pic_left, pic_top, width=pic_width, height=pic_height)
            except Exception as page_err:
                logger.warning(f"Lỗi khi thêm ảnh {os.path.basename(image_path)} vào slide: {page_err}. Bỏ qua trang này.")

        prs.save(output_path)
        logger.info(f"Đã lưu PPTX thành công tại: {output_path}")
        return True
    except ValueError as ve: # Catch specific errors like Poppler missing
        logger.error(f"Lỗi giá trị khi chuyển đổi PDF sang PPTX (hình ảnh): {ve}", exc_info=True)
        # Propagate a clearer message if it's the Poppler issue
        if "Poppler might be missing" in str(ve):
             raise ValueError("PDF conversion failed: Poppler utility not found or accessible. Please ensure Poppler is installed and in your system's PATH.") from ve
        else:
             raise # Re-raise other ValueErrors
    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang PPTX (phương pháp hình ảnh): {e}", exc_info=True)
        # Check if it's a Poppler error disguised as something else
        if "pdfinfo" in str(e) or "Poppler" in str(e):
             raise ValueError("PDF conversion failed: Error interacting with Poppler. Please ensure Poppler is installed and in your system's PATH.") from e
        return False # Return False for other general exceptions
    finally:
        safe_remove(temp_dir) # Use safe_remove for the directory

def convert_pdf_to_pptx_python(input_path, output_path):
    logger.info("Thử chuyển đổi PDF -> PPTX bằng phương pháp hình ảnh (Python)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)

def convert_jpg_to_pdf(input_path, output_path):
    """Chuyển đổi JPG/JPEG sang PDF"""
    try:
        image = Image.open(input_path)
        # Convert image to RGB if it's not (e.g., RGBA, P, CMYK)
        # This prevents potential errors during PDF saving for some modes.
        if image.mode == 'RGBA':
             # Create a white background image
             bg = Image.new('RGB', image.size, (255, 255, 255))
             # Paste the RGBA image onto the white background
             bg.paste(image, (0, 0), image)
             image = bg
             logger.info(f"Đã chuyển đổi ảnh RGBA sang RGB bằng cách thêm nền trắng.")
        elif image.mode == 'P':
            image = image.convert('RGB')
            logger.info(f"Đã chuyển đổi ảnh Palette (P) sang RGB.")
        elif image.mode not in ['RGB', 'L']: # L is grayscale, usually fine
            image = image.convert('RGB')
            logger.info(f"Đã chuyển đổi ảnh chế độ {image.mode} sang RGB.")

        image.save(output_path, "PDF", resolution=100.0, save_all=False) # resolution is optional
        logger.info(f"Đã chuyển đổi JPG -> PDF thành công: {output_path}")
        return True
    except Exception as e:
        logger.error(f"Lỗi chuyển đổi JPG sang PDF: {e}", exc_info=True)
        return False

def convert_pdf_to_jpg_zip(input_path, output_zip_path):
    """Chuyển đổi PDF sang nhiều file JPG và nén thành ZIP"""
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdf2jpg_")
        logger.info(f"Tạo thư mục tạm cho JPG: {temp_dir}")

        # Check if Poppler is likely available by trying pdfinfo first
        try:
            pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None)
            page_count = pdfinfo.get('Pages', 0)
            logger.info(f"PDF info read successfully (Poppler likely present). Pages: {page_count}")
            if page_count == 0:
                 logger.warning("PDF reported 0 pages.")
                 # Still try conversion, maybe pdfinfo was wrong
        except Exception as info_err:
             logger.error(f"Lỗi khi chạy pdfinfo (kiểm tra Poppler): {info_err}. Kiểm tra xem Poppler đã được cài đặt và trong PATH chưa.", exc_info=True)
             raise ValueError("Could not process PDF, Poppler might be missing or not configured correctly.") from info_err

        # Generate images in the temp directory
        # Use a base name for output files inside the temp dir
        output_base_name = os.path.splitext(os.path.basename(input_path))[0]
        images = convert_from_path(
            input_path,
            dpi=200,  # Adjust DPI as needed
            fmt='jpeg',
            output_folder=temp_dir,
            output_file=output_base_name, # Use base name for generated files
            thread_count=4
        )

        if not images:
            # Double check if files were created even if images list is empty
            image_files = [f for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))]
            if not image_files:
                 raise ValueError("Không tạo được ảnh nào từ PDF.")
            else:
                 logger.warning("convert_from_path trả về list rỗng nhưng file ảnh đã được tạo.")

        # Get list of created jpg files
        image_files = sorted(
             [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))]
             # Sort numerically based on the part after the last '-'
             # key=lambda x: int(os.path.splitext(os.path.basename(x))[0].split('-')[-1])
        )
        if not image_files:
              raise ValueError("Không tìm thấy file JPG nào trong thư mục tạm sau khi chuyển đổi.")

        logger.info(f"Đã tạo {len(image_files)} file JPG trong {temp_dir}")

        # Create ZIP file
        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for i, file_path in enumerate(image_files):
                # Add file to zip, using a simpler name inside the zip (e.g., page_1.jpg)
                arcname = f"page_{i+1}.jpg"
                zipf.write(file_path, arcname=arcname)
                # logger.debug(f"Adding {file_path} as {arcname} to zip.")

        logger.info(f"Đã tạo file ZIP thành công: {output_zip_path}")
        return True

    except ValueError as ve: # Catch specific errors like Poppler missing
        logger.error(f"Lỗi giá trị khi chuyển đổi PDF sang JPG: {ve}", exc_info=True)
        if "Poppler might be missing" in str(ve):
             raise ValueError("PDF conversion failed: Poppler utility not found or accessible. Please ensure Poppler is installed and in your system's PATH.") from ve
        else:
             raise # Re-raise other ValueErrors
    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang JPG/ZIP: {e}", exc_info=True)
        if "pdfinfo" in str(e) or "Poppler" in str(e):
             raise ValueError("PDF conversion failed: Error interacting with Poppler. Please ensure Poppler is installed and in your system's PATH.") from e
        return False
    finally:
        safe_remove(temp_dir) # Use safe_remove for the directory


@app.route('/')
def index():
    translations_url = url_for('get_translations')
    return render_template('index.html', translations_url=translations_url)

@app.route('/convert', methods=['POST'])
def convert_file():
    input_path = None
    output_path = None
    temp_dir = None # For potential intermediate files if not handled by specific functions
    temp_libreoffice_output = None # For LO specific temp output

    try:
        if 'file' not in request.files:
            return "Không có file được tải lên", 400
        file = request.files['file']
        if not file or file.filename == '':
            return "Không có file được chọn", 400

        filename = secure_filename(file.filename)
        if not allowed_file(filename):
            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'không có'
            allowed_str = ", ".join(ALLOWED_EXTENSIONS)
            # Use translation key for error message
            error_message_key = 'err-format-invalid' # Generic invalid format key (needs adding to translations)
            # Or be more specific if needed
            return f"Loại file '{ext}' không hợp lệ. Chỉ chấp nhận: {allowed_str}", 400 # Keep this for now

        # Get the *actual* conversion type requested by the user's selection
        # Note: JS now sends the specific type like 'pdf_to_docx', 'docx_to_pdf' etc.
        actual_conversion_type = request.form.get('conversion_type')
        if not actual_conversion_type:
            return "Không chọn loại chuyển đổi cụ thể", 400

        logger.info(f"Yêu cầu chuyển đổi: file='{filename}', type='{actual_conversion_type}'")

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        # Create a unique input path
        input_path = os.path.join(UPLOAD_FOLDER, f"input_{time.time()}_{filename}")
        file.save(input_path)
        logger.info(f"File đã lưu: {input_path}")

        base_name = filename.rsplit('.', 1)[0]
        # Determine output extension based on the *actual* conversion type
        if actual_conversion_type == 'pdf_to_docx':
            out_ext = 'docx'
        elif actual_conversion_type == 'docx_to_pdf':
            out_ext = 'pdf'
        elif actual_conversion_type == 'pdf_to_ppt':
            out_ext = 'pptx'
        elif actual_conversion_type == 'ppt_to_pdf':
            out_ext = 'pdf'
        elif actual_conversion_type == 'jpg_to_pdf':
            out_ext = 'pdf'
        elif actual_conversion_type == 'pdf_to_jpg':
            out_ext = 'zip' # Output is a zip file
        else:
            safe_remove(input_path)
            return "Loại chuyển đổi không hợp lệ hoặc không được hỗ trợ", 400

        # Generate output path
        output_filename = f"converted_{time.time()}_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        logger.info(f"File output dự kiến: {output_path}")

        conversion_success = False
        error_message = "Lỗi chuyển đổi không xác định"

        # --- Conversion Logic ---
        try:
            if actual_conversion_type == 'pdf_to_docx':
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                conversion_success = True
                logger.info("Chuyển đổi PDF -> DOCX bằng pdf2docx thành công.")

            elif actual_conversion_type == 'docx_to_pdf':
                expected_lo_output_name = os.path.basename(input_path).replace('.docx', '.pdf')
                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                safe_remove(temp_libreoffice_output) # Ensure it's clean before run

                logger.info(f"Chạy LibreOffice: {SOFFICE_PATH} --headless --convert-to pdf --outdir {UPLOAD_FOLDER} {input_path}")
                result = subprocess.run([
                    SOFFICE_PATH, '--headless', '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER, input_path
                ], check=True, timeout=120, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                logger.info(f"LibreOffice stdout: {result.stdout}")
                logger.warning(f"LibreOffice stderr: {result.stderr}")

                if os.path.exists(temp_libreoffice_output):
                    # Check size > 0? Basic sanity check
                    if os.path.getsize(temp_libreoffice_output) > 0:
                        os.rename(temp_libreoffice_output, output_path)
                        conversion_success = True
                        logger.info(f"Chuyển đổi DOCX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                    else:
                         error_message = "LibreOffice đã tạo file PDF output nhưng file bị rỗng."
                         logger.error(error_message + f" Path: {temp_libreoffice_output}")
                         safe_remove(temp_libreoffice_output) # Clean up empty file
                else:
                    error_message = "LibreOffice chạy xong nhưng không tạo file PDF output."
                    logger.error(error_message)
                    logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")

            elif actual_conversion_type == 'pdf_to_ppt':
                # Try Python method first
                try:
                    if convert_pdf_to_pptx_python(input_path, output_path):
                        conversion_success = True
                        logger.info("Chuyển đổi PDF -> PPTX bằng Python (ảnh) thành công.")
                    else:
                         # If python method returns False explicitly, don't fallback immediately
                         # It might have logged a specific non-Poppler error
                         error_message = "Phương pháp Python (ảnh) thất bại khi chuyển đổi PDF -> PPTX."
                         logger.warning(error_message)
                         # Maybe fallback here IS desired? Let's keep fallback for now.
                         raise RuntimeError("Python PPTX conversion failed, attempting LibreOffice fallback.")

                except ValueError as ve_ppt: # Catch Poppler errors from convert_pdf_to_pptx_python
                    error_message = f"Lỗi chuyển đổi PDF -> PPTX: {ve_ppt}"
                    logger.error(error_message, exc_info=True)
                    # Don't fallback if it's a Poppler issue, LO likely won't help either?
                    # Let's try fallback anyway, LO might use a different mechanism.
                    logger.warning("Lỗi Python PPTX, thử dùng LibreOffice...")
                    # Fall through to LibreOffice block below

                except Exception as py_ppt_err: # Catch other Python errors
                     error_message = f"Lỗi Python khi chuyển PDF -> PPTX: {py_ppt_err}"
                     logger.error(error_message, exc_info=True)
                     logger.warning("Thử dùng LibreOffice fallback...")
                     # Fall through to LibreOffice block

                # --- LibreOffice Fallback Block (only reached if Python fails/raises) ---
                if not conversion_success:
                    expected_lo_output_name = os.path.basename(input_path).replace('.pdf', '.pptx')
                    temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                    safe_remove(temp_libreoffice_output)

                    logger.info(f"Chạy LibreOffice (Fallback PDF->PPTX): {SOFFICE_PATH} --headless --convert-to pptx --outdir {UPLOAD_FOLDER} {input_path}")
                    result = subprocess.run([
                        SOFFICE_PATH, '--headless', '--convert-to', 'pptx',
                        '--outdir', UPLOAD_FOLDER, input_path
                    ], check=True, timeout=180, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                    logger.info(f"LibreOffice stdout: {result.stdout}")
                    logger.warning(f"LibreOffice stderr: {result.stderr}")

                    if os.path.exists(temp_libreoffice_output):
                         if os.path.getsize(temp_libreoffice_output) > 0:
                            os.rename(temp_libreoffice_output, output_path)
                            conversion_success = True
                            logger.info(f"Chuyển đổi PDF -> PPTX bằng LibreOffice thành công (fallback). Output: {output_path}")
                         else:
                            error_message = "LibreOffice (fallback) đã tạo file PPTX output nhưng file bị rỗng."
                            logger.error(error_message + f" Path: {temp_libreoffice_output}")
                            safe_remove(temp_libreoffice_output)
                    else:
                        error_message = "LibreOffice (fallback) chạy xong nhưng không tạo file PPTX output."
                        logger.error(error_message)
                        logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                        logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")

                    if not conversion_success and "Python PPTX conversion failed" not in error_message:
                         # Update error message only if it wasn't set by the initial python failure
                         error_message = "Cả phương pháp Python và LibreOffice đều thất bại khi chuyển đổi PDF -> PPTX."


            elif actual_conversion_type == 'ppt_to_pdf':
                expected_lo_output_name = os.path.basename(input_path)
                if expected_lo_output_name.lower().endswith('.pptx'):
                    expected_lo_output_name = expected_lo_output_name[:-5] + '.pdf'
                elif expected_lo_output_name.lower().endswith('.ppt'):
                    expected_lo_output_name = expected_lo_output_name[:-4] + '.pdf'
                else: # Should not happen due to allowed_extensions but handle defensively
                     expected_lo_output_name += '.pdf'

                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name)
                safe_remove(temp_libreoffice_output)

                logger.info(f"Chạy LibreOffice: {SOFFICE_PATH} --headless --convert-to pdf --outdir {UPLOAD_FOLDER} {input_path}")
                result = subprocess.run([
                    SOFFICE_PATH, '--headless', '--convert-to', 'pdf',
                    '--outdir', UPLOAD_FOLDER, input_path
                ], check=True, timeout=120, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                logger.info(f"LibreOffice stdout: {result.stdout}")
                logger.warning(f"LibreOffice stderr: {result.stderr}")

                if os.path.exists(temp_libreoffice_output):
                    if os.path.getsize(temp_libreoffice_output) > 0:
                        os.rename(temp_libreoffice_output, output_path)
                        conversion_success = True
                        logger.info(f"Chuyển đổi PPT/PPTX -> PDF bằng LibreOffice thành công. Output: {output_path}")
                    else:
                         error_message = "LibreOffice đã tạo file PDF output (từ PPT) nhưng file bị rỗng."
                         logger.error(error_message + f" Path: {temp_libreoffice_output}")
                         safe_remove(temp_libreoffice_output)
                else:
                    error_message = "LibreOffice chạy xong nhưng không tạo file PDF output (từ PPT)."
                    logger.error(error_message)
                    logger.error(f"Thư mục output của LO: {UPLOAD_FOLDER}, Tên file LO dự kiến: {expected_lo_output_name}")
                    logger.error(f"Nội dung thư mục upload sau khi chạy LO: {os.listdir(UPLOAD_FOLDER)}")

            elif actual_conversion_type == 'jpg_to_pdf':
                if convert_jpg_to_pdf(input_path, output_path):
                    conversion_success = True
                else:
                    error_message = "Chuyển đổi JPG sang PDF thất bại bằng Pillow."

            elif actual_conversion_type == 'pdf_to_jpg':
                 # This function now raises specific errors (e.g., for Poppler)
                 if convert_pdf_to_jpg_zip(input_path, output_path):
                     conversion_success = True
                     logger.info("Chuyển đổi PDF -> JPG (ZIP) thành công.")
                 else:
                     # convert_pdf_to_jpg_zip returns False for general errors
                     # Specific ValueErrors (like Poppler) are raised and caught below
                     error_message = "Chuyển đổi PDF sang JPG/ZIP thất bại (lỗi không xác định)."


        # --- Catch specific conversion errors (e.g., subprocess, Poppler) ---
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as sub_err:
             stderr_output = sub_err.stderr if hasattr(sub_err, 'stderr') else '(Không có stderr)'
             error_message = f"Lỗi thực thi tiến trình ({type(sub_err).__name__}): {sub_err}. Output: {stderr_output}"
             logger.error(f"Lỗi xử lý {actual_conversion_type}: {error_message}", exc_info=True)
        except ValueError as val_err: # Catch Poppler/Image processing errors
             error_message = f"Lỗi dữ liệu hoặc cấu hình: {val_err}"
             logger.error(f"Lỗi xử lý {actual_conversion_type}: {error_message}", exc_info=True)
             # Check if it's the Poppler specific message we crafted
             if "Poppler utility not found" in str(val_err):
                 # You could return a more specific error code or message here if needed
                 error_message = str(val_err) # Use the specific message
        except FileNotFoundError as fnf_err: # E.g., soffice not found
             error_message = f"Không tìm thấy file hoặc chương trình cần thiết: {fnf_err}"
             logger.error(f"Lỗi xử lý {actual_conversion_type}: {error_message}", exc_info=True)
             if SOFFICE_PATH in str(fnf_err): # Specifically if LibreOffice wasn't found
                  error_message = "Lỗi: Không tìm thấy LibreOffice. Chuyển đổi phụ thuộc vào LibreOffice đã thất bại."
        except ImportError as imp_err: # Handle missing optional dependencies if any added later
             error_message = f"Thiếu thư viện Python cần thiết: {imp_err}"
             logger.error(f"Lỗi xử lý {actual_conversion_type}: {error_message}", exc_info=True)
        except Exception as conv_err: # Catch any other conversion-related error
             error_message = f"Lỗi không xác định trong quá trình chuyển đổi {actual_conversion_type}: {conv_err}"
             logger.error(error_message, exc_info=True)

        # --- Handling result ---
        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            try:
                logger.info(f"Chuẩn bị gửi file output: {output_path}, Kích thước: {os.path.getsize(output_path)} bytes")
                # Determine mimetype for send_file
                mimetype = None
                if out_ext == 'zip':
                    mimetype = 'application/zip'
                elif out_ext == 'pdf':
                    mimetype = 'application/pdf'
                elif out_ext == 'docx':
                    mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                elif out_ext == 'pptx':
                    mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

                response = send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype=mimetype
                )
                # Add cleanup *after* file is sent using response.call_on_close
                # Use a lambda to capture the paths needed for cleanup
                @response.call_on_close
                def cleanup_files():
                    logger.info("Cleanup after sending file...")
                    safe_remove(input_path)
                    safe_remove(output_path)
                    # Also clean up LO temp file if it exists and wasn't the final output
                    if temp_libreoffice_output and temp_libreoffice_output != output_path:
                         safe_remove(temp_libreoffice_output)

                return response

            except Exception as send_err:
                logger.error(f"Lỗi khi gửi file {output_path}: {send_err}", exc_info=True)
                # Fall through to failure case below after logging
                error_message = f"Lỗi khi chuẩn bị file để tải về: {send_err}"
                conversion_success = False # Ensure we hit the failure block

        # --- Failure Case ---
        if not conversion_success:
            logger.error(f"Chuyển đổi thất bại cuối cùng cho {actual_conversion_type}. Lý do: {error_message}")
            # Cleanup all potentially created files on failure
            safe_remove(input_path)
            safe_remove(output_path) # Remove potentially empty/corrupt output
            safe_remove(temp_dir) # Remove intermediate dirs if any
            if temp_libreoffice_output:
                 safe_remove(temp_libreoffice_output) # Remove LO temp file
            # Return the error message to the user
            # Prepend "Error: " to make it clearer on the frontend?
            return f"Chuyển đổi thất bại: {error_message}", 500

    except Exception as e:
        # Catch-all for unexpected errors in the route handler itself
        logger.error(f"Lỗi không mong muốn trong route /convert: {e}", exc_info=True)
        # Attempt cleanup of any known paths
        safe_remove(input_path)
        safe_remove(output_path)
        safe_remove(temp_dir)
        if temp_libreoffice_output:
             safe_remove(temp_libreoffice_output)
        return f"Đã xảy ra lỗi máy chủ không mong muốn: {str(e)}", 500

# @app.after_request
# def after_request_func(response):
#     # Removed cleanup here, using response.call_on_close in send_file now
#     return response

@app.teardown_appcontext
def cleanup_old_files(exception=None):
    """Dọn dẹp các file CŨ trong thư mục upload khi context kết thúc"""
    if not os.path.exists(UPLOAD_FOLDER):
        return
    logger.info("Chạy dọn dẹp file cũ teardown_appcontext...")
    try:
        now = time.time()
        # Keep files for longer? E.g., 2 hours = 7200 seconds
        # Shorten for testing: 1 hour = 3600 seconds
        max_age = 3600
        deleted_count = 0
        checked_count = 0
        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                # Check both files and directories (like temp dirs for images)
                # Use lstat to avoid following symlinks if any
                stat_result = os.lstat(path)
                is_file = os.path.isfile(path)
                is_dir = os.path.isdir(path)

                if is_file or is_dir:
                    checked_count += 1
                    file_age = now - stat_result.st_mtime
                    if file_age > max_age:
                        logger.info(f"Teardown: Xóa {'file' if is_file else 'dir'} cũ ({file_age:.0f}s > {max_age}s): {path}")
                        if safe_remove(path): # Use safe_remove for robustness
                             deleted_count += 1
            except FileNotFoundError:
                # File might have been deleted by another process/request between listdir and stat
                continue
            except Exception as e:
                logger.error(f"Lỗi khi kiểm tra/dọn dẹp file/dir {path}: {e}")
        logger.info(f"Teardown dọn dẹp hoàn tất. Đã kiểm tra {checked_count} mục, xóa {deleted_count} mục cũ.")
    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng trong quá trình dọn dẹp teardown_appcontext: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    # Make sure UPLOAD_FOLDER exists at startup
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    logger.info(f"Thư mục Upload: {UPLOAD_FOLDER}")
    logger.info(f"Khởi động server trên cổng {port} - Chế độ Debug: {debug_mode}")
    # Use threaded=True for better handling of concurrent requests if needed,
    # but be mindful of resource usage (CPU for conversions)
    app.run(host='0.0.0.0', port=port, debug=debug_mode, threaded=True) # Added threaded=True

# --- END OF FILE app.py ---