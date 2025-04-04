from flask import Flask, request, send_file, render_template, jsonify, url_for, make_response
from flask_talisman import Talisman # Security Headers
from flask_wtf.csrf import CSRFProtect, CSRFError # CSRF Protection
from flask_limiter import Limiter # Rate Limiting
from flask_limiter.util import get_remote_address # Rate Limiting Helper
import os
import sys
import time
import subprocess
import logging
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge # Better handling for large files
from pdf2docx import Converter
import tempfile
import PyPDF2
import shutil # Cần cho safe_remove
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError, PDFSyntaxError
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image, UnidentifiedImageError
import zipfile
import magic # MIME Type Detection
from werkzeug.middleware.proxy_fix import ProxyFix

# === Basic Flask App Setup ===
app = Flask(__name__, template_folder='templates', static_folder='static')

# === Configuration ===
app.config['MAX_CONTENT_LENGTH'] = 101 * 1024 * 1024  # ~101MB limit server-side
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['WTF_CSRF_SECRET_KEY'] = app.config['SECRET_KEY']
# --- THAY ĐỔI: Có thể set thời gian CSRF nếu muốn, ví dụ 30 phút ---
# app.config['WTF_CSRF_TIME_LIMIT'] = 1800 # Mặc định là 3600 (1 giờ)
# --- KẾT THÚC THAY ĐỔI ---
app.wsgi_app = ProxyFix(
    app.wsgi_app, x_for=1, x_proto=1, x_host=1 # Các tham số quan trọng
)

# === Logging ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# === Security ===
csrf = CSRFProtect(app)
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["300 per day", "200 per hour", "50 per minute"], # Đã tăng giới hạn
    storage_uri="memory://",
    strategy="fixed-window"
)
csp = {
    'default-src': ['\'self\'', 'https://cdn.tailwindcss.com', 'https://fonts.googleapis.com', 'https://fonts.gstatic.com'],
    'style-src': ['\'self\'', '\'unsafe-inline\'', 'https://cdn.tailwindcss.com', 'https://fonts.googleapis.com'],
    'script-src': ['\'self\'', '\'unsafe-inline\'', 'https://cdn.tailwindcss.com'], # Bỏ 'unsafe-eval' nếu không cần
    'font-src': ['\'self\'', 'https://fonts.gstatic.com'],
    'img-src': ['\'self\'', 'data:'],
    'form-action': '\'self\''
}
talisman = Talisman(
    app,
    content_security_policy=csp,
    force_https=False, # Đặt thành True nếu có HTTPS
    session_cookie_secure=True, # Đặt thành True nếu có HTTPS
    session_cookie_http_only=True,
    frame_options='DENY',
    strict_transport_security=True, # Đặt thành False nếu không dùng HTTPS
)

# === Constants ===
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}
ALLOWED_IMAGE_EXTENSIONS = {'pdf', 'jpg', 'jpeg'}
ALLOWED_MIME_TYPES = {
    'pdf': ['application/pdf'],
    'docx': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword'],
    'ppt': ['application/vnd.ms-powerpoint'],
    'pptx': ['application/vnd.openxmlformats-officedocument.presentationml.presentation'],
    'jpg': ['image/jpeg'],
    'jpeg': ['image/jpeg']
}
LIBREOFFICE_TIMEOUT = 180 # Seconds
MIME_BUFFER_SIZE = 4096 # Read first 4KB for MIME detection

# === Helper Functions ===
def make_error_response(error_key, status_code=400):
    """Creates a Flask response with an error message prefixed for JS handling."""
    logger.warning(f"Returning error: {error_key} (Status: {status_code})")
    response_text = f"Conversion failed: {error_key}" # Giữ nguyên prefix để JS cũ vẫn có thể xử lý cơ bản
    response = make_response(response_text, status_code)
    response.headers["Content-Type"] = "text/plain; charset=utf-8"
    return response

# --- Logic tìm và xác minh LibreOffice (Giữ nguyên như trước) ---
_VERIFIED_SOFFICE_PATH = '/usr/lib/libreoffice/program/soffice'
SOFFICE_PATH = None
if os.path.isfile(_VERIFIED_SOFFICE_PATH):
    try:
        result = subprocess.run(
            [_VERIFIED_SOFFICE_PATH, '--headless', '--version'],
            capture_output=True, text=True, check=False, timeout=15
        )
        if result.returncode == 0 and 'LibreOffice' in result.stdout:
            logger.info(f"Using verified hardcoded LO path: {_VERIFIED_SOFFICE_PATH}")
            SOFFICE_PATH = _VERIFIED_SOFFICE_PATH
        else:
            logger.error(f"Hardcoded LO path {_VERIFIED_SOFFICE_PATH} exists, but version check failed! Code: {result.returncode}, Output: {result.stdout.strip()}")
    except subprocess.TimeoutExpired:
        logger.error(f"Timeout expired while verifying hardcoded LO path: {_VERIFIED_SOFFICE_PATH}")
    except Exception as e:
        logger.error(f"Error verifying hardcoded LO path {_VERIFIED_SOFFICE_PATH}: {e}")
else:
     logger.error(f"Hardcoded LO path {_VERIFIED_SOFFICE_PATH} does not exist or is not a file.")

if SOFFICE_PATH:
    logger.info(f"Successfully set LO path for use: {SOFFICE_PATH}")
else:
    logger.critical("LibreOffice could NOT be set/verified. Conversions requiring it WILL FAIL.")
# --- Kết thúc logic LibreOffice ---


def _allowed_file_extension(filename, allowed_set):
    """Checks only the file extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def safe_remove(item_path, retries=3, delay=0.5):
    # (No changes needed in this function's logic)
    if not item_path or not os.path.exists(item_path): return True
    is_dir = os.path.isdir(item_path); item_type = "directory" if is_dir else "file"
    for i in range(retries):
        try:
            if is_dir: shutil.rmtree(item_path)
            else: os.remove(item_path)
            logger.debug(f"Removed {item_type}: {item_path}"); return True
        except Exception as e: logger.warning(f"Error removing {item_path} (Attempt {i+1}): {e}"); time.sleep(delay*(i+1))
    logger.error(f"Failed to remove {item_type} after {retries} attempts: {item_path}"); return False

def get_actual_mime_type(file_storage):
    """Reads the beginning of a file stream to determine its MIME type."""
    try:
        original_pos = file_storage.stream.tell()
        file_storage.stream.seek(0)
        buffer = file_storage.stream.read(MIME_BUFFER_SIZE)
        file_storage.stream.seek(original_pos) # Reset stream position
        mime_type = magic.from_buffer(buffer, mime=True)
        logger.debug(f"Detected MIME type: {mime_type} for file {file_storage.filename}")
        return mime_type
    except magic.MagicException as e:
        logger.warning(f"Could not determine MIME type for {file_storage.filename}: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error during MIME detection for {file_storage.filename}: {e}")
        return None

# --- Other Helper Functions (get_pdf_page_size, setup_slide_size, etc.) ---
# (Các hàm helper khác giữ nguyên như code gốc bạn đã cung cấp)
def get_pdf_page_size(pdf_path):
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f);
            if reader.is_encrypted:
                try:
                    if reader.decrypt('') != PyPDF2.PasswordType.OWNER_PASSWORD:
                         logger.warning(f"Could not decrypt PDF {pdf_path} with empty password.")
                         raise ValueError("err-pdf-protected")
                except Exception:
                     logger.warning(f"Error during decryption attempt for {pdf_path}.")
                     raise ValueError("err-pdf-protected")
            if not reader.pages: return None, None
            page = reader.pages[0]; box = page.mediabox or page.cropbox
            if box: width = float(box.width); height = float(box.height); return width, height
    except PyPDF2.errors.PdfReadError as pdf_err: raise ValueError("err-pdf-corrupt") from pdf_err
    except ValueError as ve: raise ve
    except Exception as e: logger.error(f"Error reading PDF size {pdf_path}: {e}"); return None, None
    return None, None

def setup_slide_size(prs, pdf_path):
    pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)
    if pdf_width_pt is None: prs.slide_width, prs.slide_height = Inches(10), Inches(7.5); return prs
    try:
        pdf_width_in, pdf_height_in = pdf_width_pt / 72.0, pdf_height_pt / 72.0; max_dim = 56.0
        if pdf_width_in > max_dim or pdf_height_in > max_dim:
            ratio = pdf_width_in / pdf_height_in
            if pdf_width_in >= pdf_height_in: final_width, final_height = max_dim, max_dim / ratio
            else: final_height, final_width = max_dim, max_dim * ratio
        else: final_width, final_height = pdf_width_in, pdf_height_in
        prs.slide_width, prs.slide_height = Inches(final_width), Inches(final_height); return prs
    except Exception: prs.slide_width, prs.slide_height = Inches(10), Inches(7.5); return prs

def sort_key_for_pptx_images(filename):
    try: return int(os.path.splitext(filename)[0].split('-')[-1].split('_')[-1])
    except: return 0

def _convert_pdf_to_pptx_images(input_path, output_path):
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_")
        page_count_info = pdfinfo_from_path(input_path, poppler_path=None)
        page_count = page_count_info.get('Pages')
        if page_count is None: raise PDFInfoNotInstalledError("Poppler may be missing or invalid.")
        if page_count_info.get('Encrypted', 'no') == 'yes': raise ValueError("err-pdf-protected")

        if page_count == 0: Presentation().save(output_path); return True
        images = convert_from_path(input_path, dpi=300, fmt='jpeg', output_folder=temp_dir, thread_count=4, poppler_path=None)
        if not images: raise RuntimeError("err-conversion-img")
        prs = Presentation(); prs = setup_slide_size(prs, input_path); blank_layout = prs.slide_layouts[6]
        gen_imgs = sorted([f for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))], key=sort_key_for_pptx_images)
        if not gen_imgs: raise RuntimeError("err-conversion-img")
        slide_w, slide_h = prs.slide_width, prs.slide_height
        for img_fn in gen_imgs:
            img_path = os.path.join(temp_dir, img_fn)
            try:
                slide = prs.slides.add_slide(blank_layout)
                with Image.open(img_path) as img: img_w, img_h = img.size
                img_width_emu = int(img_w / 72 * 914400)
                img_height_emu = int(img_h / 72 * 914400)
                r_w = slide_w / img_width_emu if img_width_emu > 0 else 1
                r_h = slide_h / img_height_emu if img_height_emu > 0 else 1
                s = min(r_w, r_h); pic_w, pic_h = int(img_width_emu * s), int(img_height_emu * s)
                pic_l, pic_t = int((slide_w - pic_w) / 2), int((slide_h - pic_h) / 2)
                if pic_w > 0 and pic_h > 0: slide.shapes.add_picture(img_path, pic_l, pic_t, width=pic_w, height=pic_h)
            except UnidentifiedImageError:
                 logger.warning(f"Skipping invalid image file during PPTX creation: {img_fn}")
                 continue
            except Exception as page_err: logger.warning(f"Error adding image {img_fn} to PPTX: {page_err}")
        prs.save(output_path); return True
    except (PDFInfoNotInstalledError) as e: logger.error(f"PDF->PPTX Poppler Error: {e}"); raise ValueError("err-poppler-missing") from e
    except (PDFPageCountError, PDFSyntaxError) as e: logger.error(f"PDF->PPTX PDF Error: {e}"); raise ValueError("err-pdf-corrupt") from e
    except ValueError as ve: logger.error(f"PDF->PPTX Value Error: {ve}"); raise ve
    except RuntimeError as rte: logger.error(f"PDF->PPTX Runtime Error: {rte}"); raise rte
    except Exception as e: logger.error(f"Unexpected PDF->PPTX Error: {e}", exc_info=True); raise RuntimeError("err-unknown") from e
    finally: safe_remove(temp_dir)

def convert_pdf_to_pptx_python(input_path, output_path):
    logger.info("Attempting PDF -> PPTX via Python (image-based)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)

def convert_images_to_pdf(image_files, output_path):
    # (Giữ nguyên code gốc của hàm này)
    image_objects = []
    try:
        allowed_mimes = ALLOWED_MIME_TYPES['jpeg']
        for file_storage in image_files:
             mime_type = get_actual_mime_type(file_storage)
             if not mime_type or mime_type not in allowed_mimes:
                 logger.warning(f"Invalid MIME type detected for {secure_filename(file_storage.filename)}: {mime_type}. Allowed: {allowed_mimes}")
                 raise ValueError("err-invalid-mime-type-image")

        sorted_files = sorted(image_files, key=lambda f: secure_filename(f.filename))

        for file_storage in sorted_files:
            filename = secure_filename(file_storage.filename)
            try:
                file_storage.stream.seek(0)
                img_io = BytesIO(file_storage.stream.read())
                with Image.open(img_io) as img:
                    img.load(); converted_img = None
                    if img.mode in ['RGBA', 'LA']:
                         bg = Image.new('RGB', img.size, (255, 255, 255)); mask=None
                         try: mask = img.getchannel('A' if img.mode == 'RGBA' else 'L')
                         except: pass
                         bg.paste(img, mask=mask); converted_img = bg
                    elif img.mode not in ['RGB', 'L', 'CMYK']: converted_img = img.convert('RGB')
                    else: converted_img = img.copy()
                    image_objects.append(converted_img)
            except UnidentifiedImageError:
                 logger.error(f"File {filename} is not a valid image format recognized by Pillow.")
                 raise ValueError("err-invalid-image-file")
            except Exception as img_err:
                 logger.error(f"Error processing image {filename}: {img_err}")
                 raise RuntimeError("err-conversion") from img_err

        if not image_objects: raise ValueError("err-select-file")

        image_objects[0].save(output_path, "PDF", resolution=100.0, save_all=True, append_images=image_objects[1:])
        return True
    except ValueError as ve: raise ve
    except Exception as e:
        logger.error(f"Unexpected error converting images to PDF: {e}", exc_info=True)
        raise RuntimeError("err-unknown") from e
    finally:
        for img_obj in image_objects:
             try: img_obj.close()
             except: pass

def convert_pdf_to_image_zip(input_path, output_zip_path, img_format='jpeg'):
    # (Giữ nguyên code gốc của hàm này)
    temp_dir = None; fmt = img_format.lower(); ext = 'jpg' if fmt in ['jpeg', 'jpg'] else fmt
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdf2imgzip_")
        try:
             page_count_info = pdfinfo_from_path(input_path, poppler_path=None)
             page_count = page_count_info.get('Pages')
             if page_count is None: raise PDFInfoNotInstalledError("Poppler may be missing.")
             if page_count_info.get('Encrypted', 'no') == 'yes': raise ValueError("err-pdf-protected")
             logger.info(f"PDF Info: {page_count} pages found.")
        except (PDFInfoNotInstalledError, FileNotFoundError) as e: raise ValueError("err-poppler-missing") from e
        except (PDFPageCountError, PDFSyntaxError) as e: raise ValueError("err-pdf-corrupt") from e
        except ValueError as ve: raise ve
        except Exception as info_err: logger.error(f"pdfinfo error: {info_err}"); raise ValueError("err-poppler-check-failed") from info_err

        if page_count == 0:
            logger.warning("PDF reported 0 pages. Creating empty ZIP.")
            with zipfile.ZipFile(output_zip_path, 'w') as zf: pass
            return True

        safe_base = secure_filename(f"page_{os.path.splitext(os.path.basename(input_path))[0]}")
        images = convert_from_path(input_path, dpi=200, fmt=fmt, output_folder=temp_dir, output_file=safe_base, thread_count=4, poppler_path=None)
        if not images:
             if page_count > 0: raise RuntimeError("err-conversion-img")
             else: return True # No images generated, but PDF had 0 pages initially

        def sort_key(f):
             try: return int(os.path.splitext(f)[0].split('-')[-1])
             except: return 0
        gen_files = sorted([f for f in os.listdir(temp_dir) if f.lower().startswith(safe_base.lower()) and f.lower().endswith(f'.{ext}')], key=sort_key)
        if not gen_files and page_count > 0: # Check page_count again
            raise RuntimeError("err-conversion-img")
        elif not gen_files and page_count == 0:
             logger.warning("PDF had 0 pages and no images were generated. Creating empty ZIP.")
             with zipfile.ZipFile(output_zip_path, 'w') as zf: pass
             return True

        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
             for i, filename in enumerate(gen_files): zf.write(os.path.join(temp_dir, filename), f"page_{i+1}.{ext}")
        return True
    except (PDFInfoNotInstalledError) as e: raise ValueError("err-poppler-missing") from e
    except (PDFPageCountError, PDFSyntaxError) as e: raise ValueError("err-pdf-corrupt") from e
    except ValueError as ve: raise ve
    except RuntimeError as rte: raise rte
    except Exception as e: logger.error(f"Unexpected PDF->ZIP Error: {e}", exc_info=True); raise RuntimeError("err-unknown") from e
    finally: safe_remove(temp_dir)


# === Global Error Handlers ===
# (Giữ nguyên các error handlers: CSRFError, RequestEntityTooLarge, 429, Exception)
@app.errorhandler(CSRFError)
def handle_csrf_error(e):
    logger.warning(f"CSRF validation failed: {e.description}")
    return make_error_response("err-csrf-invalid", 400)

@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    logger.warning(f"File upload rejected (too large): {e.description}")
    return make_error_response("err-file-too-large", 413)

@app.errorhandler(429) # Rate limit exceeded
def ratelimit_handler(e):
    logger.warning(f"Rate limit exceeded: {e.description}")
    return make_error_response("err-rate-limit-exceeded", 429)

@app.errorhandler(Exception) # Generic fallback for unhandled exceptions
def handle_generic_exception(e):
     from werkzeug.exceptions import HTTPException
     if isinstance(e, HTTPException):
          # Let Flask/Werkzeug handle standard HTTP exceptions (like 404, 405)
          return e
     # Log all other unhandled exceptions
     logger.error(f"Unhandled Exception: {e}", exc_info=True)
     # Return a generic server error
     return make_error_response("err-unknown", 500)


# === Routes ===

# === Translations Route ===
@app.route('/api/translations')
def get_translations():
    """Provides translation strings to the frontend (Hardcoded Version)."""
    translations = {
        'en': {
            'lang-title': 'PDF & Office Tools', 'lang-subtitle': 'Simple, powerful tools for your documents',
            'lang-error-title': 'Error!', 'lang-convert-title': 'Convert PDF/Office',
            'lang-convert-desc': 'Transform PDF to Word/PPT and vice versa', 'lang-compress-title': 'Compress PDF',
            'lang-compress-desc': 'Reduce file size while maintaining quality', 'lang-merge-title': 'Merge PDF',
            'lang-merge-desc': 'Combine multiple PDFs into one file', 'lang-split-title': 'Split PDF',
            'lang-split-desc': 'Extract pages from your PDF', 'lang-rotate-title': 'Rotate PDF',
            'lang-rotate-desc': 'Change page orientation', 'lang-image-title': 'PDF ↔ Image',
            'lang-image-desc': 'Convert PDF to images or images to PDF', 'lang-image-input-label': 'Select PDF or Image(s) (JPG/JPEG only)',
            'lang-image-convert-btn': 'Convert Now', 'lang-image-converting': 'Converting...',
            'lang-size-limit': 'Size limit: 100MB (total)', 'lang-select-conversion': 'Select conversion type',
            'lang-converting': 'Converting...', 'lang-convert-btn': 'Convert Now',
            'lang-file-input-label': 'Select file', 'file-no-selected': 'No file selected',
            'err-select-file': 'Please select file(s) to convert.', 'err-file-too-large': 'Total file size exceeds the limit (100MB).',
            'err-select-conversion': 'Please select a conversion type.', 'err-format-docx': 'Select one PDF or DOCX file for this conversion.',
            'err-format-ppt': 'Select one PDF, PPT or PPTX file for this conversion.', 'err-conversion': 'An error occurred during conversion.',
            'err-fetch-translations': 'Could not load language data.', 'lang-select-btn-text': 'Browse',
            'lang-select-conversion-label': 'Conversion Type', 'err-multi-file-not-supported': 'Multi-file selection is only supported for Image to PDF conversion.',
            'err-invalid-image-file': 'One or more selected files are not valid images (Pillow error).', 'err-image-format': 'Invalid file type. Select PDF, JPG, or JPEG based on conversion.',
            'err-image-single-pdf': 'Please select only one PDF file to convert to images.', 'err-image-all-images': 'If selecting multiple files, all must be JPG or JPEG to convert to PDF.',
            'err-libreoffice': 'Conversion failed (Processing engine error).', 'err-conversion-timeout': 'Conversion timed out.',
            'err-poppler-missing': 'PDF processing library (Poppler) missing or failed.', 'err-pdf-corrupt': 'Could not process PDF (corrupt file?).',
            'err-unknown': 'An unexpected error occurred. Please try again later.',
            'err-csrf-invalid': 'Security validation failed. Please refresh the page and try again.', # Cập nhật thông báo CSRF
            'err-rate-limit-exceeded': 'Too many requests. Please wait a moment and try again.',
            'err-invalid-mime-type': 'Invalid file type detected. The file content does not match the expected format.',
            # --- THÊM DỊCH CHO LỖI MỚI ---
            'err-mime-unidentified-office': "Could not identify file type, it might be non-standard. Please 'Save As...' in the original application and upload the new file.",
            # --- KẾT THÚC THÊM DỊCH ---
            'err-invalid-mime-type-image': 'Invalid image type detected. Only JPEG files are allowed for Image-to-PDF.',
            'err-pdf-protected': 'Cannot process password-protected PDF.',
            'err-poppler-check-failed': 'Failed to get PDF info (Poppler check).',
            'err-conversion-img': 'Failed to convert/extract images from PDF.',
            'lang-clear-all': 'Clear All'
            # Thêm các key ngôn ngữ khác cho JS nếu cần (vd: upload-a-file)
             ,'lang-upload-a-file': 'Upload files'
             ,'lang-drag-drop': 'or drag and drop'
             ,'lang-image-types': 'PDF, JPG, JPEG up to 100MB total'
        },
        'vi': {
            'lang-title': 'Công Cụ PDF & Office', 'lang-subtitle': 'Công cụ đơn giản, mạnh mẽ cho tài liệu của bạn',
            'lang-error-title': 'Lỗi!', 'lang-convert-title': 'Chuyển đổi PDF/Office',
            'lang-convert-desc': 'Chuyển đổi PDF sang Word/PPT và ngược lại', 'lang-compress-title': 'Nén PDF',
            'lang-compress-desc': 'Giảm kích thước tệp trong khi duy trì chất lượng', 'lang-merge-title': 'Gộp PDF',
            'lang-merge-desc': 'Kết hợp nhiều tệp PDF thành một tệp', 'lang-split-title': 'Tách PDF',
            'lang-split-desc': 'Trích xuất các trang từ tệp PDF của bạn', 'lang-rotate-title': 'Xoay PDF',
            'lang-rotate-desc': 'Thay đổi hướng trang', 'lang-image-title': 'PDF ↔ Ảnh',
            'lang-image-desc': 'Chuyển PDF thành ảnh hoặc ảnh thành PDF', 'lang-image-input-label': 'Chọn PDF hoặc (các) Ảnh (chỉ JPG/JPEG)',
            'lang-image-convert-btn': 'Chuyển đổi ngay', 'lang-image-converting': 'Đang chuyển đổi...',
            'lang-size-limit': 'Giới hạn kích thước: 100MB (tổng)', 'lang-select-conversion': 'Chọn kiểu chuyển đổi',
            'lang-converting': 'Đang chuyển đổi...', 'lang-convert-btn': 'Chuyển đổi ngay',
            'lang-file-input-label': 'Chọn tệp', 'file-no-selected': 'Không có tệp nào được chọn',
            'err-select-file': 'Vui lòng chọn (các) tệp để chuyển đổi.', 'err-file-too-large': 'Tổng kích thước tệp vượt quá giới hạn (100MB).',
            'err-select-conversion': 'Vui lòng chọn kiểu chuyển đổi.', 'err-format-docx': 'Chọn một file PDF hoặc DOCX cho chuyển đổi này.',
            'err-format-ppt': 'Chọn một file PDF, PPT hoặc PPTX cho chuyển đổi này.', 'err-conversion': 'Đã xảy ra lỗi trong quá trình chuyển đổi.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.', 'lang-select-btn-text': 'Duyệt...',
            'lang-select-conversion-label': 'Kiểu chuyển đổi', 'err-multi-file-not-supported': 'Chỉ hỗ trợ chọn nhiều file khi chuyển đổi Ảnh sang PDF.',
            'err-invalid-image-file': 'Một hoặc nhiều tệp được chọn không phải là ảnh hợp lệ (lỗi Pillow).', 'err-image-format': 'Loại tệp không hợp lệ. Chọn PDF, JPG, hoặc JPEG tùy theo chuyển đổi.',
            'err-image-single-pdf': 'Vui lòng chỉ chọn một file PDF để chuyển đổi sang ảnh.', 'err-image-all-images': 'Nếu chọn nhiều tệp, tất cả phải là JPG hoặc JPEG để chuyển đổi sang PDF.',
            'err-libreoffice': 'Chuyển đổi thất bại (Lỗi bộ xử lý).', 'err-conversion-timeout': 'Quá trình chuyển đổi quá thời gian.',
            'err-poppler-missing': 'Thiếu hoặc lỗi thư viện xử lý PDF (Poppler).', 'err-pdf-corrupt': 'Không thể xử lý PDF (tệp lỗi?).',
            'err-unknown': 'Đã xảy ra lỗi không mong muốn. Vui lòng thử lại sau.',
            'err-csrf-invalid': 'Xác thực bảo mật thất bại. Vui lòng tải lại trang và thử lại.', # Cập nhật thông báo CSRF
            'err-rate-limit-exceeded': 'Quá nhiều yêu cầu. Vui lòng đợi một lát và thử lại.',
            'err-invalid-mime-type': 'Phát hiện loại tệp không hợp lệ. Nội dung tệp không khớp định dạng mong đợi.',
            # --- THÊM DỊCH CHO LỖI MỚI ---
            'err-mime-unidentified-office': "Không thể nhận dạng loại file dù có đuôi Office. Vui lòng dùng chức năng 'Lưu thành...' (Save As...) trong ứng dụng gốc để lưu bản mới và tải lên lại.",
            # --- KẾT THÚC THÊM DỊCH ---
            'err-invalid-mime-type-image': 'Phát hiện loại ảnh không hợp lệ. Chỉ cho phép tệp JPEG để chuyển đổi Ảnh sang PDF.',
            'err-pdf-protected': 'Không thể xử lý PDF được bảo vệ bằng mật khẩu.',
            'err-poppler-check-failed': 'Không thể lấy thông tin PDF (lỗi kiểm tra Poppler).',
            'err-conversion-img': 'Không thể chuyển đổi/trích xuất ảnh từ PDF.',
            'lang-clear-all': 'Xóa tất cả'
            # Thêm các key ngôn ngữ khác cho JS nếu cần
             ,'lang-upload-a-file': 'Tải tệp lên'
             ,'lang-drag-drop': 'hoặc kéo và thả'
             ,'lang-image-types': 'PDF, JPG, JPEG tối đa 100MB tổng'
        }
    }
    lang = request.args.get('lang', 'en')
    return jsonify(translations.get(lang, translations.get('en', {})))


@app.route('/')
def index():
    """Renders the main page."""
    # (Giữ nguyên route /)
    try:
        translations_url = url_for('get_translations', _external=False)
        return render_template('index.html', translations_url=translations_url)
    except Exception as e:
        logger.error(f"Error rendering index page: {e}", exc_info=True)
        # Use make_error_response for consistency, but maybe render a simple error page instead?
        return make_error_response("err-unknown", 500) # Or render_template('error.html', ...)


# === PDF / Office Conversion Route ===
@app.route('/convert', methods=['POST'])
@limiter.limit("10 per minute") # Giới hạn chung cho route này
def convert_file():
    """Handles PDF <-> DOCX and PDF <-> PPT conversions with security checks."""
    output_path = temp_libreoffice_output = input_path_for_process = None
    saved_input_paths = []; actual_conversion_type = None; start_time = time.time()
    error_key = "err-conversion"; conversion_success = False

    try:
        # Validate file presence
        if 'file' not in request.files: return make_error_response("err-select-file", 400)
        file = request.files['file']
        if not file or not file.filename: return make_error_response("err-select-file", 400)

        # Validate filename and basic extension
        filename = secure_filename(file.filename)
        file_ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
        allowed_office_ext = {'pdf', 'docx', 'ppt', 'pptx'}
        if not _allowed_file_extension(filename, allowed_office_ext):
             logger.warning(f"Rejected file by extension: {filename}")
             # Decide which error message is more appropriate based on context
             # Using a generic one might be safer if the flow is complex
             return make_error_response("err-format-docx", 400) # Or a more generic format error

        # Validate conversion type selection
        actual_conversion_type = request.form.get('conversion_type')
        valid_conversion_types = ['pdf_to_docx', 'docx_to_pdf', 'pdf_to_ppt', 'ppt_to_pdf']
        if not actual_conversion_type or actual_conversion_type not in valid_conversion_types:
             return make_error_response("err-select-conversion", 400)

        # Validate file extension against conversion type
        required_ext = []
        if actual_conversion_type == 'pdf_to_docx': required_ext = ['pdf']
        elif actual_conversion_type == 'docx_to_pdf': required_ext = ['docx']
        elif actual_conversion_type == 'pdf_to_ppt': required_ext = ['pdf']
        elif actual_conversion_type == 'ppt_to_pdf': required_ext = ['ppt', 'pptx']

        if file_ext not in required_ext:
             # Determine specific error message based on the expected format
             error_key_cv = "err-format-docx" if 'docx' in required_ext else "err-format-ppt"
             logger.warning(f"Extension mismatch: file '{filename}' ({file_ext}), required {required_ext} for type '{actual_conversion_type}'")
             return make_error_response(error_key_cv, 400)

        logger.info(f"Request /convert: file='{filename}', type='{actual_conversion_type}'")

        # === MIME Type Validation ===
        detected_mime = get_actual_mime_type(file)
        expected_mimes = []
        if actual_conversion_type in ['pdf_to_docx', 'pdf_to_ppt']:
             expected_mimes = ALLOWED_MIME_TYPES['pdf']
        elif actual_conversion_type == 'docx_to_pdf':
             expected_mimes = ALLOWED_MIME_TYPES['docx']
        elif actual_conversion_type == 'ppt_to_pdf':
             expected_mimes = ALLOWED_MIME_TYPES['ppt'] + ALLOWED_MIME_TYPES['pptx']

        # --- START: MODIFIED MIME CHECK LOGIC ---
        if not detected_mime or detected_mime not in expected_mimes:
            # Special check: If detected as generic stream but extension suggests Office, give specific error
            is_expected_office_ext = file_ext in ['ppt', 'pptx', 'docx'] # Check against relevant extensions
            # Check if conversion type *expects* one of these Office formats as input
            is_office_input_conversion = actual_conversion_type in ['ppt_to_pdf', 'docx_to_pdf']

            if detected_mime == 'application/octet-stream' and is_expected_office_ext and is_office_input_conversion:
                 logger.warning(f"Unidentified MIME type for expected Office file {filename}. Detected: {detected_mime}, Ext: {file_ext}, ConvType: {actual_conversion_type}")
                 # Return the new specific error code
                 return make_error_response("err-mime-unidentified-office", 400)
            else:
                 # Otherwise, it's a general MIME type mismatch
                 logger.warning(f"MIME type validation failed for {filename}. Detected: '{detected_mime}', Expected one of: {expected_mimes}")
                 return make_error_response("err-invalid-mime-type", 400)
        # --- END: MODIFIED MIME CHECK LOGIC ---
        logger.info(f"MIME type validated for {filename}: {detected_mime}")
        # === End MIME Type Validation ===


        # Create upload folder if it doesn't exist
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Save the uploaded file securely
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        input_filename_ts = f"input_{timestamp}_{filename}" # Use the already secured filename
        input_path_for_process = os.path.join(UPLOAD_FOLDER, input_filename_ts)
        try:
            file.save(input_path_for_process)
            saved_input_paths.append(input_path_for_process)
            logger.info(f"Input saved: {input_path_for_process}")
        except Exception as save_err:
            logger.error(f"File save failed for {filename}: {save_err}")
            # Don't need cleanup here as nothing else was created yet
            return make_error_response("err-unknown", 500)

        # Determine output filename and path
        base_name = filename.rsplit('.', 1)[0]
        out_ext_map = {'pdf_to_docx': 'docx', 'docx_to_pdf': 'pdf', 'pdf_to_ppt': 'pptx', 'ppt_to_pdf': 'pdf'}
        out_ext = out_ext_map.get(actual_conversion_type)
        # Ensure output filename is also secure
        output_filename = f"converted_{timestamp}_{secure_filename(base_name)}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)


        # --- Perform Conversion ---
        try:
            if actual_conversion_type == 'pdf_to_docx':
                cv = None
                try:
                    logger.info(f"Starting pdf2docx conversion for {input_path_for_process}")
                    cv = Converter(input_path_for_process)
                    cv.convert(output_path) # Default parameters
                    conversion_success = True
                    logger.info(f"pdf2docx conversion successful: {output_path}")
                except Exception as pdf2docx_err:
                    logger.error(f"pdf2docx conversion failed for {input_path_for_process}: {pdf2docx_err}", exc_info=True)
                    error_key = "err-conversion" # Keep generic or make more specific if possible
                finally:
                    if cv: cv.close() # Ensure resources are released

            elif actual_conversion_type in ['docx_to_pdf', 'ppt_to_pdf']:
                if not SOFFICE_PATH:
                    logger.error("LibreOffice path (SOFFICE_PATH) is not set or verified. Cannot perform conversion.")
                    raise RuntimeError("err-libreoffice") # Use specific error key

                output_dir = os.path.dirname(output_path)
                # Determine the expected output filename LibreOffice *might* create
                input_file_ext_actual = os.path.splitext(input_path_for_process)[1].lower()
                # Use the *input* basename for the expected LO output, as LO does this
                expected_lo_output_name = os.path.basename(input_path_for_process).replace(input_file_ext_actual, '.pdf')
                temp_libreoffice_output = os.path.join(output_dir, expected_lo_output_name)
                # Clean up any potential leftover file from previous failed runs
                safe_remove(temp_libreoffice_output)

                cmd = [SOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path_for_process]
                logger.info(f"Running LibreOffice command: {' '.join(cmd)}")
                try:
                    # Increased timeout slightly, added encoding handling
                    result = subprocess.run(cmd, check=True, timeout=LIBREOFFICE_TIMEOUT, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                    logger.info(f"LibreOffice stdout:\n{result.stdout}")
                    logger.info(f"LibreOffice stderr:\n{result.stderr}")

                    # Verify output file exists and has content, then rename
                    if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0:
                        os.rename(temp_libreoffice_output, output_path) # Rename to our desired output name
                        conversion_success = True
                        logger.info(f"LibreOffice conversion successful: {output_path}")
                    else:
                        logger.error(f"LibreOffice conversion ran for {input_path_for_process} but output file '{temp_libreoffice_output}' is missing or empty.")
                        error_key = "err-libreoffice" # Specific LO error
                except subprocess.TimeoutExpired:
                    logger.error(f"LibreOffice conversion timed out after {LIBREOFFICE_TIMEOUT}s for {input_path_for_process}.")
                    error_key = "err-conversion-timeout"
                except subprocess.CalledProcessError as lo_err:
                    logger.error(f"LibreOffice conversion failed for {input_path_for_process}. Return code: {lo_err.returncode}")
                    logger.error(f"LibreOffice stderr:\n{lo_err.stderr}") # Log stderr for debugging
                    error_key = "err-libreoffice"
                except FileNotFoundError:
                    # This should ideally not happen if SOFFICE_PATH is verified at startup
                    logger.error(f"LibreOffice executable not found at runtime: {SOFFICE_PATH}")
                    error_key = "err-libreoffice"
                except Exception as lo_run_err:
                    logger.error(f"Unexpected error running LibreOffice for {input_path_for_process}: {lo_run_err}", exc_info=True)
                    error_key = "err-libreoffice" # Group under LO errors

            elif actual_conversion_type == 'pdf_to_ppt':
                # Try Python method first
                python_method_success = False; python_method_error_key = None
                try:
                    if convert_pdf_to_pptx_python(input_path_for_process, output_path):
                         python_method_success = True
                         conversion_success = True # Set overall success
                         error_key = None # Clear any potential previous error
                         logger.info("PDF->PPTX conversion successful using Python method.")
                except ValueError as ve:
                    # Capture specific error keys raised by the helper
                    python_method_error_key = str(ve) if str(ve).startswith("err-") else "err-conversion"
                    logger.warning(f"Python PDF->PPTX method failed with ValueError: {python_method_error_key}")
                except RuntimeError as rte:
                     python_method_error_key = str(rte) if str(rte).startswith("err-") else "err-conversion"
                     logger.warning(f"Python PDF->PPTX runtime error: {python_method_error_key}")
                except Exception as py_ppt_err:
                     python_method_error_key = "err-conversion" # Fallback error key
                     logger.error(f"Unexpected error in Python PDF->PPTX method: {py_ppt_err}", exc_info=True)

                # Fallback to LibreOffice if Python method failed AND LO is available
                # AND the error wasn't something LO likely can't fix (like corrupt/protected PDF)
                can_fallback = (
                    not python_method_success and
                    SOFFICE_PATH and
                    python_method_error_key not in ["err-pdf-corrupt", "err-pdf-protected", "err-poppler-missing"]
                )
                if can_fallback:
                    logger.info(f"Python PDF->PPTX failed ({python_method_error_key}), attempting LibreOffice fallback...")
                    error_key = "err-conversion" # Reset error key for fallback attempt
                    output_dir = os.path.dirname(output_path)
                    # Expected LO output name (will be .pptx)
                    input_file_ext_actual = os.path.splitext(input_path_for_process)[1].lower()
                    expected_lo_output_name = os.path.basename(input_path_for_process).replace(input_file_ext_actual, '.pptx')
                    temp_libreoffice_output = os.path.join(output_dir, expected_lo_output_name)
                    safe_remove(temp_libreoffice_output) # Clean potential leftovers

                    cmd = [SOFFICE_PATH, '--headless', '--convert-to', 'pptx', '--outdir', output_dir, input_path_for_process]
                    logger.info(f"Running LibreOffice command: {' '.join(cmd)}")
                    try:
                        result = subprocess.run(cmd, check=True, timeout=LIBREOFFICE_TIMEOUT, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                        logger.info(f"LibreOffice stdout:\n{result.stdout}")
                        logger.info(f"LibreOffice stderr:\n{result.stderr}")
                        if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0:
                            os.rename(temp_libreoffice_output, output_path)
                            conversion_success = True # Set overall success
                            error_key = None # Clear error on success
                            logger.info("LibreOffice fallback for PDF->PPTX successful.")
                        else:
                            logger.error("LibreOffice fallback ran but output file is missing or empty.")
                            error_key = "err-libreoffice"
                    except subprocess.TimeoutExpired:
                        logger.error("LibreOffice fallback conversion timed out.")
                        error_key = "err-conversion-timeout"
                    except subprocess.CalledProcessError as lo_err:
                        logger.error(f"LibreOffice fallback conversion failed. Return code: {lo_err.returncode}")
                        logger.error(f"LibreOffice stderr:\n{lo_err.stderr}")
                        error_key = "err-libreoffice"
                    except FileNotFoundError:
                        logger.error(f"LibreOffice executable not found at runtime (should not happen): {SOFFICE_PATH}")
                        error_key = "err-libreoffice"
                    except Exception as lo_run_err:
                        logger.error(f"Unexpected error running LibreOffice fallback: {lo_run_err}", exc_info=True)
                        error_key = "err-libreoffice"
                elif not python_method_success:
                     # If fallback wasn't possible or wasn't attempted, use the error from the Python method
                     error_key = python_method_error_key or "err-conversion"
                     logger.warning(f"Skipping or no LibreOffice fallback available. Final error from Python method: {error_key}")

        # --- Catch specific errors raised during conversion steps ---
        except RuntimeError as rt_err:
            error_key = str(rt_err) if str(rt_err).startswith("err-") else "err-unknown"
            logger.error(f"Caught RuntimeError during conversion: {error_key}", exc_info=False) # No need for full stack trace if it's our defined error
        except ValueError as val_err:
             error_key = str(val_err) if str(val_err).startswith("err-") else "err-unknown"
             logger.error(f"Caught ValueError during conversion: {error_key}", exc_info=False)
        except Exception as conv_err:
            # Catch any other unexpected error during the conversion block
            error_key = "err-unknown"
            logger.error(f"Unexpected error during conversion process: {conv_err}", exc_info=True)
        # --- End Conversion ---

        # --- Handle Success or Failure ---
        if conversion_success and output_path and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            mimetype_map = {'pdf': 'application/pdf', 'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'}
            mimetype = mimetype_map.get(out_ext, 'application/octet-stream')
            try:
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                # Use call_on_close for cleanup after response is sent
                @response.call_on_close
                def cleanup_success():
                    logger.debug(f"Cleaning up successful /convert: Input: {input_path_for_process}, Output: {output_path}")
                    safe_remove(input_path_for_process)
                    safe_remove(output_path)
                    # No need to remove temp_libreoffice_output here as it was renamed or should have failed
                logger.info(f"Conversion successful. Sending file: {output_filename}. Time: {time.time() - start_time:.2f}s")
                return response
            except Exception as send_err:
                # Error during file sending (rare)
                logger.error(f"Error sending file {output_filename}: {send_err}", exc_info=True)
                # Raise a generic error to be caught by the outer handler, triggering cleanup
                raise RuntimeError("err-unknown") from send_err
        else:
            # Conversion failed or produced empty/missing output
            final_error_key = error_key or "err-conversion" # Use specific error if available
            logger.error(f"Conversion failed or produced invalid output. Final Error Key: {final_error_key}. Time: {time.time() - start_time:.2f}s")
            # Raise the determined error key to be caught by the outer handler
            raise RuntimeError(final_error_key)

    # --- Outer Exception Handler (Catch all errors from the try block) ---
    except Exception as e:
         # Determine the final error key and status code
         final_error_key = str(e) if str(e).startswith("err-") else "err-unknown"
         status_code = 400 # Default to Bad Request

         # Map specific error keys to appropriate HTTP status codes
         if final_error_key == "err-unknown":
             status_code = 500
             # Log unexpected errors with full trace
             logger.error(f"Unexpected error in /convert handler: {e}", exc_info=True)
         elif final_error_key == "err-file-too-large": status_code = 413
         elif final_error_key == "err-rate-limit-exceeded": status_code = 429
         elif final_error_key == "err-csrf-invalid": status_code = 400
         # Add more mappings if needed (e.g., 403 for permission errors)

         # --- Cleanup for Failed Request ---
         logger.debug(f"Cleaning up failed /convert request (Error: {final_error_key}).")
         # Cleanup input files that were successfully saved
         for p in saved_input_paths: safe_remove(p)
         # Cleanup potential output files (even if incomplete/empty)
         safe_remove(output_path)
         # Cleanup temporary LibreOffice output if it still exists
         if temp_libreoffice_output and os.path.exists(temp_libreoffice_output):
             safe_remove(temp_libreoffice_output)
         # --- End Cleanup ---

         # Return the standardized error response
         return make_error_response(final_error_key, status_code)


# === PDF / Image Conversion Route ===
# (Giữ nguyên route /convert_image như code gốc của bạn)
@app.route('/convert_image', methods=['POST'])
@limiter.limit("10 per minute")
def convert_image_route():
    """Handles PDF <-> Image conversions with security checks."""
    output_path = None; input_path_for_pdf_input = None; saved_input_paths = []
    actual_conversion_type = None; output_filename = None; start_time = time.time()
    error_key = "err-conversion"; conversion_success = False
    temp_upload_dir = None # Directory for temporary image storage during img->pdf

    try:
        uploaded_files = request.files.getlist('image_file') # Use the correct name from JS FormData
        if not uploaded_files or not all(f and f.filename for f in uploaded_files):
            return make_error_response("err-select-file", 400)

        logger.info(f"Request /convert_image: Received {len(uploaded_files)} file(s).")

        # Determine conversion type based on the first file
        first_file = uploaded_files[0]
        first_filename = secure_filename(first_file.filename)
        first_ext = first_filename.rsplit('.', 1)[-1].lower() if '.' in first_filename else ''

        validation_error_key = None
        out_ext = None
        valid_files_for_processing = [] # Will store FileStorage for PDF->Img, paths for Img->PDF

        # --- Input Validation Logic ---
        if first_ext == 'pdf':
            if len(uploaded_files) > 1:
                validation_error_key = "err-image-single-pdf"
            elif not _allowed_file_extension(first_filename, ALLOWED_IMAGE_EXTENSIONS): # Check extension first
                validation_error_key = "err-image-format"
            else:
                # Validate MIME type for the single PDF file
                mime_type = get_actual_mime_type(first_file)
                if not mime_type or mime_type not in ALLOWED_MIME_TYPES['pdf']:
                     logger.warning(f"Invalid MIME type for PDF upload {first_filename}: {mime_type}")
                     validation_error_key = "err-invalid-mime-type"
                else:
                     actual_conversion_type = 'pdf_to_image'
                     out_ext = 'zip'
                     valid_files_for_processing.append(first_file) # Store the FileStorage object

        elif first_ext in ['jpg', 'jpeg']:
            actual_conversion_type = 'image_to_pdf'
            out_ext = 'pdf'
            allowed_image_mimes = ALLOWED_MIME_TYPES['jpeg'] # Only JPEG allowed for Img->PDF
            # Create a temporary directory to store uploaded images before processing
            # This avoids keeping multiple large files in memory
            try:
                temp_upload_dir = tempfile.mkdtemp(prefix="img2pdf_")
            except Exception as temp_err:
                 logger.error(f"Failed to create temporary directory for image upload: {temp_err}")
                 return make_error_response("err-unknown", 500)

            total_size = 0
            max_size_bytes = app.config['MAX_CONTENT_LENGTH']

            for i, f in enumerate(uploaded_files):
                fname_sec = secure_filename(f.filename)
                f_ext = fname_sec.rsplit('.', 1)[-1].lower() if '.' in fname_sec else ''

                # Check extension consistency (all must be jpg/jpeg)
                if f_ext not in ['jpg', 'jpeg']:
                    validation_error_key = "err-image-all-images"; logger.warning(f"Invalid extension for image upload {fname_sec}"); break

                # Check file size
                f.stream.seek(0, os.SEEK_END)
                file_size = f.stream.tell()
                f.stream.seek(0) # Reset stream position
                total_size += file_size
                if total_size > max_size_bytes:
                     validation_error_key = "err-file-too-large"; logger.warning(f"Total image size exceeded limit at file {fname_sec}"); break

                # Check MIME type
                mime_type = get_actual_mime_type(f)
                if not mime_type or mime_type not in allowed_image_mimes:
                    validation_error_key = "err-invalid-mime-type-image"; logger.warning(f"Invalid MIME type for image upload {fname_sec}: {mime_type}"); break

                # Save valid image to the temporary directory
                # Use index to prevent overwriting files with the same name
                temp_image_path = os.path.join(temp_upload_dir, f"{i}_{fname_sec}")
                try:
                    f.save(temp_image_path)
                    valid_files_for_processing.append(temp_image_path) # Store the path
                    # Keep track of saved paths for potential cleanup on error
                    saved_input_paths.append(temp_image_path)
                except Exception as save_err:
                    logger.error(f"Failed to save temporary image {fname_sec}: {save_err}"); validation_error_key = "err-unknown"; break # Stop processing on save error

            if validation_error_key: pass # Error already set
            elif not valid_files_for_processing: validation_error_key = "err-select-file" # No valid files processed

        else:
            # First file is neither PDF nor JPG/JPEG
            validation_error_key = "err-image-format"

        # Handle any validation errors found
        if validation_error_key:
            safe_remove(temp_upload_dir) # Clean up temp dir if created
            for p in saved_input_paths: safe_remove(p) # Clean up any saved temp images
            return make_error_response(validation_error_key, 400)
        # --- End Input Validation ---

        logger.info(f"Determined conversion type: {actual_conversion_type}. Validated {len(valid_files_for_processing)} file(s).")

        # Create main upload folder if needed
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")

        # Save the input PDF if it's PDF->Image (was stored as FileStorage)
        if actual_conversion_type == 'pdf_to_image':
            pdf_file_storage = valid_files_for_processing[0]
            input_filename_ts = f"input_{timestamp}_{secure_filename(pdf_file_storage.filename)}"
            input_path_for_pdf_input = os.path.join(UPLOAD_FOLDER, input_filename_ts)
            try:
                # Ensure stream is at the beginning before saving
                pdf_file_storage.stream.seek(0)
                pdf_file_storage.save(input_path_for_pdf_input)
                # Add the *saved* path to saved_input_paths for final cleanup
                saved_input_paths.append(input_path_for_pdf_input)
                logger.info(f"Input PDF saved: {input_path_for_pdf_input}")
            except Exception as save_err:
                logger.error(f"Failed to save PDF input {secure_filename(pdf_file_storage.filename)}: {save_err}")
                # No need to clean temp_upload_dir here as it wasn't used for PDF input
                return make_error_response("err-unknown", 500)

        # Determine output filename and path
        base_name = first_filename.rsplit('.', 1)[0]
        output_filename = f"converted_{timestamp}_{secure_filename(base_name)}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # --- Perform Conversion ---
        try:
            if actual_conversion_type == 'pdf_to_image':
                # Call the helper function for PDF -> ZIP
                if convert_pdf_to_image_zip(input_path_for_pdf_input, output_path):
                    conversion_success = True
            elif actual_conversion_type == 'image_to_pdf':
                # Use the list of saved image *paths* in the temp directory
                # --- Inlined/Adapted image_to_pdf logic using file paths ---
                image_objects_pil = []
                try:
                    # Sort paths alphabetically based on the temporary filename (which includes index)
                    sorted_paths = sorted(valid_files_for_processing)
                    for img_path in sorted_paths:
                         filename_log = os.path.basename(img_path) # For logging
                         try:
                             with Image.open(img_path) as img:
                                 img.load() # Load image data to process
                                 converted_img = None
                                 # Handle transparency or unsupported modes
                                 if img.mode in ['RGBA', 'LA']:
                                      # Create a white background and paste image onto it
                                      bg = Image.new('RGB', img.size, (255, 255, 255))
                                      mask = None
                                      try: mask = img.getchannel('A' if img.mode == 'RGBA' else 'L')
                                      except: pass # Ignore if mask channel is problematic
                                      bg.paste(img, mask=mask)
                                      converted_img = bg
                                 elif img.mode not in ['RGB', 'L', 'CMYK']:
                                      # Convert other modes (like P, 1) to RGB
                                      logger.debug(f"Converting image {filename_log} from mode {img.mode} to RGB")
                                      converted_img = img.convert('RGB')
                                 else:
                                      # Already in a supported mode (RGB, L, CMYK), just copy
                                      converted_img = img.copy()
                                 image_objects_pil.append(converted_img)
                         except UnidentifiedImageError:
                              logger.error(f"File {filename_log} is not a valid image format recognized by Pillow.")
                              raise ValueError("err-invalid-image-file") # Raise specific error
                         except Exception as img_err:
                              logger.error(f"Error processing image {filename_log}: {img_err}")
                              raise RuntimeError("err-conversion") from img_err # Raise generic conversion error

                    if not image_objects_pil:
                         # Should not happen if validation passed, but check anyway
                         raise ValueError("err-select-file")

                    # Save the first image, appending the rest
                    image_objects_pil[0].save(output_path, "PDF", resolution=100.0, save_all=True, append_images=image_objects_pil[1:])
                    conversion_success = True
                except ValueError as ve: raise ve # Re-raise specific value errors
                except RuntimeError as rte: raise rte # Re-raise specific runtime errors
                except Exception as e:
                     logger.error(f"Unexpected error saving images to PDF: {e}", exc_info=True)
                     raise RuntimeError("err-unknown") from e # Raise generic unknown error
                finally:
                    # Close all opened PIL Image objects
                    for img_obj in image_objects_pil:
                        try: img_obj.close()
                        except: pass
                # --- End Inlined/Adapted Logic ---

        except ValueError as val_err:
            error_key = str(val_err) if str(val_err).startswith("err-") else "err-conversion"
            logger.error(f"Image conversion ValueError: {error_key}", exc_info=False)
        except RuntimeError as rt_err:
            error_key = str(rt_err) if str(rt_err).startswith("err-") else "err-conversion"
            logger.error(f"Image conversion RuntimeError: {error_key}", exc_info=False)
        except Exception as conv_err:
            error_key = "err-unknown"
            logger.error(f"Unexpected error during image conversion process: {conv_err}", exc_info=True)
        # --- End Conversion ---

        # --- Handle Success or Failure ---
        if conversion_success and output_path and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            mimetype = 'application/zip' if out_ext == 'zip' else 'application/pdf'
            try:
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                @response.call_on_close
                def cleanup_image_success():
                    logger.debug(f"Cleaning up successful /convert_image: Inputs: {saved_input_paths}, Output: {output_path}, TempDir: {temp_upload_dir}")
                    # Clean up the main input file (PDF) or the temp image files
                    for p in saved_input_paths: safe_remove(p)
                    safe_remove(output_path) # Clean up the final output PDF/ZIP
                    safe_remove(temp_upload_dir) # Clean up the temporary image directory
                logger.info(f"Image conversion successful. Sending file: {output_filename}. Time: {time.time() - start_time:.2f}s")
                return response
            except Exception as send_err:
                logger.error(f"Error sending image conversion file {output_filename}: {send_err}", exc_info=True)
                raise RuntimeError("err-unknown") from send_err
        else:
            final_error_key = error_key or "err-conversion"
            logger.error(f"Image conversion failed or produced invalid output. Final Error Key: {final_error_key}. Time: {time.time() - start_time:.2f}s")
            raise RuntimeError(final_error_key)

    # --- Outer Exception Handler ---
    except Exception as e:
        final_error_key = str(e) if str(e).startswith("err-") else "err-unknown"
        status_code = 400
        if final_error_key == "err-unknown":
            status_code = 500
            logger.error(f"Unexpected error in /convert_image handler: {e}", exc_info=True)
        elif final_error_key == "err-file-too-large": status_code = 413
        elif final_error_key == "err-rate-limit-exceeded": status_code = 429
        elif final_error_key == "err-csrf-invalid": status_code = 400

        # --- Cleanup for Failed Request ---
        logger.debug(f"Cleaning up failed /convert_image request (Error: {final_error_key}).")
        # Clean up main input (PDF) or temporary saved images
        for p in saved_input_paths: safe_remove(p)
        safe_remove(output_path) # Clean potential output file
        safe_remove(temp_upload_dir) # Clean temporary image directory
        # --- End Cleanup ---

        return make_error_response(final_error_key, status_code)


# --- Teardown (Cleanup old files in UPLOAD_FOLDER) ---
@app.teardown_appcontext
def cleanup_old_files(exception=None):
    # (Giữ nguyên hàm cleanup_old_files)
    if not os.path.exists(UPLOAD_FOLDER): return
    logger.debug("Running teardown cleanup for UPLOAD_FOLDER...")
    try:
        now = time.time(); max_age = 3600 # 1 hour
        deleted_count = 0; checked_count = 0
        try: items = os.listdir(UPLOAD_FOLDER)
        except OSError as list_err: logger.error(f"Teardown: Listdir error {UPLOAD_FOLDER}: {list_err}"); return

        for item_name in items:
            # --- ADDED: Skip temporary directories used by image conversion ---
            if item_name.startswith("img2pdf_") or item_name.startswith("pdfimg_") or item_name.startswith("pdf2imgzip_"):
                 logger.debug(f"Teardown: Skipping temporary item: {item_name}")
                 # Optionally, you could try to remove old *temp directories* here too,
                 # but safe_remove on the directory path should handle it eventually.
                 # For safety, only remove files directly for now.
                 continue
            # --- END ADDED ---
            path = os.path.join(UPLOAD_FOLDER, item_name)
            try:
                 if os.path.isfile(path):
                     stat_result = os.stat(path)
                     file_age = now - stat_result.st_mtime; checked_count += 1
                     if file_age > max_age:
                         if safe_remove(path): deleted_count += 1
                 # Removed directory check to avoid accidentally removing temp dirs used by active requests
                 # elif os.path.isdir(path): pass
            except FileNotFoundError: continue # File might have been removed by another process/request
            except Exception as e: logger.error(f"Teardown check error {path}: {e}")

        if checked_count > 0 or deleted_count > 0: logger.info(f"Teardown: Checked {checked_count}, removed {deleted_count} files older than {max_age}s from {UPLOAD_FOLDER}.")
        else: logger.debug("Teardown: No old files found/removed in UPLOAD_FOLDER.")
    except Exception as e: logger.error(f"Teardown critical error: {e}", exc_info=True)

# === Main Execution ===
if __name__ == '__main__':
    # (Giữ nguyên phần main execution)
    try:
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        logger.info(f"Upload folder created/exists: {os.path.abspath(UPLOAD_FOLDER)}")
    except OSError as mkdir_err:
        logger.critical(f"FATAL: Cannot create upload folder {UPLOAD_FOLDER}: {mkdir_err}.")
        sys.exit(1)

    # Log trạng thái SOFFICE_PATH đã được xử lý ở trên
    csrf_enabled = app.config.get('WTF_CSRF_ENABLED', True) # Default is True
    logger.info(f"CSRF Protection Enabled: {csrf_enabled}")
    logger.info(f"Rate Limiting Enabled: Yes (Default limits active)")

    port = int(os.environ.get('PORT', 5003))
    host = os.environ.get('HOST', '0.0.0.0')
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() in ['true', '1', 't']

    logger.info(f"Starting server on {host}:{port} - Debug: {debug_mode}")

    if debug_mode:
        logger.warning("Running in Flask DEBUG mode (Insecure for production).")
        # use_reloader=False might be needed if startup checks cause issues with reloader
        app.run(host=host, port=port, debug=True, threaded=True, use_reloader=True)
    else:
        logger.info("Running with Waitress production server.")
        try:
            from waitress import serve
            serve(app, host=host, port=port, threads=8) # Adjust threads as needed
        except ImportError:
            logger.critical("Waitress not found! Cannot start production server.")
            logger.warning("FALLING BACK TO FLASK DEVELOPMENT SERVER (NOT RECOMMENDED FOR PRODUCTION).")
            app.run(host=host, port=port, debug=False, threaded=True)