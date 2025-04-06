# --- START OF FILE app.py ---

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
from pdf2docx import Converter # Cần cho bước chuyển đổi cuối
import tempfile
import PyPDF2
import shutil # Cần cho safe_remove, which
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
# IMPORTANT: Adjust x_for based on your proxy setup (e.g., Nginx)
# If only Nginx is in front, x_for=1 is correct.
app.wsgi_app = ProxyFix(
    app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1
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
    default_limits=["419 per day", "210 per hour", "30 per minute"],
    storage_uri="memory://",
    strategy="fixed-window"
)
# Slightly refined CSP
csp = {
    'default-src': ['\'self\''],
    'style-src': ['\'self\'', '\'unsafe-inline\'', 'https://cdn.tailwindcss.com', 'https://fonts.googleapis.com'],
    'script-src': ['\'self\'', '\'unsafe-inline\''], # Consider using nonce for stricter security if needed
    'font-src': ['\'self\'', 'https://fonts.gstatic.com'],
    'img-src': ['\'self\'', 'data:'],
    'form-action': '\'self\''
}
# Talisman Configuration - Enable security features assuming Nginx handles HTTPS
talisman = Talisman(
    app,
    content_security_policy=csp,
    force_https=True,             # IMPORTANT: Ensure Nginx handles HTTPS correctly first
    session_cookie_secure=True,   # IMPORTANT: Requires HTTPS
    session_cookie_http_only=True,
    frame_options='DENY',
    strict_transport_security=True, # IMPORTANT: Only if HTTPS is stable and you intend to keep it
    strict_transport_security_max_age=31536000, # 1 year
    strict_transport_security_include_subdomains=True,
    # content_security_policy_nonce_in=['script-src'] # Add if using nonce
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
LIBREOFFICE_TIMEOUT = 180
GS_TIMEOUT = 180
MIME_BUFFER_SIZE = 4096

# === Helper Functions ===
def make_error_response(error_key, status_code=400):
    logger.warning(f"Returning error: {error_key} (Status: {status_code})")
    response_text = f"Conversion failed: {error_key}"
    response = make_response(response_text, status_code)
    response.headers["Content-Type"] = "text/plain; charset=utf-8"
    return response

# --- Logic tìm và xác minh LibreOffice (Keep as is) ---
_VERIFIED_SOFFICE_PATH = '/usr/lib/libreoffice/program/soffice'
SOFFICE_PATH = None
if os.path.isfile(_VERIFIED_SOFFICE_PATH):
    try:
        result = subprocess.run([_VERIFIED_SOFFICE_PATH, '--headless', '--version'], capture_output=True, text=True, check=False, timeout=15)
        if result.returncode == 0 and 'LibreOffice' in result.stdout:
            logger.info(f"Using verified hardcoded LO path: {_VERIFIED_SOFFICE_PATH}")
            SOFFICE_PATH = _VERIFIED_SOFFICE_PATH
        else:
            logger.warning(f"Hardcoded LO path {_VERIFIED_SOFFICE_PATH} exists, but version check failed! Code: {result.returncode}, Output: {result.stdout.strip()}")
    except subprocess.TimeoutExpired: logger.warning(f"Timeout expired verifying hardcoded LO path: {_VERIFIED_SOFFICE_PATH}")
    except Exception as e: logger.warning(f"Error verifying hardcoded LO path {_VERIFIED_SOFFICE_PATH}: {e}")
else: logger.warning(f"Hardcoded LO path {_VERIFIED_SOFFICE_PATH} does not exist or is not a file.")

if not SOFFICE_PATH:
    logger.info("Hardcoded LO path not verified, trying shutil.which('libreoffice')...")
    soffice_found = shutil.which('libreoffice') # Might find '/usr/bin/libreoffice' which is a script
    if soffice_found:
        # Attempt to resolve the actual binary path if soffice_found is a script
        resolved_soffice_path = soffice_found
        if os.path.islink(soffice_found):
            try:
                link_target = os.readlink(soffice_found)
                # Simple check if it points to the program dir (might need adjustment)
                if 'program/soffice' in link_target:
                     resolved_soffice_path = os.path.join(os.path.dirname(soffice_found), link_target) if not link_target.startswith('/') else link_target
                     logger.info(f"Resolved symlink {soffice_found} to {resolved_soffice_path}")
                # Check if the default path exists relative to the script
                script_dir = os.path.dirname(soffice_found)
                potential_path = os.path.join(script_dir, "../lib/libreoffice/program/soffice")
                if os.path.isfile(potential_path):
                    resolved_soffice_path = os.path.abspath(potential_path)
                    logger.info(f"Found soffice relative to script: {resolved_soffice_path}")

            except Exception as link_err:
                 logger.warning(f"Error resolving symlink {soffice_found}: {link_err}")

        try:
            # Use the potentially resolved path for version check
            logger.info(f"Verifying LO path: {resolved_soffice_path}")
            result = subprocess.run([resolved_soffice_path, '--headless', '--version'], capture_output=True, text=True, check=False, timeout=15)
            if result.returncode == 0 and 'LibreOffice' in result.stdout:
                logger.info(f"Using LO path found via shutil.which/resolution: {resolved_soffice_path}")
                SOFFICE_PATH = resolved_soffice_path # Use the resolved path
            else:
                logger.warning(f"Found/Resolved LO path {resolved_soffice_path}, but version check failed! Code: {result.returncode}, Output: {result.stdout.strip()}")
                # Fallback check for the hardcoded path again, just in case resolution failed
                if os.path.isfile(_VERIFIED_SOFFICE_PATH):
                     result_hc = subprocess.run([_VERIFIED_SOFFICE_PATH, '--headless', '--version'], capture_output=True, text=True, check=False, timeout=15)
                     if result_hc.returncode == 0 and 'LibreOffice' in result_hc.stdout:
                         logger.info(f"Fallback: Using verified hardcoded LO path after which failed: {_VERIFIED_SOFFICE_PATH}")
                         SOFFICE_PATH = _VERIFIED_SOFFICE_PATH

        except subprocess.TimeoutExpired: logger.warning(f"Timeout expired verifying LO path via which/resolution: {resolved_soffice_path}")
        except Exception as e: logger.warning(f"Error verifying LO path via which/resolution {resolved_soffice_path}: {e}")
    else: logger.warning("shutil.which('libreoffice') did not find an executable.")


if SOFFICE_PATH: logger.info(f"Successfully set LO path for use: {SOFFICE_PATH}")
else: logger.critical("LibreOffice could NOT be set/verified. Conversions requiring it WILL FAIL.")
# --- Kết thúc logic LibreOffice ---

# --- Logic tìm và xác minh Ghostscript (Keep as is) ---
GS_PATH = None
gs_executable_name = 'gs'
gs_found_path = shutil.which(gs_executable_name)
if gs_found_path:
    try:
        result = subprocess.run( [gs_found_path, '--version'], capture_output=True, text=True, check=False, timeout=10 )
        if result.returncode == 0 and '.' in result.stdout.strip():
            logger.info(f"Using Ghostscript path found via shutil.which: {gs_found_path} (Version: {result.stdout.strip()})")
            GS_PATH = gs_found_path
        else:
             logger.warning(f"Found potential GS path {gs_found_path}, but version check failed! Code: {result.returncode}, Output: {result.stdout.strip()}")
    except subprocess.TimeoutExpired: logger.warning(f"Timeout expired while verifying GS path: {gs_found_path}")
    except Exception as e: logger.warning(f"Error verifying GS path {gs_found_path}: {e}")
else: logger.warning(f"shutil.which('{gs_executable_name}') did not find an executable.")

if not GS_PATH: logger.critical("Ghostscript ('gs') could NOT be found or verified. PDF Compression WILL FAIL.")
# --- Kết thúc logic Ghostscript ---


def _allowed_file_extension(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def safe_remove(item_path, retries=3, delay=0.5):
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
    mime_type = None
    try:
        original_pos = file_storage.stream.tell()
        file_storage.stream.seek(0)
        buffer = file_storage.stream.read(MIME_BUFFER_SIZE)
        file_storage.stream.seek(original_pos)
        # Ensure python-magic uses the system's magic file if possible
        mime_type = magic.from_buffer(buffer, mime=True)
        logger.debug(f"Detected MIME type: {mime_type} for file {file_storage.filename}")
    except magic.MagicException as e: logger.warning(f"Could not determine MIME type for {file_storage.filename}: {e}")
    except Exception as e: logger.error(f"Unexpected error during MIME detection for {file_storage.filename}: {e}")
    return mime_type

# --- Other Helper Functions (Keep as is) ---
# get_pdf_page_size, setup_slide_size, sort_key_for_pptx_images,
# _convert_pdf_to_pptx_images, convert_pdf_to_pptx_python,
# convert_images_to_pdf, convert_pdf_to_image_zip,
# compress_pdf_ghostscript functions remain unchanged.

# === Global Error Handlers (Keep as is) ===
@app.errorhandler(CSRFError)
def handle_csrf_error(e): logger.warning(f"CSRF failed: {e.description}"); return make_error_response("err-csrf-invalid", 400)
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e): logger.warning(f"File too large: {e.description}"); return make_error_response("err-file-too-large", 413)
@app.errorhandler(429)
def ratelimit_handler(e): logger.warning(f"Rate limit exceeded: {e.description}"); return make_error_response("err-rate-limit-exceeded", 429)
@app.errorhandler(Exception)
def handle_generic_exception(e):
     from werkzeug.exceptions import HTTPException
     if isinstance(e, HTTPException): return e
     logger.error(f"Unhandled Exception: {e}", exc_info=True); return make_error_response("err-unknown", 500)

# === Routes ===

@app.route('/api/translations')
def get_translations():
    """Provides translation strings to the frontend."""
    translations = {
        'en': {
            # --- SEO RELATED STRINGS (NEW/UPDATED) ---
            'lang-title': 'Online PDF Converter - Convert & Compress PDF Free | Pdfsconvert.com', # Updated Title
            'lang-meta-description': 'Easily convert PDF to Word (DOCX), PPT, JPG & compress PDF files online for free with Pdfsconvert.com. Fast, secure, no registration needed.', # New Meta Description
            # --- EXISTING STRINGS ---
            'lang-subtitle': 'Simple, powerful tools for your documents',
            'lang-error-title': 'Error!', 'lang-convert-title': 'Convert PDF/Office',
            'lang-convert-desc': 'Transform PDF to Word (DOCX) & PowerPoint (PPTX) online', # Slightly more specific
            'lang-compress-title': 'Compress PDF', 'lang-compress-desc': 'Reduce PDF file size online while optimizing for quality',
            'lang-compress-input-label': 'Select PDF file', 'lang-compress-btn': 'Compress PDF',
            'lang-compressing': 'Compressing PDF...', 'lang-select-quality': 'Compression Level',
            'lang-quality-low': 'High Compression (Smallest Size)', # Renamed for clarity
            'lang-quality-medium': 'Medium Compression (Good Balance)',
            'lang-quality-high': 'Low Compression (Best Quality)',
            'lang-merge-title': 'Merge PDF', 'lang-merge-desc': 'Combine multiple PDFs into one file',
            'lang-split-title': 'Split PDF', 'lang-split-desc': 'Extract pages from your PDF',
            'lang-image-title': 'PDF ↔ Image (JPG)', 'lang-image-desc': 'Convert PDF to JPG images or JPG images to PDF online', # Specific format
            'lang-image-input-label': 'Select PDF or Image(s) (JPG/JPEG only)', 'lang-image-convert-btn': 'Convert Now',
            'lang-image-converting': 'Converting...', 'lang-size-limit': 'Size limit: 100MB',
            'lang-size-limit-total': 'Size limit: 100MB (total)', 'lang-select-conversion': 'Select conversion type',
            'lang-converting': 'Converting...', 'lang-convert-btn': 'Convert Now',
            'lang-file-input-label': 'Select file', 'file-no-selected': 'No file selected',
            'err-select-file': 'Please select file(s).', 'err-file-too-large': 'File size exceeds the limit (100MB).',
            'err-select-conversion': 'Please select a conversion type.',
            'err-format-docx': 'Select one DOCX file for this operation.',
            'err-format-ppt': 'Select one PDF, PPT or PPTX file for this conversion.',
            'err-format-pdf': 'Please select a PDF file.', 'err-conversion': 'An error occurred during processing.',
            'err-fetch-translations': 'Could not load language data.', 'lang-select-btn-text': 'Browse',
            'lang-select-conversion-label': 'Conversion Type', 'err-multi-file-not-supported': 'Multi-file selection is only supported for Image to PDF conversion.',
            'err-invalid-image-file': 'One or more selected files are not valid images (Pillow error).',
            'err-image-format': 'Invalid file type. Select PDF, JPG, or JPEG based on conversion.',
            'err-image-single-pdf': 'Please select only one PDF file to convert to images.',
            'err-image-all-images': 'If selecting multiple files, all must be JPG or JPEG to convert to PDF.',
            'err-libreoffice': 'Conversion failed (Processing engine error - LO).', 'err-conversion-timeout': 'Processing timed out.',
            'err-poppler-missing': 'PDF processing library (Poppler) missing or failed.',
            'err-pdf-corrupt': 'Could not process PDF (corrupt file?).', 'err-unknown': 'An unexpected error occurred. Please try again later.',
            'err-csrf-invalid': 'Security validation failed. Please refresh the page and try again.',
            'err-rate-limit-exceeded': 'Too many requests. Please wait a moment and try again.',
            'err-invalid-mime-type': 'Invalid file type detected. The file content does not match the expected format.',
            'err-mime-unidentified-office': "Could not identify file type, it might be non-standard. Please open your file in an Office application, press 'Save' or 'Save as' to save again and upload again.",
            'err-invalid-mime-type-image': 'Invalid image type detected. Only JPEG files are allowed for Image-to-PDF.',
            'err-pdf-protected': 'Cannot process password-protected PDF.',
            'err-poppler-check-failed': 'Failed to get PDF info (Poppler check).',
            'err-conversion-img': 'Failed to convert/extract images from PDF.',
            'err-gs-missing': 'Compression engine (Ghostscript) not available.',
            'err-gs-failed': 'Compression failed (Ghostscript error). Check if PDF is valid/not protected.',
            'err-gs-timeout': 'Compression timed out.', 'err-invalid-quality': 'Invalid compression quality selected.',
            'lang-clear-all': 'Clear All', 'lang-upload-a-file': 'Upload files',
            'lang-drag-drop': 'or drag and drop', 'lang-image-types': 'PDF, JPG, JPEG up to 100MB total',
            'lang-compress-docx-title': 'Compress Word',
            'lang-compress-docx-desc': 'Reduce Word DOCX file size online (via PDF compression)', # Updated desc
            'lang-compress-docx-input-label': 'Select Word file',
            'lang-compressing-docx': 'Compressing Word...',
            'lang-compress-docx-btn': 'Compress Word',
        },
        'vi': {
            # --- SEO RELATED STRINGS (NEW/UPDATED) ---
            'lang-title': 'Chuyển Đổi PDF Online Miễn Phí - Nén & Convert PDF | Pdfsconvert.com', # Updated Title
            'lang-meta-description': 'Chuyển đổi PDF sang Word (DOCX), PPT, JPG & nén file PDF trực tuyến miễn phí dễ dàng tại Pdfsconvert.com. Nhanh chóng, bảo mật, không cần đăng ký.', # New Meta Description
            # --- EXISTING STRINGS ---
            'lang-subtitle': 'Công cụ đơn giản, mạnh mẽ cho tài liệu của bạn',
            'lang-error-title': 'Lỗi!', 'lang-convert-title': 'Chuyển đổi PDF/Office',
            'lang-convert-desc': 'Chuyển đổi PDF sang Word (DOCX) & PowerPoint (PPTX) trực tuyến', # Slightly more specific
            'lang-compress-title': 'Nén PDF', 'lang-compress-desc': 'Giảm dung lượng tệp PDF trực tuyến mà vẫn tối ưu chất lượng',
            'lang-compress-input-label': 'Chọn tệp PDF', 'lang-compress-btn': 'Nén PDF',
            'lang-compressing': 'Đang nén PDF...', 'lang-select-quality': 'Mức độ nén',
            'lang-quality-low': 'Nén Mạnh (Nhẹ Nhất)', # Renamed for clarity
            'lang-quality-medium': 'Nén Vừa (Cân Bằng)',
            'lang-quality-high': 'Nén Nhẹ (Chất lượng cao)',
            'lang-merge-title': 'Gộp PDF', 'lang-merge-desc': 'Kết hợp nhiều tệp PDF thành một tệp',
            'lang-split-title': 'Tách PDF', 'lang-split-desc': 'Trích xuất các trang từ tệp PDF của bạn',
            'lang-image-title': 'PDF ↔ Ảnh (JPG)', 'lang-image-desc': 'Chuyển PDF thành ảnh JPG hoặc ảnh JPG thành PDF trực tuyến', # Specific format
            'lang-image-input-label': 'Chọn PDF hoặc (các) Ảnh (chỉ JPG/JPEG)', 'lang-image-convert-btn': 'Chuyển đổi ngay',
            'lang-image-converting': 'Đang chuyển đổi...', 'lang-size-limit': 'Giới hạn kích thước: 100MB',
            'lang-size-limit-total': 'Giới hạn kích thước: 100MB (tổng)', 'lang-select-conversion': 'Chọn kiểu chuyển đổi',
            'lang-converting': 'Đang chuyển đổi...', 'lang-convert-btn': 'Chuyển đổi ngay',
            'lang-file-input-label': 'Chọn tệp', 'file-no-selected': 'Không có tệp nào được chọn',
            'err-select-file': 'Vui lòng chọn (các) tệp.', 'err-file-too-large': 'Kích thước tệp vượt quá giới hạn (100MB).',
            'err-select-conversion': 'Vui lòng chọn kiểu chuyển đổi.',
            'err-format-docx': 'Chọn một file DOCX cho thao tác này.',
            'err-format-ppt': 'Chọn một file PDF, PPT hoặc PPTX cho chuyển đổi này.',
            'err-format-pdf': 'Vui lòng chọn một tệp PDF.', 'err-conversion': 'Đã xảy ra lỗi trong quá trình xử lý.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.', 'lang-select-btn-text': 'Duyệt...',
            'lang-select-conversion-label': 'Kiểu chuyển đổi', 'err-multi-file-not-supported': 'Chỉ hỗ trợ chọn nhiều file khi chuyển đổi Ảnh sang PDF.',
            'err-invalid-image-file': 'Một hoặc nhiều tệp được chọn không phải là ảnh hợp lệ (lỗi Pillow).',
            'err-image-format': 'Loại tệp không hợp lệ. Chọn PDF, JPG, hoặc JPEG tùy theo chuyển đổi.',
            'err-image-single-pdf': 'Vui lòng chỉ chọn một file PDF để chuyển đổi sang ảnh.',
            'err-image-all-images': 'Nếu chọn nhiều tệp, tất cả phải là JPG hoặc JPEG để chuyển đổi sang PDF.',
            'err-libreoffice': 'Chuyển đổi thất bại (Lỗi bộ xử lý - LO).', 'err-conversion-timeout': 'Quá trình xử lý quá thời gian.',
            'err-poppler-missing': 'Thiếu hoặc lỗi thư viện xử lý PDF (Poppler).',
            'err-pdf-corrupt': 'Không thể xử lý PDF (tệp lỗi?).', 'err-unknown': 'Đã xảy ra lỗi không mong muốn. Vui lòng thử lại sau.',
            'err-csrf-invalid': 'Xác thực bảo mật thất bại. Vui lòng tải lại trang và thử lại.',
            'err-rate-limit-exceeded': 'Quá nhiều yêu cầu. Vui lòng đợi một lát và thử lại.',
            'err-invalid-mime-type': 'Phát hiện loại tệp không hợp lệ. Nội dung tệp không khớp định dạng mong đợi.',
            'err-mime-unidentified-office': "Không thể nhận dạng loại file dù có đuôi Office. Vui lòng mở file của bạn lên bằng ứng dụng Office, ấn 'Lưu' hoặc 'Lưu thành' để lưu lại bản mới và tải lên lại.",
            'err-invalid-mime-type-image': 'Phát hiện loại ảnh không hợp lệ. Chỉ cho phép tệp JPEG để chuyển đổi Ảnh sang PDF.',
            'err-pdf-protected': 'Không thể xử lý PDF được bảo vệ bằng mật khẩu.',
            'err-poppler-check-failed': 'Không thể lấy thông tin PDF (lỗi kiểm tra Poppler).',
            'err-conversion-img': 'Không thể chuyển đổi/trích xuất ảnh từ PDF.',
            'err-gs-missing': 'Không tìm thấy công cụ nén (Ghostscript).',
            'err-gs-failed': 'Nén thất bại (Lỗi Ghostscript). Kiểm tra PDF hợp lệ/không bị khóa.',
            'err-gs-timeout': 'Nén quá thời gian.', 'err-invalid-quality': 'Đã chọn mức nén không hợp lệ.',
            'lang-clear-all': 'Xóa tất cả', 'lang-upload-a-file': 'Tải tệp lên',
            'lang-drag-drop': 'hoặc kéo và thả', 'lang-image-types': 'PDF, JPG, JPEG tối đa 100MB tổng',
            'lang-compress-docx-title': 'Nén Word',
            'lang-compress-docx-desc': 'Giảm dung lượng tệp Word DOCX trực tuyến (thông qua nén PDF)', # Updated desc
            'lang-compress-docx-input-label': 'Chọn tệp DOCX',
            'lang-compressing-docx': 'Đang nén Word...',
            'lang-compress-docx-btn': 'Nén Word',
        }
    }
    lang = request.args.get('lang', 'en')
    return jsonify(translations.get(lang, translations.get('en', {})))


@app.route('/')
def index():
    try:
        # Pass the base URL for translations (relative path is usually fine)
        translations_url = url_for('get_translations', _external=False)
        gs_available = GS_PATH is not None
        soffice_available = SOFFICE_PATH is not None
        return render_template('index.html',
                               translations_url=translations_url,
                               gs_available=gs_available,
                               soffice_available=soffice_available)
    except Exception as e:
        logger.error(f"Error rendering index page: {e}", exc_info=True)
        # Use the helper for consistent error response
        return make_error_response("err-unknown", 500)

# --- Routes /convert, /convert_image, /compress_pdf, /compress_docx (Keep as is) ---
# These routes handle the backend logic and don't need SEO changes directly.

@app.teardown_appcontext
def cleanup_old_files(exception=None):
    # Keep teardown logic as is
    if not os.path.exists(UPLOAD_FOLDER): return
    logger.debug("Running teardown cleanup...")
    try:
        now = time.time(); max_age = 3600; deleted_count = 0; checked_count = 0
        try: items = os.listdir(UPLOAD_FOLDER)
        except OSError as list_err: logger.error(f"Teardown listdir error: {list_err}"); return
        for item_name in items:
            # Improved check for temporary files/dirs (more robust)
            if item_name and any(item_name.startswith(prefix) for prefix in ["img2pdf_", "pdfimg_", "pdf2imgzip_", "temp_", "input_"]):
                 # Check if it's a directory specifically created by tempfile
                 path = os.path.join(UPLOAD_FOLDER, item_name)
                 if os.path.isdir(path) and any(prefix in item_name for prefix in ["img2pdf_", "pdfimg_", "pdf2imgzip_"]):
                     logger.debug(f"Teardown removing temp dir: {item_name}")
                     safe_remove(path)
                     deleted_count +=1
                 elif os.path.isfile(path) and item_name.startswith("input_"): # Clean input files too
                     stat_result = os.stat(path); file_age = now - stat_result.st_mtime; checked_count += 1
                     if file_age > max_age:
                         if safe_remove(path): deleted_count += 1
                 # Skip other temp-like names unless they are old files
                 elif os.path.isfile(path):
                     stat_result = os.stat(path); file_age = now - stat_result.st_mtime; checked_count += 1
                     if file_age > max_age:
                         if safe_remove(path): deleted_count += 1
                 continue # Skip further checks for known temp patterns

            # Check other files by age
            path = os.path.join(UPLOAD_FOLDER, item_name)
            try:
                 if os.path.isfile(path):
                     stat_result = os.stat(path); file_age = now - stat_result.st_mtime; checked_count += 1
                     if file_age > max_age:
                         if safe_remove(path): deleted_count += 1
                 elif os.path.isdir(path): # Also check and remove old directories if needed
                      stat_result = os.stat(path); dir_age = now - stat_result.st_mtime; checked_count += 1
                      if dir_age > max_age:
                           logger.warning(f"Teardown attempting remove old directory: {path}")
                           if safe_remove(path): deleted_count += 1
            except FileNotFoundError: continue
            except Exception as e: logger.warning(f"Teardown check error for {path}: {e}")
        if checked_count > 0 or deleted_count > 0: logger.info(f"Teardown: Checked {checked_count} items, removed {deleted_count} items older than {max_age}s.")
        else: logger.debug("Teardown: No old items found/removed.")
    except Exception as e: logger.error(f"Teardown critical error: {e}", exc_info=True)


if __name__ == '__main__':
    try: os.makedirs(UPLOAD_FOLDER, exist_ok=True); logger.info(f"Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    except OSError as mkdir_err: logger.critical(f"FATAL: Cannot create upload folder {UPLOAD_FOLDER}: {mkdir_err}."); sys.exit(1)
    logger.info(f"LibreOffice Path: {SOFFICE_PATH if SOFFICE_PATH else 'Not Found/Verified'}")
    logger.info(f"Ghostscript Path: {GS_PATH if GS_PATH else 'Not Found/Verified'}")
    csrf_enabled = app.config.get('WTF_CSRF_ENABLED', True); logger.info(f"CSRF Protection Enabled: {csrf_enabled}")
    logger.info(f"Rate Limiting Enabled: Yes (Default limits active)")
    logger.info(f"Talisman Security Headers Enabled: Yes (HTTPS forced: {talisman.force_https}, HSTS enabled: {talisman.strict_transport_security})")
    port = int(os.environ.get('PORT', 5003)); host = os.environ.get('HOST', '0.0.0.0'); debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() in ['true', '1', 't']
    logger.info(f"Starting server on {host}:{port} - Debug: {debug_mode}")
    if debug_mode: logger.warning("Running in Flask DEBUG mode."); app.run(host=host, port=port, debug=True, threaded=True, use_reloader=True)
    else:
        logger.info("Running with Waitress production server.")
        try: from waitress import serve; serve(app, host=host, port=port, threads=4)
        except ImportError: logger.critical("Waitress not found!"); logger.warning("FALLING BACK TO FLASK DEV SERVER."); app.run(host=host, port=port, debug=False, threaded=True)

# --- END OF FILE app.py ---