# --- START OF FILE app.py ---

from flask import Flask, request, send_file, render_template, jsonify, url_for
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
from pdf2image import convert_from_path, pdfinfo_from_path
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image, UnidentifiedImageError # Import UnidentifiedImageError
from docx import Document
import zipfile

app = Flask(__name__, template_folder='templates', static_folder='static')

app.config['MAX_CONTENT_LENGTH'] = 150 * 1024 * 1024  # 150MB

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
# Chỉ cần các extension này cho 2 chức năng chính
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}
# Extension chỉ cho chức năng PDF/Image
ALLOWED_IMAGE_EXTENSIONS = {'pdf', 'jpg', 'jpeg'}


@app.route('/health')
def health_check():
    return 'OK', 200

@app.route('/get_translations')
def get_translations():
    translations = {
        'en': {
            'lang-title': 'PDF & Office Tools', # Renamed slightly
            'lang-subtitle': 'Simple, powerful tools for your documents', # Renamed slightly
            'lang-error-title': 'Error!',
            'lang-convert-title': 'Convert PDF/Office', # More specific title
            'lang-convert-desc': 'Transform PDF to Word/PPT and vice versa', # More specific desc
            'lang-compress-title': 'Compress PDF',
            'lang-compress-desc': 'Reduce file size while maintaining quality',
            'lang-merge-title': 'Merge PDF',
            'lang-merge-desc': 'Combine multiple PDFs into one file',
            'lang-split-title': 'Split PDF',
            'lang-split-desc': 'Extract pages from your PDF',
            'lang-rotate-title': 'Rotate PDF',
            'lang-rotate-desc': 'Change page orientation',
            # --- NEW PDF <-> Image Card Translations ---
            'lang-image-title': 'PDF ↔ Image',
            'lang-image-desc': 'Convert PDF to images or images to PDF',
            'lang-image-input-label': 'Select PDF or Image(s)',
            'lang-image-convert-btn': 'Convert Now',
            'lang-image-converting': 'Converting...',
            # -----------------------------------------
            'lang-size-limit': 'Size limit: 100MB (total)',
            'lang-select-conversion': 'Select conversion type',
            'lang-converting': 'Converting...', # Keep this generic one too? Maybe remove if unused.
            'lang-convert-btn': 'Convert Now', # Keep this generic one too? Maybe remove if unused.
            'lang-file-input-label': 'Select file', # Label for the first card
            'file-no-selected': 'No file selected',
            'err-select-file': 'Please select file(s) to convert.',
            'err-file-too-large': 'Total file size exceeds the limit (100MB).',
            'err-select-conversion': 'Please select a conversion type.',
            'err-format-docx': 'Select one PDF or DOCX file for this conversion.', # Updated text
            'err-format-ppt': 'Select one PDF, PPT or PPTX file for this conversion.', # Updated text
            # 'err-format-jpg': No longer needed in the first card
            'err-conversion': 'An error occurred during conversion.',
            'err-fetch-translations': 'Could not load language data.',
            'lang-select-btn-text': 'Browse',
            'lang-select-conversion-label': 'Conversion Type',
            'err-multi-file-not-supported': 'Multi-file selection is only supported for Image to PDF conversion.',
            'err-invalid-image-file': 'One or more selected files are not valid images.',
            # --- NEW PDF <-> Image Card Errors ---
            'err-image-format': 'Invalid file type. Select PDF, JPG, or JPEG.',
            'err-image-single-pdf': 'Please select only one PDF file to convert to images.',
            'err-image-all-images': 'If selecting multiple files, all must be JPG or JPEG to convert to PDF.',
            # ------------------------------------
        },
        'vi': {
            'lang-title': 'Công Cụ PDF & Văn Phòng',
            'lang-subtitle': 'Công cụ đơn giản, mạnh mẽ cho tài liệu của bạn',
            'lang-error-title': 'Lỗi!',
            'lang-convert-title': 'Chuyển đổi PDF/Office', # More specific title
            'lang-convert-desc': 'Chuyển đổi PDF sang Word/PPT và ngược lại', # More specific desc
            'lang-compress-title': 'Nén PDF',
            'lang-compress-desc': 'Giảm kích thước tệp trong khi duy trì chất lượng',
            'lang-merge-title': 'Gộp PDF',
            'lang-merge-desc': 'Kết hợp nhiều tệp PDF thành một tệp',
            'lang-split-title': 'Tách PDF',
            'lang-split-desc': 'Trích xuất các trang từ tệp PDF của bạn',
            'lang-rotate-title': 'Xoay PDF',
            'lang-rotate-desc': 'Thay đổi hướng trang',
             # --- NEW PDF <-> Image Card Translations ---
            'lang-image-title': 'PDF ↔ Ảnh',
            'lang-image-desc': 'Chuyển PDF thành ảnh hoặc ảnh thành PDF',
            'lang-image-input-label': 'Chọn PDF hoặc (các) Ảnh',
            'lang-image-convert-btn': 'Chuyển đổi ngay',
            'lang-image-converting': 'Đang chuyển đổi...',
            # -----------------------------------------
            'lang-size-limit': 'Giới hạn kích thước: 100MB (tổng)',
            'lang-select-conversion': 'Chọn kiểu chuyển đổi',
            'lang-converting': 'Đang chuyển đổi...',
            'lang-convert-btn': 'Chuyển đổi ngay',
            'lang-file-input-label': 'Chọn tệp', # Label for the first card
            'file-no-selected': 'Không có tệp nào được chọn',
            'err-select-file': 'Vui lòng chọn (các) tệp để chuyển đổi.',
            'err-file-too-large': 'Tổng kích thước tệp vượt quá giới hạn (100MB).',
            'err-select-conversion': 'Vui lòng chọn kiểu chuyển đổi.',
            'err-format-docx': 'Chọn một file PDF hoặc DOCX cho chuyển đổi này.', # Updated text
            'err-format-ppt': 'Chọn một file PDF, PPT hoặc PPTX cho chuyển đổi này.', # Updated text
            # 'err-format-jpg': No longer needed in the first card
            'err-conversion': 'Đã xảy ra lỗi trong quá trình chuyển đổi.',
            'err-fetch-translations': 'Không thể tải dữ liệu ngôn ngữ.',
            'lang-select-btn-text': 'Duyệt...',
            'lang-select-conversion-label': 'Kiểu chuyển đổi',
            'err-multi-file-not-supported': 'Chỉ hỗ trợ chọn nhiều file khi chuyển đổi Ảnh sang PDF.',
            'err-invalid-image-file': 'Một hoặc nhiều tệp được chọn không phải là ảnh hợp lệ.',
             # --- NEW PDF <-> Image Card Errors ---
            'err-image-format': 'Loại tệp không hợp lệ. Chọn PDF, JPG, hoặc JPEG.',
            'err-image-single-pdf': 'Vui lòng chỉ chọn một file PDF để chuyển đổi sang ảnh.',
            'err-image-all-images': 'Nếu chọn nhiều tệp, tất cả phải là JPG hoặc JPEG để chuyển đổi sang PDF.',
            # ------------------------------------
        }
    }
    lang = request.args.get('lang', 'en')
    return jsonify(translations.get(lang, translations['en']))

def find_libreoffice():
    # ... (giữ nguyên) ...
    possible_paths = [
        'soffice', '/usr/bin/soffice', '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice', '/usr/lib/libreoffice/program/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
    ]
    for path in possible_paths:
        try:
            resolved_path = shutil.which(path) # Check PATH first
            if resolved_path and os.path.isfile(resolved_path):
                result = subprocess.run([resolved_path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0: logger.info(f"Tìm thấy LibreOffice tại (PATH): {resolved_path}"); return resolved_path
            elif os.path.isfile(path):
                result = subprocess.run([path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0: logger.info(f"Tìm thấy LibreOffice tại (Direct Path): {path}"); return path
        except FileNotFoundError: logger.debug(f"Không tìm thấy LibreOffice tại {path} hoặc qua shutil.which")
        except subprocess.TimeoutExpired: logger.warning(f"Kiểm tra LibreOffice tại {path} bị timeout.")
        except Exception as e: logger.warning(f"Lỗi khi kiểm tra LibreOffice tại {path}: {e}")
    logger.warning("Không tìm thấy LibreOffice thực thi. Sử dụng 'soffice' mặc định.")
    return 'soffice'

SOFFICE_PATH = find_libreoffice()
logger.info(f"Sử dụng đường dẫn LibreOffice: {SOFFICE_PATH}")

def _allowed_file(filename, allowed_set):
    """Kiểm tra extension dựa trên set được cung cấp."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def safe_remove(file_path, retries=5, delay=1):
     # ... (giữ nguyên) ...
    if not file_path or not os.path.exists(file_path): return True
    for i in range(retries):
        try:
            if os.path.isdir(file_path): shutil.rmtree(file_path)
            else: os.remove(file_path)
            # logger.info(f"Đã xóa {'thư mục' if os.path.isdir(file_path) else 'file'} tạm: {file_path}")
            return True
        except PermissionError: logger.warning(f"Không có quyền xóa {file_path} (lần thử {i + 1}). Đang đợi..."); time.sleep(delay * (i + 1))
        except Exception as e: logger.warning(f"Không thể xóa {file_path} (lần thử {i + 1}): {e}"); time.sleep(delay)
    logger.error(f"Xóa {file_path} thất bại sau {retries} lần thử.")
    return False

# --- PDF/PPTX helper functions (giữ nguyên) ---
def get_pdf_page_size(pdf_path):
    # ... (giữ nguyên) ...
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if not reader.pages: logger.warning(f"PDF không có trang nào: {pdf_path}"); return None, None
            page = reader.pages[0]; mediabox = page.mediabox
            if mediabox: return float(mediabox.width), float(mediabox.height)
            else: logger.warning(f"Không tìm thấy mediabox cho trang đầu tiên trong {pdf_path}"); return None, None
    except Exception as e: logger.error(f"Lỗi khi đọc kích thước PDF {pdf_path}: {e}"); return None, None

def setup_slide_size(prs, pdf_path):
    # ... (giữ nguyên) ...
    pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)
    if pdf_width_pt is None or pdf_height_pt is None:
        logger.warning("Không thể đọc kích thước PDF, sử dụng mặc định (10x7.5 inches)"); prs.slide_width, prs.slide_height = Inches(10), Inches(7.5); return prs
    try:
        pdf_width_in, pdf_height_in = pdf_width_pt / 72, pdf_height_pt / 72; max_slide_dim = 50.0
        if pdf_width_in > max_slide_dim or pdf_height_in > max_slide_dim:
            ratio = pdf_width_in / pdf_height_in
            if pdf_width_in >= pdf_height_in: prs.slide_width, prs.slide_height = Inches(max_slide_dim), Inches(max_slide_dim / ratio)
            else: prs.slide_height, prs.slide_width = Inches(max_slide_dim), Inches(max_slide_dim * ratio)
            logger.info(f"Kích thước gốc ({pdf_width_in:.2f}x{pdf_height_in:.2f} in) vượt giới hạn, điều chỉnh thành: {prs.slide_width.inches:.2f}x{prs.slide_height.inches:.2f} in")
        else: prs.slide_width, prs.slide_height = Inches(pdf_width_in), Inches(pdf_height_in)
        logger.info(f"Thiết lập kích thước slide theo PDF: {pdf_width_in:.2f} x {pdf_height_in:.2f} inches"); return prs
    except Exception as e: logger.warning(f"Lỗi khi thiết lập kích thước slide từ PDF, sử dụng mặc định: {e}"); prs.slide_width, prs.slide_height = Inches(10), Inches(7.5); return prs

def _convert_pdf_to_pptx_images(input_path, output_path):
    # ... (giữ nguyên) ...
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_"); logger.info(f"Tạo thư mục tạm cho ảnh PPTX: {temp_dir}")
        try: pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None); logger.info(f"PDF info read successfully. Pages: {pdfinfo.get('Pages', 'N/A')}")
        except Exception as info_err: logger.error(f"Lỗi khi chạy pdfinfo (kiểm tra Poppler): {info_err}.", exc_info=True); raise ValueError("Could not process PDF, Poppler might be missing or not configured correctly.") from info_err
        images = convert_from_path(input_path, dpi=300, fmt='jpeg', output_folder=temp_dir, thread_count=4)
        if not images: raise ValueError("Không tìm thấy trang nào trong PDF hoặc không thể chuyển đổi thành ảnh.")
        prs = Presentation(); prs = setup_slide_size(prs, input_path); blank_layout = prs.slide_layouts[6]
        image_files = sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))], key=lambda x: int(os.path.splitext(os.path.basename(x))[0].split('-')[-1]))
        if not image_files: raise ValueError("Không tìm thấy file ảnh nào trong thư mục tạm.")
        logger.info(f"Tìm thấy {len(image_files)} ảnh trang để thêm vào PPTX.")
        for image_path in image_files:
            try:
                with Image.open(image_path) as img: img_width, img_height = img.size
                slide = prs.slides.add_slide(blank_layout); img_ratio = img_width / img_height; slide_width_emu, slide_height_emu = prs.slide_width, prs.slide_height; slide_ratio = slide_width_emu / slide_height_emu
                if img_ratio > slide_ratio: pic_width, pic_height, pic_left, pic_top = slide_width_emu, int(pic_width / img_ratio), 0, int((slide_height_emu - pic_height) / 2)
                else: pic_height, pic_width, pic_left, pic_top = slide_height_emu, int(pic_height * img_ratio), int((slide_width_emu - pic_width) / 2), 0
                slide.shapes.add_picture(image_path, pic_left, pic_top, width=pic_width, height=pic_height)
            except Exception as page_err: logger.warning(f"Lỗi khi thêm ảnh {os.path.basename(image_path)} vào slide: {page_err}. Bỏ qua trang này.")
        prs.save(output_path); logger.info(f"Đã lưu PPTX thành công tại: {output_path}"); return True
    except ValueError as ve: logger.error(f"Lỗi giá trị khi chuyển đổi PDF sang PPTX (hình ảnh): {ve}", exc_info=True); raise
    except Exception as e: logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang PPTX (hình ảnh): {e}", exc_info=True); return False
    finally: safe_remove(temp_dir)

def convert_pdf_to_pptx_python(input_path, output_path):
    # ... (giữ nguyên) ...
    logger.info("Thử chuyển đổi PDF -> PPTX bằng phương pháp hình ảnh (Python)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)


# --- IMAGE CONVERSION FUNCTIONS (Renamed) ---
def convert_images_to_pdf(image_files, output_path):
    """Chuyển đổi một danh sách các file ảnh (FileStorage) thành một file PDF đa trang."""
    image_objects = []
    opened_streams = [] # Keep track of opened streams if needed (Pillow might handle this)
    try:
        for file_storage in image_files:
            try:
                # Ensure stream is at the beginning
                file_storage.stream.seek(0)
                img = Image.open(file_storage.stream)
                opened_streams.append(file_storage.stream) # Add stream if needed for later closing, though Pillow should manage

                # Convert to RGB for compatibility
                if img.mode == 'RGBA': bg = Image.new('RGB', img.size, (255, 255, 255)); bg.paste(img, (0, 0), img); img = bg
                elif img.mode not in ['RGB', 'L']: img = img.convert('RGB')

                # Crucially, append the *converted* Image object
                image_objects.append(img)
                # Do NOT close the original file_storage.stream here if Pillow needs it

            except UnidentifiedImageError:
                 logger.error(f"File không phải ảnh hợp lệ hoặc bị lỗi: {file_storage.filename}", exc_info=False) # Less verbose logging
                 raise ValueError(f"'{file_storage.filename}' is not a valid image file.") # User-friendly error
            except Exception as img_err:
                logger.error(f"Lỗi khi xử lý ảnh {file_storage.filename}: {img_err}", exc_info=True)
                raise # Re-throw other processing errors

        if not image_objects:
            logger.warning("Không có đối tượng ảnh nào được xử lý.")
            return False

        first_image = image_objects[0]
        other_images = image_objects[1:]

        # Save the first image, appending the rest
        first_image.save(
            output_path, "PDF", resolution=100.0,
            save_all=True, append_images=other_images
        )
        logger.info(f"Đã chuyển đổi {len(image_objects)} ảnh thành PDF thành công: {output_path}")
        return True

    except ValueError as ve: # Catch invalid image error
         logger.error(f"Lỗi giá trị khi chuyển đổi ảnh sang PDF: {ve}")
         raise # Re-throw to be caught by the route
    except Exception as e:
        logger.error(f"Lỗi nghiêm trọng khi chuyển đổi ảnh sang PDF: {e}", exc_info=True)
        return False
    finally:
        # Close all image objects to release resources
        for img in image_objects:
            try: img.close()
            except Exception: pass
        # Optionally, close streams if Pillow didn't
        # for stream in opened_streams:
        #     try: stream.close()
        #     except Exception: pass

def convert_pdf_to_image_zip(input_path, output_zip_path, img_format='jpeg'):
    """Chuyển đổi PDF sang nhiều file ảnh và nén thành ZIP."""
    temp_dir = None
    fmt = img_format.lower()
    if fmt not in ['jpeg', 'jpg', 'png']: fmt = 'jpeg' # Default to jpeg
    ext = 'jpg' if fmt == 'jpeg' else fmt

    try:
        temp_dir = tempfile.mkdtemp(prefix="pdf2imgzip_")
        logger.info(f"Tạo thư mục tạm cho ảnh ({ext}): {temp_dir}")

        try:
            pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None)
            page_count = pdfinfo.get('Pages', 0)
            logger.info(f"PDF info read successfully. Pages: {page_count}")
            if page_count == 0: logger.warning("PDF reported 0 pages.")
        except Exception as info_err:
            logger.error(f"Lỗi khi chạy pdfinfo (kiểm tra Poppler): {info_err}.", exc_info=True)
            raise ValueError("Could not process PDF, Poppler might be missing or not configured correctly.") from info_err

        output_base_name = os.path.splitext(os.path.basename(input_path))[0]
        images = convert_from_path(
            input_path, dpi=200, fmt=fmt, output_folder=temp_dir,
            output_file=output_base_name, thread_count=4
        )

        image_files = [f for f in os.listdir(temp_dir) if f.lower().endswith(f'.{ext}')]
        if not image_files: raise ValueError("Không tạo được ảnh nào từ PDF.")
        logger.info(f"Đã tạo {len(image_files)} file {ext.upper()} trong {temp_dir}")

        image_files_sorted = sorted(
            [os.path.join(temp_dir, f) for f in image_files],
            key=lambda x: int(os.path.splitext(os.path.basename(x))[0].split('-')[-1]) # Sort numerically
        )

        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for i, file_path in enumerate(image_files_sorted):
                arcname = f"page_{i+1}.{ext}" # Use correct extension in zip
                zipf.write(file_path, arcname=arcname)

        logger.info(f"Đã tạo file ZIP thành công: {output_zip_path}")
        return True
    except ValueError as ve: logger.error(f"Lỗi giá trị khi chuyển đổi PDF sang ảnh: {ve}", exc_info=True); raise
    except Exception as e: logger.error(f"Lỗi nghiêm trọng khi chuyển đổi PDF sang ảnh/ZIP: {e}", exc_info=True); return False
    finally: safe_remove(temp_dir)

# --- ROUTES ---

@app.route('/')
def index():
    translations_url = url_for('get_translations')
    return render_template('index.html', translations_url=translations_url)

@app.route('/convert', methods=['POST'])
def convert_file():
    """Handles PDF <-> DOCX and PDF <-> PPT conversions."""
    output_path = None
    temp_libreoffice_output = None
    input_path_for_process = None # Only one input file expected here
    saved_input_paths = []

    try:
        if 'file' not in request.files: return "No file part", 400
        file = request.files['file']
        if not file or file.filename == '': return "No selected file", 400

        filename = secure_filename(file.filename)
        if not _allowed_file(filename, ALLOWED_EXTENSIONS - ALLOWED_IMAGE_EXTENSIONS): # Check against non-image extensions
             ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'none'
             return f"Invalid file type '{ext}' for this converter. Use PDF, DOCX, PPT, PPTX.", 400

        actual_conversion_type = request.form.get('conversion_type')
        if not actual_conversion_type or actual_conversion_type not in ['pdf_to_docx', 'docx_to_pdf', 'pdf_to_ppt', 'ppt_to_pdf']:
            return "Invalid or missing conversion type for this endpoint.", 400

        logger.info(f"Yêu cầu chuyển đổi Office/PDF: file='{filename}', type='{actual_conversion_type}'")

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        input_path_for_process = os.path.join(UPLOAD_FOLDER, f"input_{time.time()}_{filename}")
        file.save(input_path_for_process)
        saved_input_paths.append(input_path_for_process)
        logger.info(f"File input đã lưu: {input_path_for_process}")

        base_name = filename.rsplit('.', 1)[0]
        if actual_conversion_type == 'pdf_to_docx': out_ext = 'docx'
        elif actual_conversion_type == 'docx_to_pdf': out_ext = 'pdf'
        elif actual_conversion_type == 'pdf_to_ppt': out_ext = 'pptx'
        elif actual_conversion_type == 'ppt_to_pdf': out_ext = 'pdf'
        # No else needed due to check above

        output_filename = f"converted_{time.time()}_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        logger.info(f"File output dự kiến: {output_path}")

        conversion_success = False
        error_message = "Conversion error"

        # --- Conversion Logic (Only PDF/Office) ---
        try:
            if actual_conversion_type == 'pdf_to_docx':
                # ... (pdf2docx logic) ...
                cv = Converter(input_path_for_process); cv.convert(output_path); cv.close()
                conversion_success = True; logger.info("PDF -> DOCX success.")
            elif actual_conversion_type == 'docx_to_pdf':
                # ... (LibreOffice logic) ...
                expected_lo_output_name = os.path.basename(input_path_for_process).replace('.docx', '.pdf')
                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name); safe_remove(temp_libreoffice_output)
                result = subprocess.run([SOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path_for_process], check=True, timeout=120, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0: os.rename(temp_libreoffice_output, output_path); conversion_success = True; logger.info(f"DOCX -> PDF success.")
                else: error_message = "LibreOffice didn't create output PDF."; logger.error(error_message)
            elif actual_conversion_type == 'pdf_to_ppt':
                # ... (Python/LibreOffice fallback logic) ...
                try:
                     if convert_pdf_to_pptx_python(input_path_for_process, output_path): conversion_success = True; logger.info("PDF -> PPTX (Python) success.")
                     else: raise RuntimeError("Python PPTX conversion failed, attempting LibreOffice fallback.")
                except (ValueError, RuntimeError, Exception) as py_ppt_err:
                     logger.warning(f"Python PPTX Error ({type(py_ppt_err).__name__}), trying LO fallback...")
                     expected_lo_output_name = os.path.basename(input_path_for_process).replace('.pdf', '.pptx')
                     temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name); safe_remove(temp_libreoffice_output)
                     try:
                         result = subprocess.run([SOFFICE_PATH, '--headless', '--convert-to', 'pptx', '--outdir', UPLOAD_FOLDER, input_path_for_process], check=True, timeout=180, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                         if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0: os.rename(temp_libreoffice_output, output_path); conversion_success = True; logger.info(f"PDF -> PPTX (LO fallback) success.")
                         else: error_message = "LibreOffice (fallback) didn't create output PPTX."; logger.error(error_message)
                     except Exception as lo_ppt_err: error_message = f"LO fallback Error: {lo_ppt_err}"; logger.error(error_message, exc_info=True)
                     if not conversion_success: error_message = "Both Python and LO failed for PDF -> PPTX."
            elif actual_conversion_type == 'ppt_to_pdf':
                # ... (LibreOffice logic) ...
                expected_lo_output_name = os.path.basename(input_path_for_process)
                if expected_lo_output_name.lower().endswith('.pptx'): expected_lo_output_name = expected_lo_output_name[:-5] + '.pdf'
                elif expected_lo_output_name.lower().endswith('.ppt'): expected_lo_output_name = expected_lo_output_name[:-4] + '.pdf'
                else: expected_lo_output_name += '.pdf'
                temp_libreoffice_output = os.path.join(UPLOAD_FOLDER, expected_lo_output_name); safe_remove(temp_libreoffice_output)
                result = subprocess.run([SOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path_for_process], check=True, timeout=120, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0: os.rename(temp_libreoffice_output, output_path); conversion_success = True; logger.info(f"PPT/PPTX -> PDF success.")
                else: error_message = "LibreOffice didn't create output PDF from PPT."; logger.error(error_message)

        # --- Catch specific conversion errors ---
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as sub_err: error_message = f"Process Error ({type(sub_err).__name__}): {sub_err}"; logger.error(f"Error processing {actual_conversion_type}: {error_message}", exc_info=True)
        except ValueError as val_err: error_message = f"Data/Config Error: {val_err}"; logger.error(f"Error processing {actual_conversion_type}: {error_message}", exc_info=True)
        except FileNotFoundError as fnf_err: error_message = f"File/Program Not Found: {fnf_err}"; logger.error(f"Error processing {actual_conversion_type}: {error_message}", exc_info=True)
        except Exception as conv_err: error_message = f"Unknown conversion error: {conv_err}"; logger.error(error_message, exc_info=True)

        # --- Handling result ---
        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            try:
                # ... (send_file logic with cleanup) ...
                mimetype = None
                if out_ext == 'pdf': mimetype = 'application/pdf'
                elif out_ext == 'docx': mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                elif out_ext == 'pptx': mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                @response.call_on_close
                def cleanup_files_after_send():
                    logger.info("Cleanup after sending file...")
                    for p in saved_input_paths: safe_remove(p)
                    safe_remove(output_path)
                    if temp_libreoffice_output and temp_libreoffice_output != output_path: safe_remove(temp_libreoffice_output)
                return response
            except Exception as send_err: logger.error(f"Error sending file {output_path}: {send_err}", exc_info=True); error_message = f"Error preparing download: {send_err}"; conversion_success = False

        # --- Failure Case ---
        if not conversion_success:
            logger.error(f"Conversion failed for {actual_conversion_type}. Reason: {error_message}")
            for p in saved_input_paths: safe_remove(p)
            safe_remove(output_path)
            if temp_libreoffice_output: safe_remove(temp_libreoffice_output)
            return f"Conversion failed: {error_message}", 500

    except Exception as e:
        logger.error(f"Unexpected error in /convert route: {e}", exc_info=True)
        for p in saved_input_paths: safe_remove(p)
        safe_remove(output_path)
        if temp_libreoffice_output: safe_remove(temp_libreoffice_output)
        return f"Unexpected server error: {str(e)}", 500


# --- NEW ROUTE FOR PDF <-> IMAGE ---
@app.route('/convert_image', methods=['POST'])
def convert_image_route():
    """Handles PDF -> Image (ZIP) and Image(s) -> PDF conversions."""
    output_path = None
    input_path_for_process = None # Used only for PDF -> Images
    saved_input_paths = [] # Track saved PDF input

    try:
        uploaded_files = request.files.getlist('image_file') # Use the new input name

        if not uploaded_files or not uploaded_files[0].filename:
            return "No file selected.", 400 # Use key 'err-select-file'

        # --- Check total size ---
        total_size = sum(f.content_length for f in uploaded_files if f.content_length is not None)
        if total_size > app.config['MAX_CONTENT_LENGTH']:
             return "Total file size exceeds limit.", 413 # Use key 'err-file-too-large'

        first_file = uploaded_files[0]
        first_filename = secure_filename(first_file.filename)
        first_ext = first_filename.rsplit('.', 1)[-1].lower() if '.' in first_filename else ''

        # --- Determine conversion type and validate ---
        actual_conversion_type = None
        if first_ext == 'pdf':
            if len(uploaded_files) > 1:
                 return "Please select only one PDF file to convert to images.", 400 # Use key 'err-image-single-pdf'
            if not _allowed_file(first_filename, ALLOWED_IMAGE_EXTENSIONS):
                 return f"Invalid file type '{first_ext}'. Select PDF, JPG, or JPEG.", 400 # Use key 'err-image-format'
            actual_conversion_type = 'pdf_to_image'
            out_ext = 'zip'
        elif first_ext in ['jpg', 'jpeg']:
            # Allow multiple JPG/JPEG files
            for f in uploaded_files:
                 fname = secure_filename(f.filename)
                 if not _allowed_file(fname, ALLOWED_IMAGE_EXTENSIONS) or fname.rsplit('.', 1)[-1].lower() not in ['jpg', 'jpeg']:
                      return "If selecting multiple files, all must be JPG or JPEG.", 400 # Use key 'err-image-all-images'
            actual_conversion_type = 'image_to_pdf'
            out_ext = 'pdf'
        else:
            return f"Invalid file type '{first_ext}'. Select PDF, JPG, or JPEG.", 400 # Use key 'err-image-format'

        logger.info(f"Yêu cầu chuyển đổi Ảnh/PDF: {len(uploaded_files)} file(s), file đầu tiên='{first_filename}', type='{actual_conversion_type}'")

        # --- Save input PDF if needed ---
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        if actual_conversion_type == 'pdf_to_image':
            input_path_for_process = os.path.join(UPLOAD_FOLDER, f"input_{time.time()}_{first_filename}")
            first_file.save(input_path_for_process)
            saved_input_paths.append(input_path_for_process)
            logger.info(f"Input PDF saved: {input_path_for_process}")

        # --- Define Output Path ---
        base_name = first_filename.rsplit('.', 1)[0]
        output_filename = f"converted_{time.time()}_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        logger.info(f"Image conversion output dự kiến: {output_path}")

        conversion_success = False
        error_message = "Image conversion error"

        # --- Conversion Logic ---
        try:
            if actual_conversion_type == 'pdf_to_image':
                # Use the renamed function
                if convert_pdf_to_image_zip(input_path_for_process, output_path):
                    conversion_success = True
                    logger.info("PDF -> Images (ZIP) success.")
                else:
                    error_message = "Failed to convert PDF to images." # Function logs details

            elif actual_conversion_type == 'image_to_pdf':
                # Use the renamed function, passing the list of FileStorage objects
                if convert_images_to_pdf(uploaded_files, output_path):
                    conversion_success = True
                    logger.info("Images -> PDF success.")
                else:
                    error_message = "Failed to convert images to PDF." # Function logs details

        except ValueError as val_err: # Catch Poppler/Invalid Image errors
            error_message = str(val_err)
            logger.error(f"Error processing {actual_conversion_type}: {error_message}", exc_info=False) # Less verbose stack trace for user errors
        except Exception as conv_err:
            error_message = f"Unknown conversion error: {conv_err}"
            logger.error(error_message, exc_info=True)

        # --- Handling result ---
        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            try:
                # ... (send_file logic with cleanup) ...
                mimetype = 'application/zip' if out_ext == 'zip' else 'application/pdf'
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                @response.call_on_close
                def cleanup_files_after_send():
                    logger.info("Cleanup after sending image conversion file...")
                    for p in saved_input_paths: safe_remove(p) # Only PDF input is saved
                    safe_remove(output_path)
                return response
            except Exception as send_err:
                logger.error(f"Error sending file {output_path}: {send_err}", exc_info=True)
                error_message = f"Error preparing download: {send_err}"
                conversion_success = False

        # --- Failure Case ---
        if not conversion_success:
            logger.error(f"Image conversion failed for {actual_conversion_type}. Reason: {error_message}")
            for p in saved_input_paths: safe_remove(p)
            safe_remove(output_path)
            # Return error using the message from the exception or conversion logic
            return f"Conversion failed: {error_message}", 500

    except Exception as e:
        # Catch-all for unexpected errors in the route handler itself
        logger.error(f"Unexpected error in /convert_image route: {e}", exc_info=True)
        for p in saved_input_paths: safe_remove(p)
        safe_remove(output_path)
        return f"Unexpected server error: {str(e)}", 500


# --- Teardown and Main ---
@app.teardown_appcontext
def cleanup_old_files(exception=None):
     # ... (giữ nguyên) ...
    if not os.path.exists(UPLOAD_FOLDER): return
    logger.info("Chạy dọn dẹp file cũ teardown_appcontext...")
    try:
        now = time.time(); max_age = 3600; deleted_count = 0; checked_count = 0
        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                stat_result = os.lstat(path); is_file = os.path.isfile(path); is_dir = os.path.isdir(path)
                if is_file or is_dir:
                    checked_count += 1; file_age = now - stat_result.st_mtime
                    if file_age > max_age:
                        logger.info(f"Teardown: Xóa {'file' if is_file else 'dir'} cũ ({file_age:.0f}s > {max_age}s): {path}")
                        if safe_remove(path): deleted_count += 1
            except FileNotFoundError: continue
            except Exception as e: logger.error(f"Lỗi khi kiểm tra/dọn dẹp {path}: {e}")
        logger.info(f"Teardown dọn dẹp hoàn tất. Đã kiểm tra {checked_count} mục, xóa {deleted_count} mục cũ.")
    except Exception as e: logger.error(f"Lỗi nghiêm trọng trong teardown_appcontext: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    logger.info(f"Thư mục Upload: {UPLOAD_FOLDER}")
    logger.info(f"Khởi động server trên cổng {port} - Chế độ Debug: {debug_mode}")
    app.run(host='0.0.0.0', port=port, debug=debug_mode, threaded=True)

# --- END OF FILE app.py ---