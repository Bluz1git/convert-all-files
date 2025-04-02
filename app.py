# --- START OF FILE app.py ---

from flask import Flask, request, send_file, render_template, jsonify, url_for, make_response
from flask_talisman import Talisman # Import Talisman
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
from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError, PDFSyntaxError
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image, UnidentifiedImageError
import zipfile
import json # <--- THÊM IMPORT JSON

# === Basic Flask App Setup ===
app = Flask(__name__, template_folder='templates', static_folder='static')

# === Configuration ===
# Increased max content length slightly for potential overhead
app.config['MAX_CONTENT_LENGTH'] = 105 * 1024 * 1024  # ~105MB limit server-side
# Secret key for session management (needed by some extensions, good practice)
# Use an environment variable in production!
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')

# === Logging ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s', # Added funcName
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# === Security Headers with Talisman ===
# Content Security Policy (CSP) - Adjust as needed, start restrictive
# Allows loading resources only from self (your domain) and the specified CDN/fonts
csp = {
    'default-src': [
        '\'self\'',
        'https://cdn.tailwindcss.com',
        'https://fonts.googleapis.com',
        'https://fonts.gstatic.com'
    ],
    'style-src': [
        '\'self\'',
        '\'unsafe-inline\'', # Allow inline styles (Tailwind needs this sometimes, or configure properly)
        'https://cdn.tailwindcss.com',
        'https://fonts.googleapis.com'
    ],
     'script-src': [
        '\'self\'',
        '\'unsafe-inline\'', # Allow inline scripts (Your JS is inline) - Consider moving JS to separate files later
        'https://cdn.tailwindcss.com',
    ],
    'font-src': [
        '\'self\'',
        'https://fonts.gstatic.com'
    ],
    'img-src': [
         '\'self\'',
         'data:' # Allow data URIs if needed for images
    ]
}
# Initialize Talisman
# force_https=False because Railway (or other reverse proxies) usually handle this redirection
# Use session_cookie_secure=True if you deploy with HTTPS (highly recommended)
talisman = Talisman(
    app,
    content_security_policy=csp,
    force_https=False, # Railway/proxy handles HTTPS redirect
    session_cookie_secure=True, # Assumes HTTPS deployment
    session_cookie_http_only=True, # Đã sửa đúng
    frame_options='DENY', # Prevent clickjacking
    strict_transport_security=True, # Enable HSTS if using HTTPS
    content_security_policy_nonce_in=['script-src'] # Optional: for CSP nonces if needed later
)


# === Constants ===
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
# Only allow these extensions overall
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}
# Extensions specifically for the PDF/Image card
ALLOWED_IMAGE_EXTENSIONS = {'pdf', 'jpg', 'jpeg'}
LIBREOFFICE_TIMEOUT = 180 # Seconds
TRANSLATIONS_FILE = 'translations.json' # Đường dẫn tới file JSON

# === Helper Functions ===

def make_error_response(error_key, status_code=400):
    """Creates a Flask response with an error message prefixed for JS handling."""
    logger.warning(f"Returning error: {error_key} (Status: {status_code})")
    # Prefix is crucial for JS error parsing in index.html
    response_text = f"Conversion failed: {error_key}"
    response = make_response(response_text, status_code)
    response.headers["Content-Type"] = "text/plain; charset=utf-8" # Thêm charset=utf-8
    return response

def find_libreoffice():
    """Tries to find the LibreOffice executable."""
    possible_paths = [
        'soffice', # Check PATH first
        '/usr/bin/soffice', '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice', '/usr/lib/libreoffice/program/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    ]
    for path in possible_paths:
        try:
            resolved_path = shutil.which(path)
            if resolved_path:
                result = subprocess.run([resolved_path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                if result.returncode == 0 and 'LibreOffice' in result.stdout:
                    logger.info(f"Found usable LibreOffice via shutil.which: {resolved_path}")
                    return resolved_path
            elif os.path.isfile(path): # Check direct path if not in PATH
                 result = subprocess.run([path, '--version'], capture_output=True, text=True, check=False, timeout=5)
                 if result.returncode == 0 and 'LibreOffice' in result.stdout:
                    logger.info(f"Found usable LibreOffice via direct path check: {path}")
                    return path
        except FileNotFoundError: logger.debug(f"LibreOffice not found at or via {path}")
        except subprocess.TimeoutExpired: logger.warning(f"Checking LibreOffice at {path} timed out.")
        except Exception as e: logger.warning(f"Error checking LibreOffice at {path}: {e}")

    logger.error("LibreOffice executable not found or verification failed. Office conversions might fail.")
    return None # Return None if not found

SOFFICE_PATH = find_libreoffice()
if SOFFICE_PATH:
    logger.info(f"Using LibreOffice path: {SOFFICE_PATH}")
else:
    logger.warning("LibreOffice not found. DOCX/PPTX related conversions will likely fail.")

def _allowed_file(filename, allowed_set):
    """Checks if the file extension is in the allowed set."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def safe_remove(item_path, retries=3, delay=0.5):
    """Safely removes a file or directory with retries."""
    if not item_path or not os.path.exists(item_path): return True
    is_dir = os.path.isdir(item_path)
    item_type = "directory" if is_dir else "file"
    for i in range(retries):
        try:
            if is_dir: shutil.rmtree(item_path)
            else: os.remove(item_path)
            logger.debug(f"Successfully removed {item_type}: {item_path}")
            return True
        except PermissionError as e: logger.warning(f"Permission error removing {item_path} (Attempt {i + 1}/{retries}): {e}")
        except OSError as e: logger.warning(f"OS error removing {item_path} (Attempt {i + 1}/{retries}): {e}")
        except Exception as e: logger.warning(f"Unexpected error removing {item_path} (Attempt {i + 1}/{retries}): {e}")
        if i < retries - 1:
            logger.debug(f"Retrying removal of {item_path} after {delay * (i + 1):.1f}s delay...")
            time.sleep(delay * (i + 1))
    logger.error(f"Failed to remove {item_type} after {retries} attempts: {item_path}")
    return False

# --- PDF Processing Helpers ---
def get_pdf_page_size(pdf_path):
    """Gets the dimensions (width, height) of the first page of a PDF in points."""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if reader.is_encrypted:
                 logger.warning(f"PDF is encrypted: {pdf_path}")
                 raise ValueError("err-pdf-corrupt") # Treat encrypted as corrupt/unsupported
            if not reader.pages:
                logger.warning(f"PDF has no pages: {pdf_path}")
                return None, None # Or raise error? Let's return None for now.

            page = reader.pages[0]
            # Use mediabox as primary, fallback to cropbox if needed
            box = page.mediabox or page.cropbox
            if box:
                width = float(box.width)
                height = float(box.height)
                if width > 0 and height > 0: return width, height
                else: logger.warning(f"Invalid page dimensions (<=0) in {pdf_path}."); return None, None
            else: logger.warning(f"Could not find page dimensions (mediabox/cropbox) in {pdf_path}."); return None, None
    except PyPDF2.errors.PdfReadError as pdf_err:
         logger.error(f"PyPDF2 read error (possibly corrupt PDF): {pdf_path} - {pdf_err}", exc_info=False)
         raise ValueError("err-pdf-corrupt") from pdf_err
    except ValueError as ve: raise ve # Re-raise specific errors
    except Exception as e:
        logger.error(f"Unexpected error reading PDF page size {pdf_path}: {e}", exc_info=True)
        raise ValueError("err-pdf-corrupt") from e

def setup_slide_size(prs, pdf_path):
    """Sets the slide dimensions in the Presentation object based on the PDF page size."""
    pdf_width_pt, pdf_height_pt = None, None
    try:
        pdf_width_pt, pdf_height_pt = get_pdf_page_size(pdf_path)
    except ValueError as e:
        logger.warning(f"Could not get PDF page size for slide setup ({e}), using default.")
        # Fallthrough to default below

    if pdf_width_pt is None or pdf_height_pt is None:
        logger.warning("Using default slide size (10x7.5 inches).")
        prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
        return prs

    try:
        # Convert points to inches (1 inch = 72 points)
        pdf_width_in, pdf_height_in = pdf_width_pt / 72.0, pdf_height_pt / 72.0
        # PowerPoint max slide dimension is 56 inches
        max_slide_dim_in = 56.0
        if pdf_width_in > max_slide_dim_in or pdf_height_in > max_slide_dim_in:
            logger.info(f"PDF dimensions ({pdf_width_in:.2f}x{pdf_height_in:.2f} in) exceed max slide size ({max_slide_dim_in} in). Scaling down.")
            ratio = pdf_width_in / pdf_height_in
            # Scale based on the larger dimension hitting the max limit
            if pdf_width_in >= pdf_height_in:
                 final_width_in = max_slide_dim_in
                 final_height_in = max_slide_dim_in / ratio
            else: # pdf_height_in > pdf_width_in
                 final_height_in = max_slide_dim_in
                 final_width_in = max_slide_dim_in * ratio
            logger.info(f"Scaled slide dimensions to: {final_width_in:.2f}x{final_height_in:.2f} in")
        else:
            final_width_in, final_height_in = pdf_width_in, pdf_height_in
            logger.info(f"Using original PDF dimensions for slide: {final_width_in:.2f}x{final_height_in:.2f} in")

        # Set slide dimensions in Inches
        prs.slide_width = Inches(final_width_in)
        prs.slide_height = Inches(final_height_in)
        return prs
    except Exception as e:
        logger.warning(f"Error applying PDF dimensions to slide, using default: {e}", exc_info=True)
        prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
        return prs

# --- START: Helper function for sorting PPTX images ---
def sort_key_for_pptx_images(filename):
    """Helper function to extract page number for sorting images from pdf2image."""
    try:
        # Lấy phần tên file không có extension
        base = os.path.splitext(filename)[0]
        # Tìm số ở cuối, có thể sau dấu '-' hoặc '_'
        num_str = base.split('-')[-1].split('_')[-1]
        return int(num_str)
    except (ValueError, IndexError):
        logger.warning(f"Could not extract page number from filename: {filename}. Using 0 for sorting.")
        return 0 # Trả về 0 để xếp vào đầu hoặc cuối
# --- END: Helper function for sorting PPTX images ---

def _convert_pdf_to_pptx_images(input_path, output_path):
    """Internal function to convert PDF to images and add them to PPTX slides."""
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="pdfimg_")
        logger.info(f"Created temp directory for PPTX images: {temp_dir}")
        page_count = 0 # Initialize page count

        # 1. Check Poppler & Get Page Count
        try:
            # Specify poppler_path=None to use system PATH or default location
            pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None)
            page_count = pdfinfo.get('Pages')
            if page_count is None:
                 # Try to raise a more specific error if poppler likely missing
                 raise PDFInfoNotInstalledError("Poppler might be missing or not in PATH.")
            logger.info(f"PDF info read successfully. Pages: {page_count}")
            if page_count == 0:
                 logger.warning("PDF has 0 pages. Creating empty PPTX.")
                 Presentation().save(output_path)
                 return True
        except (PDFInfoNotInstalledError, FileNotFoundError) as e:
            logger.error(f"Poppler error accessing {input_path}: {e}", exc_info=True)
            raise ValueError("err-poppler-missing") from e
        except (PDFPageCountError, PDFSyntaxError) as e:
            logger.error(f"PDF reading error (corrupt?) {input_path}: {e}", exc_info=False)
            raise ValueError("err-pdf-corrupt") from e
        except Exception as info_err: # Catch other potential poppler errors
            logger.error(f"Unexpected error getting PDF info {input_path}: {info_err}", exc_info=True)
            raise ValueError("err-poppler-missing") from info_err

        # 2. Convert PDF pages to images
        logger.info(f"Converting PDF pages to JPEG images (dpi=300)...")
        # Again specify poppler_path=None
        images = convert_from_path(input_path, dpi=300, fmt='jpeg', output_folder=temp_dir, thread_count=4, poppler_path=None)
        if not images:
             if page_count > 0: # If info said pages > 0 but no images generated
                 logger.error(f"pdf2image conversion returned no images for {input_path} despite page count {page_count}.")
                 raise RuntimeError("err-conversion")
             else: # 0 pages case already handled
                 logger.info("No images generated (PDF likely empty). Empty PPTX created.")
                 return True

        # 3. Create Presentation and add images
        prs = Presentation()
        prs = setup_slide_size(prs, input_path) # Set size *before* adding slides
        blank_layout = prs.slide_layouts[6] # Use a blank slide layout

        # Find and sort generated images
        generated_images = sorted(
            [f for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg'))],
            key=sort_key_for_pptx_images
        )
        if not generated_images:
            logger.error(f"No JPEG images found in temp dir {temp_dir} after conversion.")
            raise RuntimeError("err-conversion") # Or maybe a more specific error
        logger.info(f"Found {len(generated_images)} images to add to PPTX.")

        # Add images to slides, fitting them to the slide dimensions
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        for image_filename in generated_images:
            image_path = os.path.join(temp_dir, image_filename)
            try:
                slide = prs.slides.add_slide(blank_layout)
                # Add picture centered and scaled to fit (maintaining aspect ratio)
                # This is a simple fit, might leave margins if aspect ratios differ
                img = Image.open(image_path)
                img_width_px, img_height_px = img.size
                img.close()

                # Calculate scaling factor to fit within slide dimensions
                # Use EMUs directly for potentially better precision within pptx library
                img_width_emu = Inches(img_width_px / 96.0) # Approximate conversion assuming 96 DPI source for calculation
                img_height_emu = Inches(img_height_px / 96.0)

                width_ratio = slide_width / img_width_emu if img_width_emu > 0 else 1
                height_ratio = slide_height / img_height_emu if img_height_emu > 0 else 1
                scale_ratio = min(width_ratio, height_ratio)

                pic_width = int(img_width_emu * scale_ratio)
                pic_height = int(img_height_emu * scale_ratio)

                # Center the image using EMUs
                pic_left = int((slide_width - pic_width) / 2)
                pic_top = int((slide_height - pic_height) / 2)

                # Ensure non-negative dimensions before adding
                if pic_width > 0 and pic_height > 0:
                     slide.shapes.add_picture(image_path, pic_left, pic_top, width=pic_width, height=pic_height)
                else:
                     logger.warning(f"Calculated invalid dimensions ({pic_width}x{pic_height}) for image {image_filename}. Skipping.")

            except UnidentifiedImageError:
                 logger.warning(f"Skipping invalid image file: {image_filename}")
            except Exception as page_err:
                 logger.warning(f"Error adding image {image_filename} to slide: {page_err}. Skipping.")

        # 4. Save presentation
        prs.save(output_path)
        logger.info(f"Saved PPTX with {len(prs.slides)} slides to: {output_path}")
        return True

    except ValueError as ve: raise ve # Re-raise specific errors like err-poppler-missing
    except RuntimeError as rt_err: raise rt_err # Re-raise err-conversion
    except Exception as e:
        logger.error(f"Unexpected error during PDF->PPTX conversion: {e}", exc_info=True)
        raise RuntimeError("err-unknown") from e # Generic error for unexpected issues
    finally:
        if temp_dir: safe_remove(temp_dir)


def convert_pdf_to_pptx_python(input_path, output_path):
    """Wrapper for PDF -> PPTX conversion using Python image method."""
    logger.info("Attempting PDF -> PPTX via Python (image-based)...")
    return _convert_pdf_to_pptx_images(input_path, output_path)

# --- Image Conversion Helpers ---
def convert_images_to_pdf(image_files, output_path):
    """Converts a list of image FileStorage objects into a single multi-page PDF."""
    image_objects = []
    try:
        # Sort by filename provided by the user/browser
        sorted_files = sorted(image_files, key=lambda f: secure_filename(f.filename))
        logger.info(f"Processing {len(sorted_files)} images for PDF conversion.")

        for file_storage in sorted_files:
            filename = secure_filename(file_storage.filename)
            try:
                # Read stream into BytesIO for PIL
                file_storage.stream.seek(0)
                img_io = BytesIO(file_storage.stream.read())

                with Image.open(img_io) as img: # Use 'with' to ensure closing
                    img.load() # Load image data to catch errors early

                    # Convert modes incompatible with PDF saving (RGBA, P, LA, etc.) to RGB
                    converted_img = None
                    if img.mode in ['RGBA', 'LA']:
                        logger.debug(f"Converting image {filename} from {img.mode} to RGB with white background.")
                        # Create a white background image
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        # Paste the image onto the background using the alpha channel as mask
                        try:
                            mask = img.getchannel('A') if img.mode == 'RGBA' else (img.getchannel('L') if img.mode == 'LA' else None)
                            bg.paste(img, mask=mask)
                            converted_img = bg
                        except ValueError: # Handle cases where alpha might not be separable easily
                            logger.warning(f"Could not use alpha mask for {filename}, converting to RGB directly.")
                            converted_img = img.convert('RGB')

                    elif img.mode == 'P': # Palette mode - convert to RGB
                        logger.debug(f"Converting image {filename} from P to RGB.")
                        converted_img = img.convert('RGB')
                    elif img.mode not in ['RGB', 'L', 'CMYK']: # Allow RGB, Grayscale, CMYK
                        logger.debug(f"Converting image {filename} from {img.mode} to RGB.")
                        converted_img = img.convert('RGB')
                    else:
                        # If already compatible, create a copy to avoid issues with save_all
                        # modifying the original object potentially needed later (though less likely here)
                        converted_img = img.copy()

                    image_objects.append(converted_img)

            except UnidentifiedImageError:
                logger.error(f"Cannot identify image file: {filename}", exc_info=False)
                raise ValueError("err-invalid-image-file")
            except OSError as img_os_err:
                logger.error(f"OS error processing image file {filename}: {img_os_err}", exc_info=False)
                raise ValueError("err-invalid-image-file") from img_os_err
            except Exception as img_err:
                logger.error(f"Unexpected error processing image {filename}: {img_err}", exc_info=True)
                raise RuntimeError("err-conversion") from img_err

        if not image_objects:
            logger.warning("No valid images found to convert to PDF.")
            raise ValueError("err-select-file") # Or a more specific error

        logger.info(f"Saving {len(image_objects)} images to PDF: {output_path}")
        # Use the first image object's save method with save_all=True
        image_objects[0].save(
            output_path,
            "PDF",
            resolution=100.0, # Standard resolution
            save_all=True,    # Crucial for multi-page PDF
            append_images=image_objects[1:] # List of subsequent image objects
        )
        logger.info(f"Successfully converted images to PDF: {output_path}")
        return True
    except ValueError as ve: raise ve
    except Exception as e:
        logger.error(f"Unexpected error converting images to PDF: {e}", exc_info=True)
        raise RuntimeError("err-unknown") from e
    finally:
        # Ensure all PIL Image objects in the list are closed
        for img_obj in image_objects:
             try: img_obj.close()
             except Exception: pass

def convert_pdf_to_image_zip(input_path, output_zip_path, img_format='jpeg'):
    """Converts PDF pages to images and creates a ZIP archive."""
    temp_dir = None
    fmt = img_format.lower()
    if fmt not in ['jpeg', 'jpg', 'png']: fmt = 'jpeg'
    ext = 'jpg' if fmt == 'jpeg' else fmt

    try:
        temp_dir = tempfile.mkdtemp(prefix="pdf2imgzip_")
        logger.info(f"Created temp dir for PDF->Image ({ext}): {temp_dir}")
        page_count = 0
        # 1. Check Poppler & Get Page Count
        try:
            pdfinfo = pdfinfo_from_path(input_path, userpw=None, poppler_path=None)
            page_count = pdfinfo.get('Pages')
            if page_count is None: raise PDFInfoNotInstalledError("Poppler might be missing.")
            logger.info(f"PDF info read. Pages: {page_count}")
            if page_count == 0:
                 logger.warning("PDF has 0 pages. Creating empty ZIP.")
                 with zipfile.ZipFile(output_zip_path, 'w') as zipf: pass # Create empty zip
                 return True
        except (PDFInfoNotInstalledError, FileNotFoundError) as e:
            logger.error(f"Poppler error accessing {input_path}: {e}", exc_info=True)
            raise ValueError("err-poppler-missing") from e
        except (PDFPageCountError, PDFSyntaxError) as e:
            logger.error(f"PDF reading error (corrupt?) {input_path}: {e}", exc_info=False)
            raise ValueError("err-pdf-corrupt") from e
        except Exception as info_err:
            logger.error(f"Unexpected error getting PDF info {input_path}: {info_err}", exc_info=True)
            raise ValueError("err-poppler-missing") from info_err

        # 2. Convert PDF pages to images
        # Use a safer base name for output files within the temp dir
        safe_output_base = secure_filename(f"page_{os.path.splitext(os.path.basename(input_path))[0]}")
        logger.info(f"Converting PDF pages to {ext.upper()} (dpi=200)...")
        # Ensure output_file uses the base name, pdf2image adds numbers/ext
        images = convert_from_path(
            input_path,
            dpi=200,
            fmt=fmt,
            output_folder=temp_dir,
            output_file=safe_output_base, # Base name for generated files
            thread_count=4,
            poppler_path=None
        )
        if not images:
             if page_count > 0:
                 logger.error(f"pdf2image conversion returned no images for {input_path} despite page count {page_count}.")
                 raise RuntimeError("err-conversion")
             else: # 0 pages case handled above
                 logger.info("No images generated (0 pages). Empty ZIP created.")
                 return True

        # --- START: Sorting function specific to image naming used by pdf2image ---
        def sort_key_for_zip_images(filename):
            """Extract page number based on pdf2image naming (e.g., base-XXXX.ext)."""
            try:
                # Assuming format like 'safe_base-XXXX.ext' where XXXX is page number
                # Split by '-' and take the last part before the extension
                page_num_str = os.path.splitext(filename)[0].split('-')[-1]
                return int(page_num_str)
            except (ValueError, IndexError, TypeError):
                logger.warning(f"Could not extract page number from filename: {filename} for ZIP sorting. Using 0.")
                return 0
        # --- END: Sorting function ---

        # 3. Find and sort generated images based on expected naming
        generated_files = sorted(
             [f for f in os.listdir(temp_dir) if f.lower().startswith(safe_output_base.lower()) and f.lower().endswith(f'.{ext}')],
             key=sort_key_for_zip_images # Use specific sorting key
        )
        if not generated_files:
             logger.error(f"No {ext.upper()} images found in temp dir {temp_dir} matching base '{safe_output_base}'.")
             raise RuntimeError("err-conversion")
        logger.info(f"Generated {len(generated_files)} {ext.upper()} files.")

        # 4. Create ZIP archive
        logger.info(f"Creating ZIP archive: {output_zip_path}")
        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for i, filename in enumerate(generated_files):
                file_path = os.path.join(temp_dir, filename)
                # Use a consistent naming scheme within the ZIP (page_1.ext, page_2.ext)
                arcname = f"page_{i+1}.{ext}"
                zipf.write(file_path, arcname=arcname)
        logger.info(f"Created ZIP file: {output_zip_path}")
        return True

    except ValueError as ve: raise ve # err-poppler-missing, err-pdf-corrupt
    except RuntimeError as rt_err: raise rt_err # err-conversion
    except Exception as e:
        logger.error(f"Unexpected error converting PDF to image ZIP: {e}", exc_info=True)
        raise RuntimeError("err-unknown") from e
    finally:
        if temp_dir: safe_remove(temp_dir)

# === Routes ===

# === THÊM ROUTE get_translations ===
@app.route('/api/translations')
def get_translations():
    """Provides translation strings to the frontend."""
    lang = request.args.get('lang', 'en') # Default to English
    try:
        # Ensure the translations file exists
        if not os.path.exists(TRANSLATIONS_FILE):
            logger.error(f"{TRANSLATIONS_FILE} not found!")
            # Return minimal English fallback if file missing
            return jsonify({
                "en": { "err-fetch-translations": "Could not load language data (file missing)." }
            }).get('en', {})

        with open(TRANSLATIONS_FILE, 'r', encoding='utf-8') as f:
            all_translations = json.load(f)

        # Return the requested language dictionary, fallback to English, then empty
        return jsonify(all_translations.get(lang, all_translations.get('en', {})))

    except json.JSONDecodeError:
        logger.error(f"Error decoding JSON from {TRANSLATIONS_FILE}", exc_info=True)
        return jsonify({"error": "Translation file is corrupted"}), 500
    except Exception as e:
        logger.error(f"Error loading translations: {e}", exc_info=True)
        return jsonify({"error": "Could not load translations"}), 500
# === KẾT THÚC ROUTE get_translations ===


@app.route('/')
def index():
    """Renders the main page."""
    try:
        # url_for will now work because get_translations function/endpoint exists
        translations_url = url_for('get_translations', _external=True)
        return render_template('index.html', translations_url=translations_url)
    except Exception as e:
        # Fallback in case url_for still fails for some reason or template render fails
        logger.error(f"Error rendering index page: {e}", exc_info=True)
        # You might want a simple static error page here
        return "An error occurred loading the page.", 500

# === PDF / Office Conversion Route ===
@app.route('/convert', methods=['POST'])
def convert_file():
    """Handles PDF <-> DOCX and PDF <-> PPT conversions."""
    output_path = temp_libreoffice_output = input_path_for_process = None
    saved_input_paths = [] # Store paths of successfully saved input files
    actual_conversion_type = None
    start_time = time.time()
    error_key = "err-conversion" # Initialize error key outside try-except for broader scope
    conversion_success = False   # Initialize success flag

    try:
        # 1. Validation
        if 'file' not in request.files:
            return make_error_response("err-select-file", 400)
        file = request.files['file']
        if not file or not file.filename:
            return make_error_response("err-select-file", 400)

        filename = secure_filename(file.filename)
        allowed_office_ext = {'pdf', 'docx', 'ppt', 'pptx'}
        file_ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
        if not _allowed_file(filename, allowed_office_ext):
             logger.warning(f"Disallowed file type uploaded to /convert: {filename}")
             return make_error_response("err-format-docx", 400) # Generic error

        actual_conversion_type = request.form.get('conversion_type')
        valid_conversion_types = ['pdf_to_docx', 'docx_to_pdf', 'pdf_to_ppt', 'ppt_to_pdf']
        if not actual_conversion_type or actual_conversion_type not in valid_conversion_types:
             logger.warning(f"Invalid conversion_type received: {actual_conversion_type}")
             return make_error_response("err-select-conversion", 400)

        # Cross-validation
        error_key_cv = None
        if actual_conversion_type == 'pdf_to_docx' and file_ext != 'pdf': error_key_cv = "err-format-docx"
        elif actual_conversion_type == 'docx_to_pdf' and file_ext != 'docx': error_key_cv = "err-format-docx"
        elif actual_conversion_type == 'pdf_to_ppt' and file_ext != 'pdf': error_key_cv = "err-format-ppt"
        elif actual_conversion_type == 'ppt_to_pdf' and file_ext not in ['ppt', 'pptx']: error_key_cv = "err-format-ppt"
        if error_key_cv:
            logger.warning(f"File type '{file_ext}' mismatch for conversion '{actual_conversion_type}'")
            return make_error_response(error_key_cv, 400)

        logger.info(f"Request /convert: file='{filename}', type='{actual_conversion_type}'")

        # 2. Size Check
        try:
            file.stream.seek(0, os.SEEK_END)
            file_size = file.stream.tell()
            file.stream.seek(0)
            if file_size > 100 * 1024 * 1024: # 100 MiB
                logger.warning(f"File too large: {filename} ({file_size} bytes)")
                return make_error_response("err-file-too-large", 413)
        except Exception as size_err:
            logger.error(f"Could not determine file size for {filename}: {size_err}", exc_info=True)
            return make_error_response("err-unknown", 500)

        # 3. Save Input & Define Output
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        input_filename_ts = f"input_{timestamp}_{filename}"
        input_path_for_process = os.path.join(UPLOAD_FOLDER, input_filename_ts)
        try:
            file.save(input_path_for_process)
            saved_input_paths.append(input_path_for_process)
            logger.info(f"Input saved: {input_path_for_process}")
        except Exception as save_err:
            logger.error(f"Failed to save uploaded file {filename}: {save_err}", exc_info=True)
            return make_error_response("err-unknown", 500)

        base_name = filename.rsplit('.', 1)[0]
        out_ext_map = {'pdf_to_docx': 'docx', 'docx_to_pdf': 'pdf', 'pdf_to_ppt': 'pptx', 'ppt_to_pdf': 'pdf'}
        out_ext = out_ext_map.get(actual_conversion_type)
        output_filename = f"converted_{timestamp}_{secure_filename(base_name)}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # 4. Perform Conversion (Inside a try block to catch conversion errors)
        # Note: error_key and conversion_success are initialized before this try
        try:
            # --- PDF to DOCX (pdf2docx library) ---
            if actual_conversion_type == 'pdf_to_docx':
                logger.info(f"Converting PDF -> DOCX (pdf2docx): {input_path_for_process} -> {output_path}")
                cv = None
                try:
                     cv = Converter(input_path_for_process)
                     cv.convert(output_path, start=0, end=None)
                     conversion_success = True
                except Exception as pdf2docx_err:
                     logger.error(f"pdf2docx conversion failed: {pdf2docx_err}", exc_info=True)
                     error_key = "err-conversion" # Keep generic or map specific pdf2docx errors
                finally:
                     if cv: cv.close()

            # --- DOCX/PPT to PDF (LibreOffice) ---
            elif actual_conversion_type in ['docx_to_pdf', 'ppt_to_pdf']:
                if not SOFFICE_PATH:
                    logger.error("LibreOffice path not configured.")
                    raise RuntimeError("err-libreoffice")

                logger.info(f"Converting {file_ext.upper()} -> PDF (LibreOffice): {input_path_for_process} -> {output_path}")
                output_dir = os.path.dirname(output_path)
                expected_lo_output_name = os.path.splitext(os.path.basename(input_path_for_process))[0] + '.pdf'
                temp_libreoffice_output = os.path.join(output_dir, expected_lo_output_name)
                safe_remove(temp_libreoffice_output)

                cmd = [SOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path_for_process]
                logger.debug(f"Executing LibreOffice command: {' '.join(cmd)}")

                # Inner try/except specific to LibreOffice call
                try:
                    result = subprocess.run(cmd, check=True, timeout=LIBREOFFICE_TIMEOUT, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                    logger.info(f"LibreOffice conversion stdout: {result.stdout[:200]}...")
                    if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0:
                        os.rename(temp_libreoffice_output, output_path)
                        conversion_success = True
                    else:
                         logger.error(f"LibreOffice command succeeded but output file '{temp_libreoffice_output}' invalid.")
                         error_key = "err-libreoffice"
                except subprocess.TimeoutExpired:
                    logger.error(f"LibreOffice process timed out.")
                    error_key = "err-conversion-timeout"
                    # No need to raise here, error_key is set
                except subprocess.CalledProcessError as lo_err:
                    logger.error(f"LibreOffice process error (Code {lo_err.returncode}): {lo_err.stderr}")
                    error_key = "err-libreoffice"
                except FileNotFoundError:
                    logger.error(f"LibreOffice command failed: '{SOFFICE_PATH}' not found.")
                    error_key = "err-libreoffice"
                # Catching generic Exception here for LO call is okay as it's specific
                except Exception as lo_ex:
                    logger.error(f"Unexpected error during LibreOffice conversion: {lo_ex}", exc_info=True)
                    error_key = "err-libreoffice"


            # --- PDF to PPTX (Python Image Method or LibreOffice Fallback) ---
            elif actual_conversion_type == 'pdf_to_ppt':
                python_method_success = False
                python_method_error_key = None
                try:
                    logger.info("Attempting PDF -> PPTX (Python Image Method)...")
                    if convert_pdf_to_pptx_python(input_path_for_process, output_path):
                        python_method_success = True
                        conversion_success = True
                        error_key = None # Clear default error if Python method works
                except ValueError as py_ppt_err:
                    python_method_error_key = str(py_ppt_err)
                    logger.warning(f"Python PDF->PPTX failed: {python_method_error_key}. Checking fallback.")
                except RuntimeError as py_rt_err:
                    python_method_error_key = str(py_rt_err)
                    logger.warning(f"Python PDF->PPTX failed: {python_method_error_key}. Checking fallback.")
                except Exception as py_gen_err:
                    python_method_error_key = "err-unknown"
                    logger.error(f"Unexpected Python PDF->PPTX error: {py_gen_err}", exc_info=True)
                    logger.warning("Checking fallback.")

                # Set error_key based on Python method outcome *before* fallback attempt
                if not python_method_success:
                     error_key = python_method_error_key or "err-conversion" # Use specific error or default

                # LibreOffice Fallback Logic
                can_fallback = not python_method_success and SOFFICE_PATH and python_method_error_key not in ["err-pdf-corrupt"]
                if can_fallback:
                    logger.info("Attempting PDF -> PPTX (LibreOffice fallback)...")
                    output_dir = os.path.dirname(output_path)
                    expected_lo_output_name = os.path.splitext(os.path.basename(input_path_for_process))[0] + '.pptx'
                    temp_libreoffice_output = os.path.join(output_dir, expected_lo_output_name)
                    safe_remove(temp_libreoffice_output)

                    cmd = [SOFFICE_PATH, '--headless', '--convert-to', 'pptx', '--outdir', output_dir, input_path_for_process]
                    logger.debug(f"Executing LO fallback: {' '.join(cmd)}")
                    try:
                        result = subprocess.run(cmd, check=True, timeout=LIBREOFFICE_TIMEOUT, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                        logger.info(f"LO fallback stdout: {result.stdout[:200]}...")
                        if os.path.exists(temp_libreoffice_output) and os.path.getsize(temp_libreoffice_output) > 0:
                            os.rename(temp_libreoffice_output, output_path)
                            conversion_success = True
                            error_key = None # IMPORTANT: Clear error if fallback succeeds
                            logger.info("LibreOffice fallback successful.")
                        else:
                             logger.error("LO fallback succeeded but output file invalid.")
                             # Keep the error from Python method if fallback also fails effectively
                             # error_key remains unchanged from python_method_error_key
                    except subprocess.TimeoutExpired:
                         logger.error(f"LO fallback timed out.")
                         error_key = "err-conversion-timeout" # Overwrite previous error with timeout
                    except subprocess.CalledProcessError as lo_err:
                         logger.error(f"LO fallback error (Code {lo_err.returncode}): {lo_err.stderr}")
                         # Keep python error if exists, otherwise set LO error
                         if not error_key or error_key == "err-conversion": error_key = "err-libreoffice"
                    except FileNotFoundError:
                         logger.error(f"LO fallback failed: '{SOFFICE_PATH}' not found.")
                         if not error_key or error_key == "err-conversion": error_key = "err-libreoffice"
                    except Exception as lo_ex:
                        logger.error(f"Unexpected LO fallback error: {lo_ex}", exc_info=True)
                        if not error_key or error_key == "err-conversion": error_key = "err-libreoffice"
                elif not python_method_success:
                     logger.warning(f"Skipping LO fallback. Python error: '{python_method_error_key}', LO available: {bool(SOFFICE_PATH)}")
                     # error_key already holds the python_method_error_key or default


        # <--- START: ADDED MISSING EXCEPT BLOCKS for the outer 'try' around conversion logic --->
        except subprocess.TimeoutExpired:
             error_key = "err-conversion-timeout"; logger.error(f"Conversion process timed out.")
        except subprocess.CalledProcessError as sub_err: # Should be caught by inner blocks, but as safeguard
             error_key = "err-libreoffice"; logger.error(f"Caught CalledProcessError: {sub_err}\nStderr: {sub_err.stderr}")
        except ValueError as val_err: # Catch specific errors from helpers if they bubble up
             error_key = str(val_err); logger.error(f"Caught Data/Value Error: {error_key}")
        except FileNotFoundError as fnf_err: # Should be caught by inner blocks, but as safeguard
             error_key = "err-libreoffice"; logger.error(f"Caught FileNotFoundError (LibreOffice?): {fnf_err}")
        except RuntimeError as rt_err: # Catch RuntimeErrors explicitly raised (like err-libreoffice)
             error_key = str(rt_err); logger.error(f"Caught RuntimeError: {error_key}")
        except Exception as conv_err: # Catch any other unexpected error during conversion itself
             error_key = "err-unknown"; logger.error(f"Caught unexpected conversion error: {conv_err}", exc_info=True)
        # <--- END: ADDED MISSING EXCEPT BLOCKS --->


        # 5. Handle Result (This section is now OUTSIDE the inner try/except for conversion logic)
        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            # ... (Phần gửi file và cleanup thành công như cũ) ...
            logger.info(f"Conversion successful for '{filename}'. Sending file '{output_filename}'. Time: {time.time() - start_time:.2f}s")
            mimetype_map = {
                'pdf': 'application/pdf',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            }
            mimetype = mimetype_map.get(out_ext, 'application/octet-stream')
            try:
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                @response.call_on_close
                def cleanup():
                    logger.debug(f"Cleaning up after successful /convert request: {input_path_for_process}, {output_path}")
                    safe_remove(input_path_for_process)
                    safe_remove(output_path)
                return response
            except Exception as send_err:
                 logger.error(f"Error sending file {output_path}: {send_err}", exc_info=True)
                 safe_remove(input_path_for_process) # Try cleanup even if sending failed
                 safe_remove(output_path)
                 return make_error_response("err-unknown", 500)
        else:
            # Conversion failed or produced invalid output
            if not error_key: error_key = "err-conversion" # Ensure an error key exists
            logger.error(f"Conversion failed for '{filename}'. Final Error key: {error_key}. Time: {time.time() - start_time:.2f}s")
            # Raise error to trigger the main cleanup block below
            raise RuntimeError(error_key)

    # Main exception handler for the entire request processing
    except Exception as e:
         # Catch errors raised explicitly (like RuntimeError(error_key)) or unexpected errors
         final_error_key = str(e) if str(e).startswith("err-") else "err-unknown"
         status_code = 500 if final_error_key == "err-unknown" else 400 # Use 500 for unexpected

         if final_error_key == "err-unknown":
             logger.error(f"Unexpected error in /convert handler: {e}", exc_info=True)
         # else: Specific errors were already logged when they occurred or were raised

         # --- Cleanup Phase on Any Error ---
         logger.debug(f"Cleaning up after failed /convert request (Error: {final_error_key}).")
         for p in saved_input_paths: safe_remove(p)
         safe_remove(output_path) # Attempt removal even if None or not existing
         if temp_libreoffice_output and os.path.exists(temp_libreoffice_output):
             safe_remove(temp_libreoffice_output)

         return make_error_response(final_error_key, status_code)


# === PDF / Image Conversion Route ===
# ... (Hàm convert_image_route không thay đổi so với phiên bản trước) ...
@app.route('/convert_image', methods=['POST'])
def convert_image_route():
    """Handles PDF -> Image (ZIP) and Image(s) -> PDF conversions."""
    output_path = None
    input_path_for_pdf_input = None # Only used if input is PDF
    saved_input_paths = [] # Store path of saved PDF input
    actual_conversion_type = None # 'pdf_to_image' or 'image_to_pdf'
    output_filename = None
    start_time = time.time()
    error_key = "err-conversion" # Default error
    conversion_success = False   # Initialize flag

    try:
        # 1. Validation
        uploaded_files = request.files.getlist('image_file')
        if not uploaded_files or not all(f and f.filename for f in uploaded_files):
            logger.warning("Request /convert_image received no valid files.")
            return make_error_response("err-select-file", 400)

        logger.info(f"Request /convert_image: Received {len(uploaded_files)} file(s).")

        # --- Size Check ---
        total_size = 0
        filenames_for_log = []
        for f in uploaded_files:
            try:
                f.stream.seek(0, os.SEEK_END)
                total_size += f.stream.tell()
                f.stream.seek(0)
                filenames_for_log.append(secure_filename(f.filename))
            except Exception as size_err:
                logger.error(f"Size check failed for {f.filename}: {size_err}", exc_info=True)
                return make_error_response("err-unknown", 500)
        logger.info(f"Files: {', '.join(filenames_for_log)}. Total size: {total_size} bytes.")
        if total_size > 100 * 1024 * 1024: # 100 MiB limit
            logger.warning(f"Total file size exceeds limit ({total_size} bytes).")
            return make_error_response("err-file-too-large", 413)

        # --- Type and Count Validation ---
        first_file = uploaded_files[0]
        first_filename = secure_filename(first_file.filename)
        first_ext = first_filename.rsplit('.', 1)[-1].lower() if '.' in first_filename else ''

        validation_error_key = None
        out_ext = None

        if first_ext == 'pdf':
            if len(uploaded_files) > 1:
                validation_error_key = "err-image-single-pdf"
            elif not _allowed_file(first_filename, ALLOWED_IMAGE_EXTENSIONS):
                validation_error_key = "err-image-format"
            else:
                actual_conversion_type = 'pdf_to_image'
                out_ext = 'zip'
        elif first_ext in ['jpg', 'jpeg']:
            actual_conversion_type = 'image_to_pdf'
            out_ext = 'pdf'
            for f in uploaded_files:
                 fname_sec = secure_filename(f.filename)
                 f_ext = fname_sec.rsplit('.', 1)[-1].lower() if '.' in fname_sec else ''
                 if f_ext not in ['jpg', 'jpeg']:
                     validation_error_key = "err-image-all-images"; break
        else:
            validation_error_key = "err-image-format"

        if validation_error_key:
            logger.warning(f"Image conversion validation failed: {validation_error_key}")
            return make_error_response(validation_error_key, 400)

        logger.info(f"Determined conversion type: {actual_conversion_type}")

        # 2. Save Input PDF (if needed) & Define Output
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")

        if actual_conversion_type == 'pdf_to_image':
            input_filename_ts = f"input_{timestamp}_{first_filename}"
            input_path_for_pdf_input = os.path.join(UPLOAD_FOLDER, input_filename_ts)
            try:
                first_file.save(input_path_for_pdf_input)
                saved_input_paths.append(input_path_for_pdf_input)
                logger.info(f"Input PDF saved: {input_path_for_pdf_input}")
            except Exception as save_err:
                logger.error(f"Failed to save input PDF {first_filename}: {save_err}", exc_info=True)
                return make_error_response("err-unknown", 500)

        base_name = first_filename.rsplit('.', 1)[0]
        output_filename = f"converted_{timestamp}_{secure_filename(base_name)}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # 3. Perform Conversion
        try:
            if actual_conversion_type == 'pdf_to_image':
                logger.info(f"Converting PDF -> Images (ZIP): {input_path_for_pdf_input} -> {output_path}")
                img_fmt = 'jpeg'
                if convert_pdf_to_image_zip(input_path_for_pdf_input, output_path, img_format=img_fmt):
                    conversion_success = True
            elif actual_conversion_type == 'image_to_pdf':
                logger.info(f"Converting {len(uploaded_files)} Image(s) -> PDF: -> {output_path}")
                if convert_images_to_pdf(uploaded_files, output_path):
                    conversion_success = True

        except ValueError as val_err:
            error_key = str(val_err)
            logger.error(f"Image conversion Value Error: {error_key}", exc_info=False)
        except RuntimeError as rt_err:
            error_key = str(rt_err)
            logger.error(f"Image conversion Runtime Error: {error_key}", exc_info=True)
        except Exception as conv_err:
            error_key = "err-unknown"
            logger.error(f"Unknown image conversion error: {conv_err}", exc_info=True)

        # 4. Handle Result
        if conversion_success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            logger.info(f"Image conversion successful for '{first_filename}'. Sending '{output_filename}'. Time: {time.time() - start_time:.2f}s")
            mimetype = 'application/zip' if out_ext == 'zip' else 'application/pdf'

            try:
                response = send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
                @response.call_on_close
                def cleanup():
                    logger.debug(f"Cleaning up after successful /convert_image: {input_path_for_pdf_input}, {output_path}")
                    if input_path_for_pdf_input: safe_remove(input_path_for_pdf_input)
                    safe_remove(output_path)
                return response
            except Exception as send_err:
                 logger.error(f"Error sending file {output_path}: {send_err}", exc_info=True)
                 if input_path_for_pdf_input: safe_remove(input_path_for_pdf_input)
                 safe_remove(output_path)
                 return make_error_response("err-unknown", 500)
        else:
            if not error_key: error_key = "err-conversion"
            logger.error(f"Image conversion failed. Final Error key: {error_key}. Time: {time.time() - start_time:.2f}s")
            raise RuntimeError(error_key)

    # Main exception handler for the route
    except Exception as e:
         final_error_key = str(e) if str(e).startswith("err-") else "err-unknown"
         status_code = 500 if final_error_key == "err-unknown" else 400

         if final_error_key == "err-unknown":
             logger.error(f"Unexpected error in /convert_image handler: {e}", exc_info=True)

         logger.debug(f"Cleaning up after failed /convert_image (Error: {final_error_key}).")
         for p in saved_input_paths: safe_remove(p)
         safe_remove(output_path)

         return make_error_response(final_error_key, status_code)


# --- Teardown (Cleanup old files) ---
@app.teardown_appcontext
def cleanup_old_files(exception=None):
    """Cleans up files older than max_age in the UPLOAD_FOLDER."""
    if not os.path.exists(UPLOAD_FOLDER): return

    logger.debug("Running teardown cleanup...")
    try:
        now = time.time()
        max_age = 3600 # 1 hour
        deleted_count = 0
        checked_count = 0
        try: # Add try/except around listdir itself
            items = os.listdir(UPLOAD_FOLDER)
        except OSError as list_err:
             logger.error(f"Error listing upload directory {UPLOAD_FOLDER}: {list_err}")
             return # Cannot proceed if directory cannot be listed

        for filename in items:
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                stat_result = os.stat(path)
                # Currently only cleaning files, not directories
                if not os.path.isdir(path):
                     file_age = now - stat_result.st_mtime
                     checked_count += 1
                     if file_age > max_age:
                         logger.debug(f"Removing old file: {path} (Age: {file_age:.0f}s)")
                         if safe_remove(path):
                             deleted_count += 1
            except FileNotFoundError: continue
            except Exception as e: logger.error(f"Error during teardown check for {path}: {e}")

        if checked_count > 0 or deleted_count > 0: # Log even if only checked
            logger.info(f"Teardown cleanup finished. Checked {checked_count}, removed {deleted_count} files older than {max_age}s.")
        else:
            logger.debug("Teardown cleanup: No items checked or removed.")
    except Exception as e:
        logger.error(f"Critical error during teardown cleanup process: {e}", exc_info=True)

# === Main Execution ===
if __name__ == '__main__':
    # Ensure upload folder exists at startup
    try:
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        logger.info(f"Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    except OSError as mkdir_err:
         logger.critical(f"FATAL: Could not create upload folder {UPLOAD_FOLDER}: {mkdir_err}. Exiting.")
         sys.exit(1) # Exit if upload folder cannot be created

    logger.info(f"LibreOffice path: {SOFFICE_PATH or 'Not Found'}")

    # Check for translations file
    if not os.path.exists(TRANSLATIONS_FILE):
        logger.warning(f"Translation file '{TRANSLATIONS_FILE}' not found. Language features might be limited.")

    port = int(os.environ.get('PORT', 5003))
    host = os.environ.get('HOST', '0.0.0.0')
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() in ['true', '1', 't']

    logger.info(f"Attempting to start server on {host}:{port} - Debug mode: {debug_mode}")

    # Server execution logic
    if debug_mode:
        logger.warning("Running in DEBUG mode with Flask development server.")
        # Consider use_reloader=False if having issues
        app.run(host=host, port=port, debug=True, threaded=True, use_reloader=True)
    else:
        try:
            from waitress import serve
            logger.info("Running with Waitress production server.")
            serve(app, host=host, port=port, threads=8) # Adjust threads as needed
        except ImportError:
            logger.error("Waitress not found! Install with 'pip install waitress'.")
            logger.warning("FALLING BACK TO FLASK DEVELOPMENT SERVER (NOT RECOMMENDED FOR PRODUCTION).")
            app.run(host=host, port=port, debug=False, threaded=True)

# --- END OF FILE app.py ---