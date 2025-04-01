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
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from PIL import Image

app = Flask(__name__, template_folder='templates')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Configure upload folder
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'jpg', 'jpeg'}


# Health check endpoint
@app.route('/health')
def health_check():
    return 'OK', 200


def find_libreoffice():
    possible_paths = [
        'soffice',
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/usr/lib/libreoffice/program/soffice',
    ]

    for path in possible_paths:
        try:
            if os.path.isfile(path) or shutil.which(path):
                result = subprocess.run([path, '--version'],
                                        capture_output=True, text=True, check=False)
                if result.returncode == 0:
                    logger.info(f"Found LibreOffice at: {path}")
                    return path
        except Exception as e:
            logger.warning(f"Error checking LibreOffice at {path}: {e}")

    logger.warning("Using default 'soffice' path")
    return 'soffice'


SOFFICE_PATH = find_libreoffice()
logger.info(f"Using LibreOffice path: {SOFFICE_PATH}")


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_remove(file_path, retries=5, delay=1):
    for i in range(retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
        except Exception as e:
            logger.warning(f"Failed to remove {file_path} (attempt {i + 1}): {e}")
            time.sleep(delay)
    return False


def convert_pdf_to_pptx_python(input_path, output_path):
    try:
        # Convert PDF to images with high quality
        images = convert_from_path(input_path, dpi=300, fmt='jpeg')

        if not images:
            raise ValueError("No pages found in PDF")

        # Create presentation with proper aspect ratio
        prs = Presentation()

        # Detect page ratio from first page
        first_page = images[0]
        page_ratio = first_page.width / first_page.height

        # Set slide size based on page ratio (16:9 or 4:3)
        if abs(page_ratio - 16 / 9) < abs(page_ratio - 4 / 3):
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(5.625)  # 16:9 ratio
        else:
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)  # 4:3 ratio

        blank_layout = prs.slide_layouts[6]

        for image in images:
            # Use in-memory buffer instead of temp files
            img_bytes = BytesIO()
            image.save(img_bytes, format='JPEG', quality=95)
            img_bytes.seek(0)

            slide = prs.slides.add_slide(blank_layout)

            # Calculate image dimensions to maintain aspect ratio
            img_ratio = image.width / image.height
            slide_ratio = prs.slide_width / prs.slide_height

            if img_ratio > slide_ratio:
                # Image is wider than slide - fit to width
                width = prs.slide_width
                height = width / img_ratio
            else:
                # Image is taller than slide - fit to height
                height = prs.slide_height
                width = height * img_ratio

            # Center the image
            left = (prs.slide_width - width) / 2
            top = (prs.slide_height - height) / 2

            slide.shapes.add_picture(img_bytes, left, top, width, height)

        prs.save(output_path)
        return True

    except Exception as e:
        logger.error(f"PDF to PPTX conversion error: {e}")
        return False


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    input_path = output_path = None
    try:
        # Validate input
        if 'file' not in request.files:
            return "No file uploaded", 400

        file = request.files['file']
        if not file or file.filename == '':
            return "No file selected", 400

        if not allowed_file(file.filename):
            return "Invalid file type", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            return "No conversion type selected", 400

        # Prepare upload directory
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved: {input_path}")

        # Determine conversion type
        ext = filename.rsplit('.', 1)[1].lower()
        conversions = {
            'pdf_to_docx': ('pdf', 'docx'),
            'docx_to_pdf': ('docx', 'pdf'),
            'pdf_to_ppt': ('pdf', 'pptx'),
            'ppt_to_pdf': (['ppt', 'pptx'], 'pdf'),
            'pdf_docx': ('pdf', 'docx') if ext == 'pdf' else ('docx', 'pdf'),
            'pdf_ppt': ('pdf', 'pptx') if ext == 'pdf' else (['ppt', 'pptx'], 'pdf')
        }

        if conversion_type not in conversions:
            return "Invalid conversion type", 400

        valid_exts, out_ext = conversions[conversion_type]
        if isinstance(valid_exts, list):
            if ext not in valid_exts:
                return "File type mismatch", 400
        elif ext != valid_exts:
            return "File type mismatch", 400

        # Generate output filename
        base_name = filename.rsplit('.', 1)[0]
        output_filename = f"converted_{base_name}.{out_ext}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # Perform conversion
        if conversion_type in ['pdf_to_docx', 'pdf_docx'] and ext == 'pdf':
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        elif conversion_type in ['docx_to_pdf', 'pdf_docx'] and ext == 'docx':
            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], check=True, timeout=60)

            temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
            if os.path.exists(temp_output):
                os.rename(temp_output, output_path)
            else:
                raise Exception("Conversion failed - no output file")

        elif conversion_type in ['pdf_to_ppt', 'pdf_ppt'] and ext == 'pdf':
            # Try LibreOffice first
            try:
                result = subprocess.run([
                    SOFFICE_PATH,
                    '--headless',
                    '--infilter="draw_pdf_import"',
                    '--convert-to', 'pptx',
                    '--outdir', UPLOAD_FOLDER,
                    input_path
                ], check=True, timeout=120)

                temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pptx")
                if os.path.exists(temp_output):
                    os.rename(temp_output, output_path)
                else:
                    raise Exception("LibreOffice conversion failed")

            except Exception as e:
                logger.warning(f"LibreOffice failed, trying python-pptx: {e}")
                if not convert_pdf_to_pptx_python(input_path, output_path):
                    raise Exception("All conversion methods failed")

        elif conversion_type in ['ppt_to_pdf', 'pdf_ppt'] and ext in ['ppt', 'pptx']:
            result = subprocess.run([
                SOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], check=True, timeout=60)

            temp_output = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
            if os.path.exists(temp_output):
                os.rename(temp_output, output_path)
            else:
                raise Exception("Conversion failed - no output file")

        else:
            return "Unsupported conversion", 400

        # Return converted file
        with open(output_path, 'rb') as f:
            file_data = f.read()

        mimetypes = {
            'pdf': 'application/pdf',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        }

        return Response(
            file_data,
            mimetype=mimetypes.get(out_ext, 'application/octet-stream'),
            headers={'Content-Disposition': f'attachment; filename={output_filename}'}
        )

    except Exception as e:
        logger.error(f"Conversion error: {e}")
        return f"Conversion failed: {str(e)}", 500

    finally:
        for path in [input_path, output_path]:
            if path and os.path.exists(path):
                safe_remove(path)


@app.teardown_appcontext
def cleanup(exception=None):
    if not os.path.exists(UPLOAD_FOLDER):
        return

    try:
        now = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            try:
                if os.path.isfile(path) and os.path.getmtime(path) < now - 3600:
                    safe_remove(path)
            except Exception as e:
                logger.error(f"Cleanup error for {path}: {e}")
    except Exception as e:
        logger.error(f"Cleanup error: {e}")


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5003))
    logger.info(f"Starting server on port {port}")
    app.run(host='0.0.0.0', port=port)