from flask import Flask, request, send_file, render_template, Response
import os
import time
import subprocess
import logging
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import img2pdf  # Để chuyển JPG sang PDF
import fitz  # PyMuPDF để chuyển PDF sang JPG

app = Flask(__name__, template_folder='templates')

# Cấu hình logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Cấu hình thư mục tạm
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'ppt', 'pptx', 'xls', 'xlsx', 'jpg', 'jpeg'}

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
            logger.error("Invalid file type, only PDF, DOCX, PPT, Excel, JPG supported")
            return "Only PDF, DOCX, PPT, Excel, JPG files are supported", 400

        conversion_type = request.form.get('conversion_type')
        if not conversion_type:
            logger.error("No conversion type selected")
            return "Please select a conversion type", 400

        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)
        logger.info(f"File saved at: {input_path}")

        # Xử lý các loại chuyển đổi
        if conversion_type == 'pdf_to_docx' and filename.endswith('.pdf'):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.docx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            logger.info(f"Starting PDF to DOCX conversion: {input_path} -> {output_path}")
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
            logger.info("PDF to DOCX conversion completed")

        elif conversion_type == 'docx_to_pdf' and filename.endswith('.docx'):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
            logger.info(f"Starting DOCX to PDF conversion: {input_path} -> {output_path}")
            try:
                soffice_check = subprocess.run(['soffice', '--version'], capture_output=True, text=True, check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            result = subprocess.run([
                'soffice',
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

        # Chuyển đổi PDF sang PPT (dùng LibreOffice)
        elif conversion_type == 'pdf_to_ppt' and filename.endswith('.pdf'):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.ppt"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.ppt')
            logger.info(f"Starting PDF to PPT conversion: {input_path} -> {output_path}")
            try:
                soffice_check = subprocess.run(['soffice', '--version'], capture_output=True, text=True, check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            result = subprocess.run([
                'soffice',
                '--headless',
                '--convert-to', 'ppt',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], capture_output=True, text=True, check=True, timeout=60)
            logger.info(f"soffice stdout: {result.stdout}")
            if result.stderr:
                logger.warning(f"soffice stderr: {result.stderr}")
            if not os.path.exists(actual_output_path):
                logger.error(f"Output file not created: {actual_output_path}")
                return "Error converting PDF to PPT", 500
            if actual_output_path != output_path:
                os.rename(actual_output_path, output_path)
                logger.info(f"Renamed file from {actual_output_path} to {output_path}")
            logger.info("PDF to PPT conversion completed")

        # Chuyển đổi PPT sang PDF
        elif conversion_type == 'ppt_to_pdf' and (filename.endswith('.ppt') or filename.endswith('.pptx')):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
            logger.info(f"Starting PPT to PDF conversion: {input_path} -> {output_path}")
            try:
                soffice_check = subprocess.run(['soffice', '--version'], capture_output=True, text=True, check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            result = subprocess.run([
                'soffice',
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

        # Chuyển đổi PDF sang Excel (dùng LibreOffice)
        elif conversion_type == 'pdf_to_excel' and filename.endswith('.pdf'):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.xlsx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.xlsx')
            logger.info(f"Starting PDF to Excel conversion: {input_path} -> {output_path}")
            try:
                soffice_check = subprocess.run(['soffice', '--version'], capture_output=True, text=True, check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            result = subprocess.run([
                'soffice',
                '--headless',
                '--convert-to', 'xlsx',
                '--outdir', UPLOAD_FOLDER,
                input_path
            ], capture_output=True, text=True, check=True, timeout=60)
            logger.info(f"soffice stdout: {result.stdout}")
            if result.stderr:
                logger.warning(f"soffice stderr: {result.stderr}")
            if not os.path.exists(actual_output_path):
                logger.error(f"Output file not created: {actual_output_path}")
                return "Error converting PDF to Excel", 500
            if actual_output_path != output_path:
                os.rename(actual_output_path, output_path)
                logger.info(f"Renamed file from {actual_output_path} to {output_path}")
            logger.info("PDF to Excel conversion completed")

        # Chuyển đổi Excel sang PDF
        elif conversion_type == 'excel_to_pdf' and (filename.endswith('.xls') or filename.endswith('.xlsx')):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            actual_output_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
            logger.info(f"Starting Excel to PDF conversion: {input_path} -> {output_path}")
            try:
                soffice_check = subprocess.run(['soffice', '--version'], capture_output=True, text=True, check=True)
                logger.info(f"LibreOffice version: {soffice_check.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice not working: {e}")
                return "Error: LibreOffice is not installed or not working", 500
            result = subprocess.run([
                'soffice',
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
                return "Error converting Excel to PDF", 500
            if actual_output_path != output_path:
                os.rename(actual_output_path, output_path)
                logger.info(f"Renamed file from {actual_output_path} to {output_path}")
            logger.info("Excel to PDF conversion completed")

        # Chuyển đổi PDF sang JPG (dùng PyMuPDF - fitz)
        elif conversion_type == 'pdf_to_jpg' and filename.endswith('.pdf'):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.jpg"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            logger.info(f"Starting PDF to JPG conversion: {input_path} -> {output_path}")
            pdf_document = fitz.open(input_path)
            if pdf_document.page_count == 0:
                pdf_document.close()
                logger.error("PDF file is empty")
                return "Error: PDF file is empty", 400
            page = pdf_document.load_page(0)  # Chỉ lấy trang đầu tiên
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))  # Tăng độ phân giải (300 DPI)
            pix.save(output_path)
            pdf_document.close()
            logger.info("PDF to JPG conversion completed")

        # Chuyển đổi JPG sang PDF (dùng img2pdf)
        elif conversion_type == 'jpg_to_pdf' and (filename.endswith('.jpg') or filename.endswith('.jpeg')):
            output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            logger.info(f"Starting JPG to PDF conversion: {input_path} -> {output_path}")
            with open(output_path, "wb") as f:
                f.write(img2pdf.convert(input_path))
            logger.info("JPG to PDF conversion completed")

        else:
            logger.error("File type does not match conversion type or conversion not supported")
            return "File type does not match conversion type or conversion not supported", 400

        with open(output_path, 'rb') as f:
            file_data = f.read()
        logger.info(f"Output file read: {output_path}")

        safe_remove(input_path)
        safe_remove(output_path)

        return Response(
            file_data,
            mimetype='application/octet-stream',
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5003)))