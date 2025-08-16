"""
PDF to Word OCR Converter - Web Application
Upload PDF files and download converted Word documents
"""
import os
import sys
from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for, Response
from werkzeug.utils import secure_filename
import uuid
from datetime import datetime
import threading
import time

# Add the parent directory to Python path to import our converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import fitz  # PyMuPDF
    from docx import Document
    from spellchecker import SpellChecker
    import pytesseract
    from PIL import Image
    import io
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install: pip install flask PyMuPDF python-docx pyspellchecker pytesseract Pillow")
    sys.exit(1)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'pdf-converter-secret-key-2025'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# Store conversion status
conversion_status = {}

def setup_tesseract():
    """Setup Tesseract path for Windows"""
    possible_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(os.getenv('USERNAME'))
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    
    try:
        pytesseract.image_to_string(Image.new('RGB', (100, 100), color='white'))
        return True
    except:
        return False

def extract_text_from_page(page):
    """Extract text from both regular text and images in a PDF page"""
    regular_text = page.get_text().strip()
    image_text = ""
    
    try:
        image_list = page.get_images()
        
        for img_index, img in enumerate(image_list):
            xref = img[0]
            pix = fitz.Pixmap(page.parent, xref)
            
            if pix.n - pix.alpha < 4:
                img_data = pix.tobytes("png")
                pil_image = Image.open(io.BytesIO(img_data))
                
                ocr_text = pytesseract.image_to_string(pil_image, lang="eng", config='--psm 6')
                if ocr_text.strip():
                    image_text += f"\n[Text from image {img_index + 1}:]\n{ocr_text.strip()}\n"
            
            pix = None
            
    except Exception as e:
        try:
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            pil_image = Image.open(io.BytesIO(img_data))
            
            page_ocr_text = pytesseract.image_to_string(pil_image, lang="eng", config='--psm 6')
            
            if len(page_ocr_text.strip()) > len(regular_text):
                return page_ocr_text.strip(), "Full page OCR"
                
        except Exception as ocr_error:
            pass
    
    combined_text = regular_text
    if image_text:
        combined_text = regular_text + "\n" + image_text
    
    extraction_method = "Text"
    if image_text and regular_text:
        extraction_method = "Text + Image OCR"
    elif image_text and not regular_text:
        extraction_method = "Image OCR only"
    elif not regular_text and not image_text:
        extraction_method = "No text found"
    
    return combined_text.strip(), extraction_method

def convert_pdf_to_word_web(pdf_path, output_path, job_id):
    """Convert PDF to Word with progress tracking for web interface"""
    try:
        conversion_status[job_id] = {
            'status': 'starting',
            'progress': 0,
            'message': 'Initializing conversion...',
            'current_page': 0,
            'total_pages': 0
        }
        
        # Setup Tesseract
        setup_tesseract()
        
        # Open PDF
        pdf_document = fitz.open(pdf_path)
        total_pages = len(pdf_document)
        
        conversion_status[job_id].update({
            'total_pages': total_pages,
            'message': f'Processing {total_pages} pages...'
        })
        
        # Create Word document
        doc = Document()
        pdf_name = os.path.basename(pdf_path)
        doc.add_heading(f'{pdf_name}', 0)
        
        total_words = 0
        pages_with_text = 0
        
        for page_num in range(total_pages):
            # Update progress
            progress = int((page_num / total_pages) * 100)
            conversion_status[job_id].update({
                'status': 'processing',
                'progress': progress,
                'current_page': page_num + 1,
                'message': f'Processing page {page_num + 1} of {total_pages}...'
            })
            
            # Extract text
            page = pdf_document[page_num]
            text, extraction_method = extract_text_from_page(page)
            
            if text:
                pages_with_text += 1
                word_count = len(text.split())
                total_words += word_count
                
                # Add to document
                doc.add_heading(f'Page {page_num + 1}', level=1)
                doc.add_paragraph(text)
                doc.add_page_break()
        
        # Add summary
        doc.add_heading('Document Information', level=1)
        doc.add_paragraph(f"Source File: {pdf_name}")
        doc.add_paragraph(f"Total Pages: {total_pages}")
        doc.add_paragraph(f"Pages with Text: {pages_with_text}")
        doc.add_paragraph(f"Total Words: {total_words}")
        doc.add_paragraph(f"Converted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Save document
        doc.save(output_path)
        pdf_document.close()
        
        # Update final status
        conversion_status[job_id].update({
            'status': 'completed',
            'progress': 100,
            'message': 'Conversion completed successfully!',
            'output_file': os.path.basename(output_path),
            'stats': {
                'total_pages': total_pages,
                'pages_with_text': pages_with_text,
                'total_words': total_words
            }
        })
        
        return True
        
    except Exception as e:
        conversion_status[job_id].update({
            'status': 'error',
            'progress': 0,
            'message': f'Error: {str(e)}'
        })
        return False

@app.route('/')
def index():
    """Main page with upload form"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and start conversion"""
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'No file selected'}), 400
    
    file = request.files['pdf_file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and file.filename.lower().endswith('.pdf'):
        # Generate unique job ID
        job_id = str(uuid.uuid4())
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_{filename}")
        file.save(pdf_path)
        
        # Generate output filename
        output_filename = f"{os.path.splitext(filename)[0]}_converted.docx"
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], f"{job_id}_{output_filename}")
        
        # Start conversion in background thread
        thread = threading.Thread(
            target=convert_pdf_to_word_web,
            args=(pdf_path, output_path, job_id)
        )
        thread.start()
        
        return jsonify({
            'job_id': job_id,
            'message': 'Upload successful! Conversion started.',
            'filename': filename
        })
    
    return jsonify({'error': 'Please select a valid PDF file'}), 400

@app.route('/progress/<job_id>')
def get_progress(job_id):
    """Get conversion progress"""
    if job_id in conversion_status:
        return jsonify(conversion_status[job_id])
    return jsonify({'error': 'Job not found'}), 404

@app.route('/download/<job_id>')
def download_file(job_id):
    """Download converted Word document"""
    print(f"Download request for job_id: {job_id}")
    
    if job_id in conversion_status:
        status = conversion_status[job_id]
        print(f"Job status: {status}")
        
        if status['status'] == 'completed':
            output_file = status['output_file']
            # The output_file already has the job_id prefix, so use it directly
            file_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_file)
            
            print(f"Looking for file: {file_path}")
            print(f"File exists: {os.path.exists(file_path)}")
            
            if os.path.exists(file_path):
                # Create a clean filename for download (remove job_id prefix)
                clean_filename = output_file.replace(f"{job_id}_", "")
                print(f"Sending file as: {clean_filename}")
                
                # Simple file response
                def generate():
                    with open(file_path, 'rb') as f:
                        while True:
                            data = f.read(4096)
                            if not data:
                                break
                            yield data
                
                return Response(
                    generate(),
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    headers={
                        'Content-Disposition': f'attachment; filename="{clean_filename}"',
                        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    }
                )
            else:
                print("File not found on disk")
                return jsonify({'error': f'File not found: {output_file}'}), 404
        else:
            print(f"Job not completed. Status: {status['status']}")
            return jsonify({'error': f'Conversion not completed. Status: {status["status"]}'}), 400
    else:
        print("Job ID not found in conversion_status")
        return jsonify({'error': 'Job not found'}), 404

@app.route('/test')
def test_page():
    """Test page for debugging downloads"""
    return render_template('test.html')

@app.route('/test-download')
def test_download():
    """Test download functionality with existing file"""
    download_folder = app.config['DOWNLOAD_FOLDER']
    files = os.listdir(download_folder) if os.path.exists(download_folder) else []
    
    if files:
        # Use the first available file for testing
        test_file = files[0]
        file_path = os.path.join(download_folder, test_file)
        
        # Extract original filename from the prefixed filename
        original_name = test_file.split('_', 1)[1] if '_' in test_file else test_file
        
        try:
            return send_file(
                os.path.abspath(file_path),
                as_attachment=True,
                attachment_filename=original_name
            )
        except Exception as e:
            return jsonify({'error': f'Download failed: {str(e)}'}), 500
    else:
        return jsonify({'error': 'No files available for download'}), 404

@app.route('/debug/<job_id>')
def debug_job(job_id):
    """Debug endpoint to check job status and files"""
    info = {
        'job_id': job_id,
        'job_exists': job_id in conversion_status,
        'download_folder': app.config['DOWNLOAD_FOLDER'],
        'download_folder_exists': os.path.exists(app.config['DOWNLOAD_FOLDER']),
        'files_in_download_folder': []
    }
    
    if os.path.exists(app.config['DOWNLOAD_FOLDER']):
        info['files_in_download_folder'] = os.listdir(app.config['DOWNLOAD_FOLDER'])
    
    if job_id in conversion_status:
        info['job_status'] = conversion_status[job_id]
    
    return jsonify(info)

@app.route('/cleanup')
def cleanup_files():
    """Clean up old files (optional endpoint)"""
    try:
        # Remove files older than 1 hour
        current_time = time.time()
        
        for folder in [app.config['UPLOAD_FOLDER'], app.config['DOWNLOAD_FOLDER']]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getctime(file_path)
                    if file_age > 3600:  # 1 hour
                        os.remove(file_path)
        
        return jsonify({'message': 'Cleanup completed'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    # Create directories if they don't exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
    
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    
    print("🌐 PDF to Word OCR Converter Web App")
    print("=" * 50)
    print("Starting web server...")
    if debug:
        print("Open your browser and go to: http://localhost:5000")
    else:
        print("Production mode - serving on all interfaces")
    print("=" * 50)
    
    app.run(debug=debug, host='0.0.0.0', port=port)
