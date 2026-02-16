from dotenv import load_dotenv
import os
load_dotenv() 



from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import sys

# Add project root to system path to ensure modules are found
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ocr_engines import tesseract_engine, easyocr_engine, paddle_engine

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_demo_purposes'

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_ocr_engine(engine_name):
    if engine_name == 'tesseract':
        return tesseract_engine
    elif engine_name == 'easyocr':
        return easyocr_engine
    elif engine_name == 'paddle':
        return paddle_engine
    else:
        return None

def process_file(file_path, engine):
    """
    Processes a file (Image or PDF) using the selected OCR engine.
    """
    filename = os.path.basename(file_path)
    ext = filename.rsplit('.', 1)[1].lower()
    
    extracted_text = ""

    if ext == 'pdf':
        try:
            # Convert PDF to images
            # poppler_path might need to be configured on Windows if not in PATH
            # For this environment, we assume it's set up or we handle the error.
            images = convert_from_path(file_path)
            
            for i, image in enumerate(images):
                # Save temp image for processing
                temp_img_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_page_{i}_{filename}.jpg")
                image.save(temp_img_path, 'JPEG')
                
                # Extract text
                page_text = engine.extract_text(temp_img_path)
                extracted_text += f"\n--- Page {i+1} ---\n{page_text}\n"
                
                # Cleanup temp image
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)
                    
        except Exception as e:
            return f"Error processing PDF: {str(e)}. Ensure Poppler is installed and in PATH."
            
    else:
        # It's an image
        extracted_text = engine.extract_text(file_path)
        
    return extracted_text

@app.route('/', methods=['GET', 'POST'])
def index():
    extracted_text = None
    
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        engine_name = request.form.get('engine')
        
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
            
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            engine = get_ocr_engine(engine_name)
            if not engine:
                flash('Invalid OCR engine selected')
                return redirect(request.url)
                
            # Process the file
            extracted_text = process_file(file_path, engine)
            
            # Cleanup uploaded file (optional, keeping it simple for now)
            # if os.path.exists(file_path):
            #     os.remove(file_path)
                
    return render_template('index.html', extracted_text=extracted_text)

if __name__ == '__main__':
    app.run(debug=True)
