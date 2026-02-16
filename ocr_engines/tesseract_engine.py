import pytesseract
from PIL import Image
import os
from preprocessing.image_preprocessing import preprocess_image

# POINT THIS TO YOUR TESSERACT EXECUTABLE IF NOT IN PATH
# POINT THIS TO YOUR TESSERACT EXECUTABLE IF NOT IN PATH
import shutil
if not shutil.which('tesseract'):
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\Users\anshu\AppData\Local\Tesseract-OCR\tesseract.exe'
    ]
    for p in possible_paths:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            break

def extract_text(image_path, preprocess=True):
    """
    Extracts text from an image using Tesseract OCR.
    Args:
        image_path (str): Path to the image file.
        preprocess (bool): Whether to apply preprocessing before OCR.
    Returns:
        str: Extracted text.
    """
    try:
        if preprocess:
            # Get preprocessed PIL image
            image = preprocess_image(image_path)
        else:
            image = Image.open(image_path)

        # Perform OCR
        text = pytesseract.image_to_string(image)
        return text.strip()
    except Exception as e:
        return f"Error in Tesseract OCR: {str(e)}"
