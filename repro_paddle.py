import os
os.environ['KMP_DUPLICATE_LIB_OK'] = 'True'

try:
    import torch
    print("Torch imported successfully.")
except ImportError:
    print("Torch not found (will be imported by paddleocr/modelscope later).")
except Exception as e:
    print(f"Error importing torch: {e}")

try:
    import paddle
    # ✅ CRITICAL FIX: disable oneDNN
    paddle.set_flags({'FLAGS_use_onednn': False})
    print("Paddle imported and flags set.")
except Exception as e:
    print(f"Error importing paddle: {e}")

from paddleocr import PaddleOCR
from PIL import Image, ImageDraw

# Create dummy image
filename = "test_repro.png"
if not os.path.exists(filename):
    img = Image.new('RGB', (100, 100), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.text((10, 10), "Test", fill=(0, 0, 0))
    img.save(filename)

try:
    print("Initializing PaddleOCR (use_angle_cls=False)...")
    # Disable angle cls
    ocr = PaddleOCR(use_angle_cls=False, lang='en')
    
    print("Running OCR with predict()...")
    result = ocr.predict(filename) 
    print("OCR result type:", type(result))
    print("OCR result:", result)

except Exception as e:
    import traceback
    traceback.print_exc()
