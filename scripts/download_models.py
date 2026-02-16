import os
import sys

# Add project root to path to ensure we can import if needed, though we just need packages
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def download_paddle():
    print("-" * 30)
    print("Checking PaddleOCR Models...")
    try:
        from paddleocr import PaddleOCR
        # Initializing triggers download if missing
        ocr = PaddleOCR(use_angle_cls=True, lang='en', show_log=False)
        print("✅ PaddleOCR models are ready.")
    except Exception as e:
        print(f"❌ Error setting up PaddleOCR: {e}")

def download_easyocr():
    print("-" * 30)
    print("Checking EasyOCR Models...")
    try:
        import easyocr
        # Initializing triggers download if missing
        reader = easyocr.Reader(['en'], gpu=False)
        print("✅ EasyOCR models are ready.")
    except Exception as e:
        print(f"❌ Error setting up EasyOCR: {e}")

if __name__ == "__main__":
    print("Starting Model Verification/Download...")
    download_paddle()
    download_easyocr()
    print("-" * 30)
    print("Model setup complete.")
