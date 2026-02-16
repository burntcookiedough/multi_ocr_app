# ocr_engines/paddle_engine.py

import paddle
from paddleocr import PaddleOCR

# ✅ CRITICAL FIX: disable oneDNN
# paddle.set_flags({'FLAGS_use_onednn': False})

# ✅ Initialize once
ocr = PaddleOCR(
    use_angle_cls=True,
    lang='en'
)

def extract_text(image_path):
    try:
        result = ocr.ocr(image_path)

        extracted_text = []
        if result and result[0]:
            for line in result[0]:
                extracted_text.append(line[1][0])

        return "\n".join(extracted_text)

    except Exception as e:
        return f"Error in PaddleOCR: {str(e)}"
