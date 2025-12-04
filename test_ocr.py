import easyocr
from pathlib import Path
import sys
import logging

# Suppress EasyOCR warnings
logging.getLogger('easyocr').setLevel(logging.ERROR)

def test_ocr(image_path):
    try:
        print(f"Initializing EasyOCR for {image_path}...")
        # Initialize reader for English
        reader = easyocr.Reader(['en'], gpu=False) # Use CPU for compatibility
        
        print("Reading text...")
        # Read text
        result = reader.readtext(str(image_path), detail=0)
        
        # Join all detected text
        full_text = " ".join(result)
        
        print(f"OCR Result for {image_path}:")
        print("-" * 40)
        print(full_text)
        print("-" * 40)
        return full_text
        
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) > 1:
        test_ocr(Path(sys.argv[1]))
    else:
        print("Usage: uv run test_ocr.py <image_file>")
