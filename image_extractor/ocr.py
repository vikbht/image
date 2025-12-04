import re
from pathlib import Path
from .utils import sanitize_filename

# Global variable to hold the EasyOCR reader instance
easyocr_reader = None

def get_ocr_reader():
    """
    Initializes and returns the EasyOCR reader.
    This function is designed to be called in a threaded context.
    """
    global easyocr_reader
    if easyocr_reader is None:
        import easyocr
        import logging
        logging.getLogger('easyocr').setLevel(logging.ERROR)
        # Try to use GPU (MPS on Mac) for speed
        try:
            easyocr_reader = easyocr.Reader(['en'], gpu=True)
        except Exception:
            print("Warning: GPU initialization failed, falling back to CPU")
            easyocr_reader = easyocr.Reader(['en'], gpu=False)
    return easyocr_reader

def init_worker():
    """Initialize the EasyOCR reader in the worker process (if using multiprocessing)."""
    # Kept for compatibility if we switch back to multiprocessing
    get_ocr_reader()

def process_image_threaded(image_path: Path, lock) -> tuple:
    """Worker function for threaded OCR."""
    try:
        reader = get_ocr_reader()
            
        # Optimization: Resize image logic (Parallel)
        import numpy as np
        from PIL import Image
        
        with Image.open(image_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Aggressive resizing for speed
            max_dimension = 500
            width, height = img.size
            if width > max_dimension or height > max_dimension:
                ratio = min(max_dimension / width, max_dimension / height)
                new_size = (int(width * ratio), int(height * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            img_np = np.array(img)
            
        # Serialized Inference
        with lock:
            result = reader.readtext(img_np, detail=0)
        
        if not result:
            return None
            
        full_text = " ".join(result)
        
        # Sanitize logic
        text = full_text
        text = re.sub(r'[<>:"/\\|?*]', '', text)
        text = re.sub(r'\s+', '_', text)
        text = re.sub(r'[^\w\-_]', '', text)
        text = text.strip('_')
        if not text: text = "untitled"
        new_name = text[:50]
        
        return (image_path, new_name, full_text)
        
    except Exception as e:
        return None

def perform_ocr_and_rename(image_path: Path) -> Path:
    """
    Perform OCR on an image and rename it based on the content.
    (Sequential version, kept for fallback)
    """
    try:
        reader = get_ocr_reader()
        
        # Optimization: Resize image for faster OCR
        import numpy as np
        from PIL import Image
        
        with Image.open(image_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            max_dimension = 500
            width, height = img.size
            if width > max_dimension or height > max_dimension:
                ratio = min(max_dimension / width, max_dimension / height)
                new_size = (int(width * ratio), int(height * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            img_np = np.array(img)
        
        result = reader.readtext(img_np, detail=0)
        
        if not result:
            print(f"  No text found in {image_path.name}")
            return image_path
            
        full_text = " ".join(result)
        new_name = sanitize_filename(full_text)
        
        if new_name == "untitled":
            return image_path
            
        new_path = image_path.with_name(f"{new_name}{image_path.suffix}")
        
        counter = 1
        while new_path.exists() and new_path != image_path:
            new_path = image_path.with_name(f"{new_name}_{counter}{image_path.suffix}")
            counter += 1
            
        image_path.rename(new_path)
        print(f"  ✓ Renamed: {image_path.name} -> {new_path.name}")
        print(f"    (Text: '{full_text[:60]}...')")
        
        return new_path
        
    except Exception as e:
        print(f"  ⚠ OCR failed for {image_path.name}: {e}")
        return image_path
