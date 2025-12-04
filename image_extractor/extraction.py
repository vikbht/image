import os
import zipfile
import shutil
from pathlib import Path
from typing import List

# Magic numbers for file type identification
MAGIC_NUMBERS = {
    b'\x89PNG\r\n\x1a\n': '.png',
    b'\xff\xd8\xff': '.jpg',
    b'GIF87a': '.gif',
    b'GIF89a': '.gif',
    b'BM': '.bmp',
    b'MM\x00*': '.tiff',
    b'II*\x00': '.tiff',
    b'\x01\x00\x00\x00': '.emf',  # EMF often starts with this
    b'\x00\x00\x01\x00': '.ico',
}

def get_file_extension(file_path: Path) -> str:
    """
    Determine the file extension based on magic numbers.
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(32)
            
        for magic, ext in MAGIC_NUMBERS.items():
            if header.startswith(magic):
                return ext
                
        # Check for EMF specifically if magic number check failed
        # EMF files are complex, sometimes the header varies slightly or offset
        if file_path.suffix.lower() == '.emf' or file_path.suffix.lower() == '.wmf':
            return '.emf'
            
        return file_path.suffix.lower()
    except Exception:
        return file_path.suffix.lower()

def extract_images_from_docx(docx_path: Path, output_dir: Path) -> List[Path]:
    """
    Extract images from a .docx file (which is a zip archive).
    """
    extracted_files = []
    
    if not output_dir.exists():
        output_dir.mkdir(parents=True)
        
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # Look for files in word/media/
            media_files = [f for f in zip_ref.namelist() if f.startswith('word/media/')]
            
            if not media_files:
                print("No images found in the document.")
                return []
                
            print(f"Found {len(media_files)} image(s) in the document.")
            
            for file in media_files:
                # Extract to temporary location
                zip_ref.extract(file, output_dir)
                
                # Move to root of output dir and rename
                source_path = output_dir / file
                filename = Path(file).name
                target_path = output_dir / filename
                
                # Move file
                shutil.move(str(source_path), str(target_path))
                
                # Determine correct extension
                ext = get_file_extension(target_path)
                final_path = target_path.with_suffix(ext)
                
                if target_path != final_path:
                    target_path.rename(final_path)
                    print(f"  Extracted: {filename} -> {final_path.name}")
                else:
                    print(f"  Extracted: {final_path.name}")
                    
                extracted_files.append(final_path)
                
            # Cleanup empty directories
            shutil.rmtree(output_dir / 'word', ignore_errors=True)
            
    except zipfile.BadZipFile:
        print(f"Error: {docx_path} is not a valid .docx file.")
    except Exception as e:
        print(f"Error extracting images: {e}")
        
    return extracted_files
