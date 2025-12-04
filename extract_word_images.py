#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "aspose-words",
#     "pillow",
#     "easyocr",
# ]
# ///
"""
Extract images from Word documents (.docx) and convert EMF files to PNG.
Also renames images based on text content found within them using OCR.
"""

import os
import sys
import zipfile
import shutil
import re
import time
from pathlib import Path
from typing import List, Tuple
import argparse

# Lazy load easyocr to avoid overhead if not used or during initial imports
easyocr_reader = None

def get_ocr_reader():
    """Get or initialize the global EasyOCR reader."""
    global easyocr_reader
    if easyocr_reader is None:
        import easyocr
        import logging
        # Suppress EasyOCR warnings
        logging.getLogger('easyocr').setLevel(logging.ERROR)
        print("Initializing OCR engine (this may take a moment)...")
        # Use CPU to avoid potential GPU issues on some systems, unless configured otherwise
        easyocr_reader = easyocr.Reader(['en'], gpu=False)
    return easyocr_reader

def sanitize_filename(text: str, max_length: int = 50) -> str:
    """
    Sanitize text to be safe for filenames.
    
    Args:
        text: Input text
        max_length: Maximum length of the filename
        
    Returns:
        Sanitized filename string
    """
    # Remove invalid characters
    text = re.sub(r'[<>:"/\\|?*]', '', text)
    # Replace spaces and non-alphanumeric chars with underscores
    text = re.sub(r'\s+', '_', text)
    text = re.sub(r'[^\w\-_]', '', text)
    # Remove leading/trailing underscores
    text = text.strip('_')
    
    if not text:
        return "untitled"
        
    return text[:max_length]

def perform_ocr_and_rename(image_path: Path) -> Path:
    """
    Perform OCR on an image and rename it based on the content.
    
    Args:
        image_path: Path to the image file
        
    Returns:
        Path to the renamed file (or original if failed)
    """
    try:
        reader = get_ocr_reader()
        
        # Read text from image
        result = reader.readtext(str(image_path), detail=0)
        
        if not result:
            print(f"  No text found in {image_path.name}")
            return image_path
            
        # Use the first few words as the title
        # Join first few results, but limit total length
        full_text = " ".join(result)
        
        # Sanitize to create a valid filename
        new_name = sanitize_filename(full_text)
        
        if new_name == "untitled":
            return image_path
            
        # Construct new path
        new_path = image_path.with_name(f"{new_name}{image_path.suffix}")
        
        # Handle duplicates
        counter = 1
        while new_path.exists() and new_path != image_path:
            new_path = image_path.with_name(f"{new_name}_{counter}{image_path.suffix}")
            counter += 1
            
        # Rename the file
        image_path.rename(new_path)
        print(f"  ✓ Renamed: {image_path.name} -> {new_path.name}")
        print(f"    (Text: '{full_text[:60]}...')")
        
        return new_path
        
    except Exception as e:
        print(f"  ⚠ OCR failed for {image_path.name}: {e}")
        return image_path

def identify_file_type(file_path: Path) -> str:
    """
    Identify the file type by reading the file signature (magic bytes).
    
    Args:
        file_path: Path to the file to identify
        
    Returns:
        File extension (e.g., 'png', 'jpg', 'emf', 'unknown')
    """
    with open(file_path, 'rb') as f:
        header = f.read(16)
    
    # Check for common image formats
    if header.startswith(b'\x89PNG\r\n\x1a\n'):
        return 'png'
    elif header.startswith(b'\xff\xd8\xff'):
        return 'jpg'
    elif header.startswith(b'GIF87a') or header.startswith(b'GIF89a'):
        return 'gif'
    elif header.startswith(b'BM'):
        return 'bmp'
    elif header.startswith(b'\x01\x00\x00\x00'):
        # EMF files start with this signature
        return 'emf'
    elif header[0:4] == b'\x00\x00\x01\x00':
        return 'ico'
    else:
        return 'unknown'


def extract_images_from_docx(docx_path: Path, output_dir: Path) -> List[Path]:
    """
    Extract all images from a Word document.
    
    Args:
        docx_path: Path to the .docx file
        output_dir: Directory to save extracted images
        
    Returns:
        List of paths to extracted image files
    """
    if not docx_path.exists():
        raise FileNotFoundError(f"Word document not found: {docx_path}")
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    extracted_files = []
    
    # Word documents are ZIP archives
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        # Find all files in the media folder
        media_files = [f for f in zip_ref.namelist() if f.startswith('word/media/')]
        
        if not media_files:
            print("No images found in the document.")
            return []
        
        print(f"Found {len(media_files)} image(s) in the document.")
        
        # Extract each media file
        for media_file in media_files:
            # Get the filename
            filename = os.path.basename(media_file)
            output_path = output_dir / filename
            
            # Extract the file
            with zip_ref.open(media_file) as source, open(output_path, 'wb') as target:
                shutil.copyfileobj(source, target)
            
            # Identify the actual file type
            file_type = identify_file_type(output_path)
            
            # Rename with correct extension if needed
            if file_type != 'unknown' and not filename.endswith(f'.{file_type}'):
                new_path = output_path.with_suffix(f'.{file_type}')
                output_path.rename(new_path)
                output_path = new_path
                print(f"  Extracted: {filename} -> {output_path.name}")
            else:
                print(f"  Extracted: {output_path.name}")
            
            extracted_files.append(output_path)
    
    return extracted_files


def convert_emf_to_png(emf_path: Path, png_path: Path = None) -> Path:
    """
    Convert an EMF file to PNG format using available tools.
    
    EMF (Enhanced Metafile) is a Windows-specific format that's challenging to convert on macOS.
    This function tries multiple approaches in order of reliability.
    
    Args:
        emf_path: Path to the EMF file
        png_path: Optional output path for PNG file (defaults to same name with .png extension)
        
    Returns:
        Path to the converted PNG file, or None if conversion failed
    """
    if png_path is None:
        png_path = emf_path.with_suffix('.png')
    
    import subprocess
    
    # Add common Homebrew paths to environment
    env = os.environ.copy()
    if 'PATH' in env:
        env['PATH'] = f"/opt/homebrew/bin:/usr/local/bin:{env['PATH']}"
    else:
        env['PATH'] = "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"
    
    # Method 1: Try Aspose.Words (Python-only solution)
    try:
        import aspose.words as aw
        
        # Suppress Aspose logging if possible
        # aw.utils.logging.set_logger_level(aw.utils.logging.LoggerLevel.OFF)
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        
        # Insert the EMF image
        shape = builder.insert_image(str(emf_path))
        
        # Save the shape as PNG
        shape.get_shape_renderer().save(str(png_path), aw.saving.ImageSaveOptions(aw.SaveFormat.PNG))
        
        if png_path.exists() and png_path.stat().st_size > 0:
            print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using Aspose.Words)")
            return png_path
            
    except ImportError:
        pass  # Aspose.Words not installed
    except Exception as e:
        # print(f"Debug: Aspose failed: {e}")
        pass

    # Method 2: Try direct LibreOffice on macOS (most reliable for EMF)
    libreoffice_path = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    if libreoffice_path.exists():
        try:
            # LibreOffice requires an output directory, not a specific output file
            # It will save the file with the same basename but new extension
            out_dir = png_path.parent
            
            result = subprocess.run(
                [str(libreoffice_path), '--headless', '--convert-to', 'png', '--outdir', str(out_dir), str(emf_path)],
                capture_output=True,
                timeout=60,
                env=env
            )
            
            # Check if the expected output file exists
            # LibreOffice might name it slightly differently or just change extension
            expected_output = out_dir / emf_path.with_suffix('.png').name
            
            if result.returncode == 0 and expected_output.exists() and expected_output.stat().st_size > 0:
                print(f"  ✓ Converted: {emf_path.name} -> {expected_output.name} (using LibreOffice)")
                return expected_output
        except Exception as e:
            # print(f"Debug: LibreOffice failed: {e}")
            pass

    # Method 3: Try ImageMagick with LibreOffice delegate (if available)
    for cmd_base in ['magick', 'convert']:
        try:
            result = subprocess.run(
                [cmd_base, str(emf_path), str(png_path)],
                capture_output=True,
                timeout=30,
                env=env
            )
            if result.returncode == 0 and png_path.exists() and png_path.stat().st_size > 0:
                print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using ImageMagick)")
                return png_path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
        except Exception:
            continue
    
    # Method 4: Try Pillow with pillow-emf extension
    try:
        from PIL import Image
        from PIL import ImageFile
        ImageFile.LOAD_TRUNCATED_IMAGES = True
        
        # Try to import pillow-emf for EMF support
        try:
            import pillow_emf
            has_emf_support = True
        except ImportError:
            has_emf_support = False
        
        if has_emf_support:
            # Open and convert the EMF file
            with Image.open(emf_path) as img:
                # Convert to RGB if necessary
                if img.mode not in ('RGB', 'RGBA'):
                    img = img.convert('RGB')
                
                # Save as PNG
                img.save(png_path, 'PNG')
            
            if png_path.exists() and png_path.stat().st_size > 0:
                print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using Pillow+EMF)")
                return png_path
        
    except ImportError:
        pass  # Pillow not installed
    except Exception:
        pass  # Conversion failed
    
    # All methods failed
    print(f"  ⚠ Could not convert {emf_path.name} (EMF format)")
    print(f"    EMF is a Windows format. The file has been extracted but not converted.")
    print(f"    Options:")
    print(f"      1. Install Aspose.Words: pip install aspose-words")
    print(f"      2. Install LibreOffice: brew install --cask libreoffice")
    print(f"      3. Use online converter: https://convertio.co/emf-png/")
    return None


def process_word_document(docx_path: str, output_dir: str = None, convert_emf: bool = True, rename_images: bool = True):
    """
    Main function to extract and process images from a Word document.
    
    Args:
        docx_path: Path to the Word document
        output_dir: Output directory for extracted images (default: extracted_images)
        convert_emf: Whether to convert EMF files to PNG (default: True)
        rename_images: Whether to rename images based on OCR content (default: True)
    """
    start_time = time.time()
    docx_path = Path(docx_path)
    
    if output_dir is None:
        output_dir = docx_path.parent / "extracted_images"
    else:
        output_dir = Path(output_dir)
    
    print(f"Processing: {docx_path}")
    print(f"Output directory: {output_dir}")
    print("-" * 60)
    
    # Extract images
    extracted_files = extract_images_from_docx(docx_path, output_dir)
    
    if not extracted_files:
        return
    
    print(f"\nExtracted {len(extracted_files)} file(s)")
    
    # List of final images to potentially rename
    final_images = []
    
    # Convert EMF files if requested
    if convert_emf:
        print("\nConverting EMF files to PNG...")
        emf_files = [f for f in extracted_files if f.suffix.lower() == '.emf']
        
        if emf_files:
            converted_count = 0
            for emf_file in emf_files:
                result = convert_emf_to_png(emf_file)
                if result:
                    converted_count += 1
                    final_images.append(result)
                else:
                    final_images.append(emf_file)
            
            print(f"\nConverted {converted_count}/{len(emf_files)} EMF file(s) to PNG")
        else:
            print("No EMF files found to convert.")
            final_images.extend(extracted_files)
    else:
        final_images.extend(extracted_files)
        
    # Add non-EMF files that weren't processed above
    for f in extracted_files:
        if f.suffix.lower() != '.emf' and f not in final_images:
            final_images.append(f)
            
    # Rename images using OCR
    if rename_images:
        print("\nAnalyzing images and renaming based on content...")
        print("(This requires OCR and may take some time)")
        
        renamed_count = 0
        for img_path in final_images:
            # Skip if file doesn't exist (e.g. was already renamed or deleted)
            if not img_path.exists():
                continue
                
            # Only process common image formats
            if img_path.suffix.lower() not in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff']:
                continue
                
            new_path = perform_ocr_and_rename(img_path)
            if new_path != img_path:
                renamed_count += 1
                
        print(f"\nRenamed {renamed_count} file(s) based on content")
    
    elapsed_time = time.time() - start_time
    print("\n" + "=" * 60)
    print(f"✓ Processing complete! Images saved to: {output_dir}")
    print(f"  Time taken: {elapsed_time:.2f} seconds")
    print("=" * 60)


def main():
    """Command-line interface for the script."""
    parser = argparse.ArgumentParser(
        description="Extract images from Word documents, convert EMF to PNG, and rename based on content",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx -o my_images
  %(prog)s document.docx --no-ocr
        """
    )
    
    parser.add_argument(
        'docx_file',
        help='Path to the Word document (.docx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        dest='output_dir',
        help='Output directory for extracted images (default: extracted_images)',
        default=None
    )
    
    parser.add_argument(
        '--no-convert',
        dest='convert_emf',
        action='store_false',
        help='Do not convert EMF files to PNG'
    )
    
    parser.add_argument(
        '--no-ocr',
        dest='rename_images',
        action='store_false',
        help='Do not rename images based on OCR content'
    )
    
    args = parser.parse_args()
    
    try:
        process_word_document(
            args.docx_file, 
            args.output_dir, 
            args.convert_emf,
            args.rename_images
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
