import sys
import time
import argparse
import concurrent.futures
import threading
from pathlib import Path

from .extraction import extract_images_from_docx
from .conversion import convert_emf_to_png
from .ocr import perform_ocr_and_rename, get_ocr_reader, process_image_threaded

def process_word_document(docx_path: str, output_dir: str = None, convert_emf: bool = True, rename_images: bool = True):
    """
    Main function to process a Word document.
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
        print("\nConverting EMF files to PNG (Parallel)...")
        emf_files = [f for f in extracted_files if f.suffix.lower() == '.emf']
        
        if emf_files:
            converted_count = 0
            # Use ThreadPoolExecutor for parallel conversion
            with concurrent.futures.ThreadPoolExecutor() as executor:
                # Submit all conversion tasks
                future_to_emf = {executor.submit(convert_emf_to_png, emf): emf for emf in emf_files}
                
                for future in concurrent.futures.as_completed(future_to_emf):
                    emf_file = future_to_emf[future]
                    try:
                        result = future.result()
                        if result:
                            converted_count += 1
                            final_images.append(result)
                        else:
                            final_images.append(emf_file)
                    except Exception as exc:
                        print(f"  ⚠ Conversion generated an exception for {emf_file.name}: {exc}")
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
        print("\nAnalyzing images and renaming based on content (Parallel I/O)...")
        print("(This requires OCR and may take some time)")
        
        # Filter images eligible for OCR
        images_to_process = []
        for img_path in final_images:
            if img_path.exists() and img_path.suffix.lower() in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff']:
                images_to_process.append(img_path)
        
        renamed_count = 0
        if images_to_process:
            # Initialize reader once
            try:
                get_ocr_reader()
            except Exception as e:
                print(f"Warning: Could not initialize OCR engine: {e}")
                return

            ocr_lock = threading.Lock()
            
            with concurrent.futures.ThreadPoolExecutor() as executor:
                # Submit tasks
                future_to_img = {
                    executor.submit(process_image_threaded, img, ocr_lock): img 
                    for img in images_to_process
                }
                
                for future in concurrent.futures.as_completed(future_to_img):
                    try:
                        result = future.result()
                        if result:
                            original_path, new_name, full_text = result
                            
                            if new_name == "untitled":
                                continue
                                
                            # Construct new path
                            new_path = original_path.with_name(f"{new_name}{original_path.suffix}")
                            
                            # Handle duplicates
                            counter = 1
                            while new_path.exists() and new_path != original_path:
                                new_path = original_path.with_name(f"{new_name}_{counter}{original_path.suffix}")
                                counter += 1
                                
                            original_path.rename(new_path)
                            print(f"  ✓ Renamed: {original_path.name} -> {new_path.name}")
                            print(f"    (Text: '{full_text[:60]}...')")
                            renamed_count += 1
                    except Exception as exc:
                        print(f"  ⚠ OCR generated an exception: {exc}")
                
        print(f"\nRenamed {renamed_count} file(s) based on content")
    
    elapsed_time = time.time() - start_time
    print("\n" + "=" * 60)
    print(f"✓ Processing complete! Images saved to: {output_dir}")
    print(f"  Time taken: {elapsed_time:.2f} seconds")
    print("=" * 60)

def main():
    """Command-line interface for the script."""
    parser = argparse.ArgumentParser(description="Extract images from a Word document (.docx)")
    parser.add_argument("docx_file", help="Path to the .docx file")
    parser.add_argument("-o", "--output", help="Output directory (default: extracted_images)")
    parser.add_argument("--no-convert", action="store_true", help="Do not convert EMF files to PNG")
    parser.add_argument("--no-ocr", action="store_true", help="Do not rename images based on OCR content")
    
    args = parser.parse_args()
    
    docx_path = Path(args.docx_file)
    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
        
    process_word_document(
        docx_path, 
        args.output, 
        convert_emf=not args.no_convert,
        rename_images=not args.no_ocr
    )

if __name__ == "__main__":
    main()
