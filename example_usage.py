#!/usr/bin/env python3
"""
Example usage of the extract_word_images module.

This demonstrates how to use the image extraction functionality
programmatically in your own Python scripts.
"""

from pathlib import Path
from extract_word_images import extract_images_from_docx, convert_emf_to_png

def example_basic_extraction():
    """Basic example: Extract images without conversion."""
    print("=" * 60)
    print("Example 1: Basic Image Extraction")
    print("=" * 60)
    
    docx_path = Path("o1.docx")
    output_dir = Path("example_output_1")
    
    # Extract all images
    extracted_files = extract_images_from_docx(docx_path, output_dir)
    
    print(f"\nExtracted {len(extracted_files)} images to {output_dir}")
    for file in extracted_files:
        print(f"  - {file.name} ({file.stat().st_size:,} bytes)")


def example_with_conversion():
    """Example: Extract and convert EMF files."""
    print("\n" + "=" * 60)
    print("Example 2: Extract and Convert EMF Files")
    print("=" * 60)
    
    docx_path = Path("o1.docx")
    output_dir = Path("example_output_2")
    
    # Extract all images
    extracted_files = extract_images_from_docx(docx_path, output_dir)
    
    # Find and convert EMF files
    emf_files = [f for f in extracted_files if f.suffix.lower() == '.emf']
    
    if emf_files:
        print(f"\nFound {len(emf_files)} EMF file(s). Attempting conversion...")
        for emf_file in emf_files:
            png_file = convert_emf_to_png(emf_file)
            if png_file:
                print(f"  ✓ Successfully converted: {png_file.name}")
            else:
                print(f"  ✗ Could not convert: {emf_file.name}")
    else:
        print("\nNo EMF files found.")


def example_filter_by_type():
    """Example: Extract and organize images by type."""
    print("\n" + "=" * 60)
    print("Example 3: Organize Images by Type")
    print("=" * 60)
    
    docx_path = Path("o1.docx")
    output_dir = Path("example_output_3")
    
    # Extract all images
    extracted_files = extract_images_from_docx(docx_path, output_dir)
    
    # Organize by file type
    by_type = {}
    for file in extracted_files:
        ext = file.suffix.lower()
        if ext not in by_type:
            by_type[ext] = []
        by_type[ext].append(file)
    
    print(f"\nExtracted images organized by type:")
    for ext, files in sorted(by_type.items()):
        total_size = sum(f.stat().st_size for f in files)
        print(f"\n  {ext.upper()} files ({len(files)} total, {total_size:,} bytes):")
        for file in files:
            print(f"    - {file.name}")


def example_custom_processing():
    """Example: Extract and perform custom processing."""
    print("\n" + "=" * 60)
    print("Example 4: Custom Processing")
    print("=" * 60)
    
    docx_path = Path("o1.docx")
    output_dir = Path("example_output_4")
    
    # Extract all images
    extracted_files = extract_images_from_docx(docx_path, output_dir)
    
    # Custom processing: Create a report
    report_path = output_dir / "image_report.txt"
    
    with open(report_path, 'w') as f:
        f.write(f"Image Extraction Report\n")
        f.write(f"Source: {docx_path}\n")
        f.write(f"Date: {Path(__file__).stat().st_mtime}\n")
        f.write(f"\n{'=' * 60}\n\n")
        
        for i, file in enumerate(extracted_files, 1):
            f.write(f"{i}. {file.name}\n")
            f.write(f"   Type: {file.suffix}\n")
            f.write(f"   Size: {file.stat().st_size:,} bytes\n")
            f.write(f"   Path: {file}\n\n")
    
    print(f"\nCreated report: {report_path}")
    print(f"Extracted {len(extracted_files)} images")


if __name__ == "__main__":
    # Run all examples
    try:
        example_basic_extraction()
        example_with_conversion()
        example_filter_by_type()
        example_custom_processing()
        
        print("\n" + "=" * 60)
        print("All examples completed!")
        print("=" * 60)
        
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        print("Make sure 'o1.docx' exists in the current directory.")
    except Exception as e:
        print(f"\nError: {e}")
