# Word Document Image Extractor

A Python script to extract all images from Microsoft Word documents (.docx) and optionally convert EMF files to PNG format.

## Features

- ✅ Extracts all images from .docx files
- ✅ Automatically identifies file types (PNG, JPG, EMF, etc.)
- ✅ Renames files with correct extensions
- ✅ **Automatically renames images based on content** (using OCR)
- ✅ Attempts to convert EMF files to PNG (when possible)
- ✅ No external dependencies required for basic extraction
- ✅ Command-line interface with options

## Installation & Usage with uv (Recommended)

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

### 1. Install uv
```bash
# On macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. Run the script
You can run the script directly with `uv run`. It will automatically handle dependencies.

```bash
uv run extract_word_images.py o1.docx
```

### 3. Project Setup (Optional)
If you want to work on the project or install dependencies in a virtual environment:

```bash
uv sync
```

## Manual Installation (Legacy)

If you prefer not to use `uv`, you can install dependencies manually:

```bash
pip install aspose-words pillow
```

## Usage

### Basic Usage (with uv)

Extract all images from a Word document, convert EMFs, and rename based on content:

```bash
uv run extract_word_images.py document.docx
```

This will create an `extracted_images` folder containing all images.

### Specify Output Directory

```bash
uv run extract_word_images.py document.docx -o my_images
```

### Disable OCR Renaming

If you want to skip the OCR process (faster):

```bash
uv run extract_word_images.py document.docx --no-ocr
```

### Extract Without Converting EMF Files

If you just want to extract the files without attempting conversion:

```bash
uv run extract_word_images.py document.docx --no-convert
```

### Help

```bash
uv run extract_word_images.py --help
```

## How It Works

1. **Extraction**: Word documents (.docx) are actually ZIP archives. The script:
   - Unzips the document
   - Finds all files in the `word/media/` folder
   - Extracts them to the output directory

2. **File Type Detection**: The script reads the file signature (magic bytes) to identify the actual file type, regardless of the extension.

3. **EMF Conversion**: The script attempts to convert EMF files to PNG using:
   - **Aspose.Words** (primary, pure Python)
   - LibreOffice (if installed on macOS)
   - ImageMagick (fallback)
   - Pillow (fallback)

4. **OCR Renaming**: The script uses **EasyOCR** to:
   - Read text from each image
   - Generate a descriptive filename (e.g., `Biceps_brachii_Site_4.png`)
   - Rename the file automatically

## Example

```bash
$ uv run extract_word_images.py o1.docx
Processing: o1.docx
Output directory: extracted_images
------------------------------------------------------------
Found 9 image(s) in the document.
  Extracted: image8.tmp -> image8.emf
  Extracted: image1.png
  ...

Extracted 9 file(s)

Converting EMF files to PNG...
  ✓ Converted: image2.emf -> image2.png (using Aspose.Words)
  ...

Analyzing images and renaming based on content...
(This requires OCR and may take some time)
  ✓ Renamed: image5.png -> Biceps_brachii_Site_4.png
    (Text: 'Biceps brachii Site 4...')
  ✓ Renamed: image6.png -> Motor_NCS_L_Median.png
    (Text: 'Motor NCS L Median...')

Renamed 9 file(s) based on content

============================================================
✓ Processing complete! Images saved to: extracted_images
============================================================
```

## Troubleshooting

### OCR Slowness
OCR is computationally intensive. It may take a few seconds per image. If you have many images and don't need descriptive names, use `--no-ocr`.

### EMF Files Won't Convert
EMF is a Windows-specific format. The script tries multiple methods. If all fail:
1. Ensure `aspose-words` is installed (`uv sync`)
2. Install LibreOffice: `brew install --cask libreoffice`
3. Use an online converter: https://convertio.co/emf-png/

## Supported Image Formats

The script can extract and identify:
- PNG, JPEG/JPG, GIF, BMP, TIFF
- EMF (Enhanced Metafile)
- ICO
- And other formats embedded in Word documents

## License

Free to use and modify.

## Author

Created to simplify image extraction from Word documents.
