# Word Document Image Extractor

A Python script to extract all images from Microsoft Word documents (.docx) and optionally convert EMF files to PNG format.

## Features

- ✅ Extracts all images from .docx files
- ✅ Automatically identifies file types (PNG, JPG, EMF, etc.)
- ✅ Renames files with correct extensions
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

### Basic Usage

Extract all images from a Word document:

```bash
python3 extract_word_images.py document.docx
```

This will create an `extracted_images` folder containing all images.

### Specify Output Directory

```bash
python3 extract_word_images.py document.docx -o my_images
```

### Extract Without Converting EMF Files

If you just want to extract the files without attempting conversion:

```bash
python3 extract_word_images.py document.docx --no-convert
```

### Help

```bash
python3 extract_word_images.py --help
```

## How It Works

1. **Extraction**: Word documents (.docx) are actually ZIP archives. The script:
   - Unzips the document
   - Finds all files in the `word/media/` folder
   - Extracts them to the output directory

2. **File Type Detection**: The script reads the file signature (magic bytes) to identify the actual file type, regardless of the extension.

3. **Renaming**: Files are renamed with the correct extension (.png, .jpg, .emf, etc.)

4. **EMF Conversion** (optional): The script attempts to convert EMF files to PNG using:
   - ImageMagick (if installed)
   - Pillow with pillow-emf extension (if available)

## Example

```bash
$ python3 extract_word_images.py o1.docx
Processing: o1.docx
Output directory: extracted_images
------------------------------------------------------------
Found 9 image(s) in the document.
  Extracted: image8.tmp -> image8.emf
  Extracted: image1.png
  Extracted: image2.tmp -> image2.emf
  ...

Extracted 9 file(s)

Converting EMF files to PNG...
  ✓ Converted: image2.emf -> image2.png (using ImageMagick)
  ...

============================================================
✓ Processing complete! Images saved to: extracted_images
============================================================
```

## Troubleshooting

### EMF Files Won't Convert

EMF is a Windows-specific format. If conversion fails:

1. **Use LibreOffice** (best option for macOS):
   ```bash
   brew install --cask libreoffice
   brew install imagemagick
   ```

2. **Use an online converter**: Upload the extracted .emf files to https://convertio.co/emf-png/

3. **Use Windows**: Open the .emf files on a Windows machine and save as PNG

### "Command not found: python"

Use `python3` instead:
```bash
python3 extract_word_images.py document.docx
```

### Permission Denied

Make the script executable:
```bash
chmod +x extract_word_images.py
./extract_word_images.py document.docx
```

## Supported Image Formats

The script can extract and identify:
- PNG
- JPEG/JPG
- GIF
- BMP
- EMF (Enhanced Metafile)
- ICO
- And other formats embedded in Word documents

## License

Free to use and modify.

## Author

Created to simplify image extraction from Word documents.
