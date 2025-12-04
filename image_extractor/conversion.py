import os
import shutil
import subprocess
from pathlib import Path

def convert_emf_to_png(emf_path: Path) -> Path:
    """
    Convert an EMF file to PNG using available tools.
    Prioritizes Aspose.Words (Python), then LibreOffice, then ImageMagick.
    """
    png_path = emf_path.with_suffix('.png')
    
    # Method 1: Aspose.Words (Best Python-only solution)
    try:
        import aspose.words as aw
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        shape = builder.insert_image(str(emf_path))
        shape.get_shape_renderer().save(str(png_path), aw.saving.ImageSaveOptions(aw.SaveFormat.PNG))
        print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using Aspose.Words)")
        return png_path
    except ImportError:
        pass # Try next method
    except Exception as e:
        print(f"  ⚠ Aspose.Words conversion failed: {e}")

    # Method 2: LibreOffice (Best for macOS/Linux if installed)
    # Check for soffice or libreoffice command
    lo_cmd = None
    possible_cmds = ['soffice', 'libreoffice']
    
    # On macOS, check standard path if not in PATH
    if os.uname().sysname == 'Darwin':
        mac_lo_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
        if os.path.exists(mac_lo_path):
            lo_cmd = mac_lo_path
            
    if not lo_cmd:
        for cmd in possible_cmds:
            if shutil.which(cmd):
                lo_cmd = cmd
                break
                
    if lo_cmd:
        try:
            # LibreOffice headless convert
            # --headless --convert-to png --outdir <dir> <file>
            cmd = [
                lo_cmd, 
                '--headless', 
                '--convert-to', 'png', 
                '--outdir', str(emf_path.parent), 
                str(emf_path)
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            if png_path.exists():
                print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using LibreOffice)")
                return png_path
        except Exception as e:
            print(f"  ⚠ LibreOffice conversion failed: {e}")

    # Method 3: ImageMagick
    # Check for 'magick' (v7) or 'convert' (v6)
    im_cmd = None
    if shutil.which('magick'):
        im_cmd = ['magick', str(emf_path), str(png_path)]
    elif shutil.which('convert'):
        im_cmd = ['convert', str(emf_path), str(png_path)]
        
    if im_cmd:
        try:
            # On macOS with Homebrew, we might need to set DYLD_LIBRARY_PATH or similar
            # but usually it works if installed correctly.
            subprocess.run(im_cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if png_path.exists():
                print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using ImageMagick)")
                return png_path
        except Exception as e:
            print(f"  ⚠ ImageMagick conversion failed: {e}")
            
    # Method 4: Pillow (Fallback, limited EMF support)
    try:
        from PIL import Image
        with Image.open(emf_path) as img:
            img.save(png_path)
        print(f"  ✓ Converted: {emf_path.name} -> {png_path.name} (using Pillow)")
        return png_path
    except Exception:
        pass

    print(f"  ❌ Failed to convert: {emf_path.name}")
    return None
