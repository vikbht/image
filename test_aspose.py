import aspose.words as aw

def convert_emf_to_png(emf_path, png_path):
    try:
        # Load the EMF file into a Document object
        # Note: Aspose.Words usually expects a Word document, but can sometimes load images directly
        # or we might need to insert it into a blank document first
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        
        # Insert the EMF image
        shape = builder.insert_image(str(emf_path))
        
        # Save the shape as PNG
        shape.get_shape_renderer().save(str(png_path), aw.saving.ImageSaveOptions(aw.SaveFormat.PNG))
        
        print(f"Converted {emf_path} to {png_path}")
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    import sys
    from pathlib import Path
    
    if len(sys.argv) > 1:
        emf_file = Path(sys.argv[1])
        png_file = emf_file.with_suffix('.png')
        convert_emf_to_png(emf_file, png_file)
    else:
        print("Usage: python3 test_aspose.py <emf_file>")
