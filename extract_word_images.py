# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "aspose-words",
#     "pillow",
#     "easyocr",
#     "opencv-python-headless",
# ]
# ///
"""
Wrapper script for image_extractor package.
"""
import sys
from pathlib import Path

# Add current directory to path so we can import the package
sys.path.append(str(Path(__file__).parent))

from image_extractor.main import main

if __name__ == "__main__":
    main()
