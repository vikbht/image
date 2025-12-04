import re

def sanitize_filename(text: str) -> str:
    """
    Sanitize text to be safe for use as a filename.
    """
    # Remove invalid characters
    text = re.sub(r'[<>:"/\\|?*]', '', text)
    # Replace whitespace with underscores
    text = re.sub(r'\s+', '_', text)
    # Remove non-alphanumeric characters (except underscores and hyphens)
    text = re.sub(r'[^\w\-_]', '', text)
    # Remove leading/trailing underscores
    text = text.strip('_')
    
    if not text:
        return "untitled"
        
    # Limit length
    return text[:50]
