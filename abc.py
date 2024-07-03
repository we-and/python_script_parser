import os
from docx import Document

file_path = '/Users/jd/Downloads/Ruthless (Film) - CDSL (2).docx'

# Check if the file exists and is a file
if not os.path.isfile(file_path):
    print(f"File does not exist at path: {file_path}")
else:
    print(f"File found at path: {file_path}")

    # Attempt to open the document
    try:
        doc = Document(file_path)
        print("Document loaded successfully.")
    except Exception as e:
        print(f"Error loading document: {e}")

# Additional troubleshooting: Check file access permissions
try:
    with open(file_path, 'rb') as f:
        f.read()
    print("File is readable.")
except Exception as e:
    print(f"Cannot read file: {e}")
