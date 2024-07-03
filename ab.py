import os
from docx import Document

file_path = '/Users/jd/Downloads/Ruthless (Film) - CDSL.docx'

# Check if the file exists
if not os.path.isfile(file_path):
    print(f"File does not exist at path: {file_path}")
else:
    print(f"File found at path: {file_path}")

    # Check if the file can be opened by python-docx
    try:
        doc = Document(file_path)
        print("Document loaded successfully.")
    except Exception as e:
        print(f"Error loading document: {e}")
