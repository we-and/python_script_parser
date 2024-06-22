import tkinter as tk
from tkinter import ttk, filedialog, Text,Menu,Toplevel,Scrollbar, Scale, HORIZONTAL, VERTICAL
import os
import chardet
import io
import tkinter.font as tkFont
import subprocess
import traceback
import tkinter.font as tkfont
import math
import webbrowser
import time
import logging
from tkinter import ttk, messagebox
import sys
import csv
import re
import platform
import pandas as pd
from docx import Document
from PyPDF2 import PdfWriter, PdfReader

from pdfplumber.pdf import PDF
import pdfplumber

from PIL import Image, ImageTk,ImageDraw

if False:
    from script_parser import process_script,get_pdf_page_blocks,detect_word_table,run_convert_pdf_to_txt,split_elements, get_pdf_text_elements, is_supported_extension,convert_word_to_txt,convert_xlsx_to_txt,convert_rtf_to_txt,convert_pdf_to_txt,filter_speech
    #import pypdfium2 as pdfium

    #import pytesseract
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.converter import TextConverter, PDFPageAggregator
    from pdfminer.layout import LAParams
    from pdfminer.pdfpage import PDFPage
    from pdfminer.layout import LTImage




def main():
    # Create the main window
    root = tk.Tk()
    root.title("Minimal Tkinter App")

    # Create a label widget
    label = tk.Label(root, text="Hello, Tkinter!")
    label.pack(padx=20, pady=20)

    if True:
    # Create a new image with a white background
        width, height = 300, 200
        img = Image.new('RGB', (width, height), color='white')

        # Get a drawing context
        draw = ImageDraw.Draw(img)

        # Draw a simple shape (a red rectangle)
        draw.rectangle([50, 50, 250, 150], fill='red', outline='black')

        # Save the image
        img.save("generated_image.png")
        print("Image created and saved as 'generated_image.png'")

    data = {
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'City': ['New York', 'San Francisco', 'Los Angeles']
    }
    df = pd.DataFrame(data)

    # Perform a simple operation (calculate mean age)
    mean_age = df['Age'].mean()

    # Display the DataFrame and the result
    print("DataFrame:")
    print(df)
    print(f"\nMean Age: {mean_age}")

    # Create a button widget
    button = tk.Button(root, text="Click me!", command=root.quit)
    button.pack(pady=10)

    # Start the Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()