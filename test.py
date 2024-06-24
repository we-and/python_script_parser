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

#import pytesseract
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTImage
if False:
    from script_parser import process_script,get_pdf_page_blocks,detect_word_table,run_convert_pdf_to_txt,split_elements, get_pdf_text_elements, is_supported_extension,convert_word_to_txt,convert_xlsx_to_txt,convert_rtf_to_txt,convert_pdf_to_txt,filter_speech
    #import pypdfium2 as pdfium

    #import pytesseract
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.converter import TextConverter, PDFPageAggregator
    from pdfminer.layout import LAParams
    from pdfminer.pdfpage import PDFPage
    from pdfminer.layout import LTImage

## .INI FILE
def get_setting_ini_path():
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "ScriptiTest", "settings.ini")

    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path
def check_settings_ini_exists():
    # Get the absolute path of the directory where the script is located
#    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = get_setting_ini_path()

    # Check if the settings.ini file exists
    if os.path.isfile(ini_file_path):
        print(f"settings.ini file exists at: {ini_file_path}")
        return True
    else:
        print(f"settings.ini file does not exist in the directory: {ini_file_path}")
        return False


def write_settings_ini():
    # Get the absolute path of the directory where the script is located
    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the content to write to the settings.ini file
    content = f"SCRIPT_FOLDER = {script_folder}"
    
    # Define the path to the settings.ini file in the same directory as the script
#    ini_file_path = os.path.join(script_folder, 'settings.ini')
    ini_file_path = get_setting_ini_path()
    # Write the content to the settings.ini file
    with open(ini_file_path, 'w') as ini_file:
        ini_file.write(content)
    
    print(f"settings.ini file created at: {ini_file_path}")


def read_settings_ini():
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = get_setting_ini_path()
    
    # Check if the settings.ini file exists
    if not os.path.isfile(ini_file_path):
        raise FileNotFoundError(f"settings.ini file does not exist in the directory: {ini_file_path}")
    
    # Read the settings.ini file and store settings in a dictionary
    settings = {}
    with open(ini_file_path, 'r') as ini_file:
        for line in ini_file:
            line = line.strip()
            if line and '=' in line:  # Ensure the line contains an '=' character
                key, value = line.split('=', 1)
                settings[key.strip()] = value.strip()
    
    return settings
def update_ini_settings_file(field,new_folder):
    # Get the absolute path of the directory where the script is located
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = get_setting_ini_path()

    # Check if the settings.ini file exists
    if not os.path.isfile(ini_file_path):
        raise FileNotFoundError(f"settings.ini file does not exist in the directory: {ini_file_path}")
    
    # Read the current settings and store them in a dictionary
    settings = {}
    with open(ini_file_path, 'r') as ini_file:
        for line in ini_file:
            line = line.strip()
            if line and '=' in line:
                key, value = line.split('=', 1)
                settings[key.strip()] = value.strip()
    
    # Update the SCRIPT_FOLDER field
    settings[field] = new_folder
    
    # Write the updated settings back to the settings.ini file
    with open(ini_file_path, 'w') as ini_file:
        for key, value in settings.items():
            ini_file.write(f"{key} = {value}\n")
    
    print(f"settings.ini file updated with SCRIPT_FOLDER = {new_folder}")



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
     #   img.save("generated_image.png")
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