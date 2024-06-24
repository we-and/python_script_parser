
import ctypes
import chardet
import pandas as pd
import os
import webbrowser
import platform
import shutil
def make_dpi_aware():
    try:
        # Attempt to set the process DPI awareness to the system DPI awareness
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except AttributeError:
        # Fallback if SetProcessDpiAwareness does not exist (possible in older Windows versions)
        ctypes.windll.user32.SetProcessDPIAware()
def get_excel_column_name(column_index):
    """Convert a 1-based column index to an Excel column name (e.g., 1 -> A, 27 -> AA)."""
    column_name = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        column_name = chr(65 + remainder) + column_name
    return column_name
def help_word_table():
    url="https://www.youtube.com/watch?v=XxZBVQKvihI"
    webbrowser.open(url)
def help_pdf_text():
    url="https://youtu.be/1F8yu14x6u8"
    webbrowser.open(url)
def help_merge():
    url="https://youtu.be/9MqWrRHDRPk"
    webbrowser.open(url)

def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result
    

def convert_csv_to_xlsx(csv_file_path, xlsx_file_path):
    # Read the CSV file
    df = pd.read_csv(csv_file_path)

    # Write the DataFrame to an Excel file
    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')




def get_encoding(enc):
    #print("Guess encoding from"+str(enc))
    if enc=="ascii":
        return "ISO-8859-1"
    elif enc=="ISO-8859-1":
        return "ISO-8859-1"
    elif enc=="Windows-1252":
        return "Windows-1252"       
    return "utf-8"
def copy_folder_contents(source_folder, destination_folder):
    """
    Copy all files from the source folder to the destination folder.
    
    Args:
    source_folder (str): Path to the source folder.
    destination_folder (str): Path to the destination folder.
    
    Returns:
    bool: True if successful, False if an error occurred.
    """
    try:
        print(f"copy files from {source_folder} to {destination_folder} ")
        # Create the destination folder if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        
        # Iterate through all items in the source folder
        for item in os.listdir(source_folder):
            source_item = os.path.join(source_folder, item)
            destination_item = os.path.join(destination_folder, item)
            
            # If it's a file, copy it
            if os.path.isfile(source_item):
                shutil.copy2(source_item, destination_item)
            # If it's a directory, recursively copy its contents
            elif os.path.isdir(source_item):
                shutil.copytree(source_item, destination_item)
        
        print(f"All contents copied from {source_folder} to {destination_folder}")
        return True
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False

def get_intial_treeview_folder_path():
    print(f"get_intial_treeview_folder_path")
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "Scripti","examples")
    print(f"create {ini_file_path}")
    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path

def get_setting_ini_path():
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "Scripti", "settings.ini")

    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path
def get_os():
    if os.name == 'nt':
        return 'Windows'
    elif os.name == 'posix':
        if 'darwin' in platform.system().lower():
            return 'macOS'
        elif 'linux' in platform.system().lower():
            return 'Linux'
    else:
        return 'Unknown'
    
def get_log_file_path():
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "Scripti", "app_log.txt")

    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path


def get_temp_folder_path():
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "Scripti", "tmp")

    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path
def get_recentfiles_file_path():
    
   # ini_file_path = os.path.join(script_folder, 'settings.ini')

    user_home = os.path.expanduser("~")
    ini_file_path = os.path.join(user_home, "Library", "Application Support", "Scripti", "recentfiles.txt")

    # Ensure the directory exists
    os.makedirs(os.path.dirname(ini_file_path), exist_ok=True)

    return ini_file_path


def save_string_to_file(text, filename):
        """Saves a given string `text` to a file named `filename`."""
        
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)