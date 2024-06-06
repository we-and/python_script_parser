import os
import tkinter as tk
from tkinter import ttk, filedialog, Text,Menu,Toplevel
from script_parser import process_script, is_supported_extension,convert_word_to_txt,convert_rtf_to_txt,convert_pdf_to_txt,filter_speech
import pandas as pd
import chardet
import tkinter.font as tkFont
import subprocess
import platform
import re
import sys
import csv
import math
import logging

logging.basicConfig(filename='app.log',level=logging.DEBUG)
logging.debug("Script starting...")


import ctypes

import threading
last_row_id = None

countingMethods=[
#    "LINE_COUNT",
 #   "WORD_COUNT",
    "ALL",
    "BLOCKS_50",
    #"BLOCKS_40",
   # "ALL_NOSPACE",
   # "ALL_NOPUNC",
  #  "ALL_NOSPACE_NOPUNC",
 #   "ALL_NOAPOS",    
]

countingMethodNames={
  #  "LINE_COUNT":"Lines",
   # "WORD_COUNT":"Words",
    "ALL":"Caracteres",
    "BLOCKS_50":"Repliques",
#    "BLOCKS_40":"Blocks (40)",
#    "ALL_NOSPACE":"No space",
 #   "ALL_NOPUNC":"No punctuation",
  #  "ALL_NOSPACE_NOPUNC":"No space, no punctuation",
   # "ALL_NOAPOS":"No apostrophe",
}

countingMethod="ALL"
currentOutputFolder=""
currentFilePath=""
currentScriptFilename=""
outputFolder="tmp"
currentRightclickRowId=None
currentXlsxPath=""
currentDialogPath=""
currentTimelinePath=""
currentBreakdown=None
currentFig=None
currentCanvas=None
currentCharacterMergeFromName=None
currentCharacterSelectRowId=None
currentCharacterMultiSelectRowIds=None
currentDisabledCharacters=[]
currentResultCharacterOrderMap=None
currentResultEnc=""
currentResultName=""
currentResultLinecountMap=None
currentResultSceneCharacterMap=None
currentMergePopupTable=None
currentMergedCharacters={}
currentMergedCharactersTo={}
currentMergePopupWindow=None
def myprint2(s):
    logging.debug(s)
    print(s)
def myprint3(s):
    logging.debug(s)
    print(s)
def myprint4(s):
    logging.debug(s)
    print(s)
def myprint5(s):
    logging.debug(s)
    print(s)

def make_dpi_aware():
    try:
        # Attempt to set the process DPI awareness to the system DPI awareness
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except AttributeError:
        # Fallback if SetProcessDpiAwareness does not exist (possible in older Windows versions)
        ctypes.windll.user32.SetProcessDPIAware()

# Only call this function if your application is running on Windows
if sys.platform.startswith('win32'):
    make_dpi_aware()

if not os.path.exists(outputFolder):
    os.mkdir(outputFolder)


def compute_length_by_method(line,method):
    res=0
    if method=="ALL":
        res=len(line)
    elif method=="LINE_COUNT":
        res=1
    elif method=="ALL_NOSPACE":
        res= len(line.replace(" ",""))
    elif method=="ALL_NOPUNC":
        res= len(line.replace(",","").replace("?","").replace(".","").replace("!",""))
    elif method=="ALL_NOSPACE_NOPUNC":
        res= len(line.replace(" ","").replace(",","").replace("?","").replace(".","").replace("!",""))
    elif method=="ALL_NOAPOS":
        res= len(line.replace("'",""))
    elif method=="WORD_COUNT":
        res= len(line.split(" "))
    elif method=="BLOCKS_50":
        res= len(line)/50
    elif method=="BLOCKS_40":
        res= len(line)/40
    else:
        res= -1
    #print("compute_length_by_method METHOD "+method+" "+str(res))
    return res


def convert_csv_to_xlsx(csv_file_path, xlsx_file_path):
    # Read the CSV file
    df = pd.read_csv(csv_file_path)

    # Write the DataFrame to an Excel file
    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')



def load_tree(parent, root_path):
    # Clear the tree view if root_path is the starting directory
    if parent == "":
        folders.delete(*folders.get_children())
        parent = folders.insert('', 'end', text=os.path.basename(root_path), open=True, values=["","",])

    # List all entries in the directory
    try:
        dir_entries = os.listdir(root_path)
    except PermissionError:
        return  # Skip directories without permission

    # Separate and sort directories and files
    dirs = sorted([entry for entry in dir_entries if os.path.isdir(os.path.join(root_path, entry))])
    files = sorted([entry for entry in dir_entries if not os.path.isdir(os.path.join(root_path, entry))])

    # Insert directories first
    for entry in dirs:
        entry_path = os.path.join(root_path, entry)

        dir_id = folders.insert(parent, 'end', text=" "+entry, image=folder_icon, open=False, values=[entry_path,"Dossier",],tags=('folder'))
        #folders.insert(dir_id, 'end', text=os.path.basename(entry_path), open=True, values=[entry_path])
        folders.insert(dir_id, 'end', text="Loading...", values=["dummy"])  # Dummy node

        #load_tree(dir_id, entry_path)  # Recursively load subdirectories

    # Insert files
    for entry in files:
        entry_path = os.path.join(root_path, entry)
        supported_tag="not_supported"
        
        name, extension = os.path.splitext(entry)
        extension_without_dot=extension
        if extension.startswith("."):
            extension_without_dot=extension_without_dot[1:]
            
        if is_supported_extension(extension):
            supported_tag="supported" 
            if extension==".docx" or extension==".doc":
                folders.insert(parent, 'end', text=" "+entry,image=docx_icon, values=[entry_path, extension_without_dot,],tags=(supported_tag))
            else:
                folders.insert(parent, 'end', text=" "+entry,image=txt_icon, values=[entry_path, extension_without_dot,],tags=(supported_tag))
        else:        
            folders.insert(parent, 'end', text=" "+entry, values=[entry_path, extension_without_dot,],tags=(supported_tag))

def on_motion(event):
    # Identify the row on which the mouse is currently hovering
    row_id = folders.identify_row(event.y)
    if row_id:
        # Retrieve current tags and add 'hover' tag
        current_tags = set(folders.item(row_id, 'tags'))
        
        current_tags.add('hover')
        folders.item(row_id, tags=list(current_tags))

    # Reset the background color of previously hovered rows
    global last_row_id
    if last_row_id and last_row_id != row_id:
        current_tags = set(folders.item(last_row_id, 'tags'))
        current_tags.discard('hover')  # Remove the hover tag
        folders.item(last_row_id, tags=list(current_tags))
    last_row_id = row_id

def on_leave(event):
    # When the mouse leaves the Treeview, reset the background of the last hovered row
    global last_row_id
    if last_row_id:
        current_tags = set(folders.item(last_row_id, 'tags'))
        current_tags.discard('hover')  # Remove the hover tag
        folders.item(last_row_id, tags=list(current_tags))
    last_row_id = None



def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result
def get_encoding(enc):
    #print("Guess encoding from"+str(enc))
    if enc=="ascii":
        return "ISO-8859-1"
    elif enc=="ISO-8859-1":
        return "ISO-8859-1"
    elif enc=="Windows-1252":
        return "Windows-1252"       
    return "utf-8"

def reset_tables(): 
    #print("reset_tables")
    
    
    for item in breakdown_table.get_children():
        breakdown_table.delete(item)
    for item in character_list_table.get_children():
        character_list_table.delete(item)
    for item in character_table.get_children():
        character_table.delete(item)
    for item in stats_table.get_children():
        stats_table.delete(item)
    for item in character_stats_table.get_children():
        character_stats_table.delete(item)
    
def runJob(file_path,method):
    global currentFilePath
    global currentScriptFilename
    global currentBreakdown
    global currentOutputFolder
    global currentTimelinePath

    global currentResultCharacterOrderMap
    global currentResultEnc
    global currentResultName
    global currentResultLinecountMap
    global currentResultSceneCharacterMap
    
    currentFilePath=file_path
    reset_tables()
    # Check if the selected item is a file and display its content
    if os.path.isfile(file_path):
        #show_loading()

        try:
            file_name = os.path.basename(file_path)
            currentScriptFilename=file_name
            name, extension = os.path.splitext(file_name)
            
            myprint3("Extension           :"+extension)
            if is_supported_extension(extension):
                myprint3("Supported           : YES")

                encoding_info = detect_file_encoding(file_path)
                encoding=encoding_info['encoding']
                myprint3("Encoding detected   : "+str(encoding))
                myprint3("Encoding confidence : "+str(encoding_info['confidence']))

                enc=get_encoding(encoding)
                myprint3("Encoding used       : "+str(enc))
    #            encodings = ['windows-1252', 'iso-8859-1', 'utf-16','utf-8']
     #           for encod in encodings:
     #               print("try encoding"+encod)


                currentOutputFolder=outputFolder+"/"+name+"/"
                if not os.path.exists(currentOutputFolder):
                    os.mkdir(currentOutputFolder)
                extension=extension.lower()
                if extension==".docx" or extension==".doc":
                    myprint3("Conversion Word to txt")

                    forceMode=""
                    forceCols={}
                    print(" !!!!!!!!!! CHECK FILENAME")
                    if "CLEAR CUT" in file_name:
                        print(" !!!!!!!!!! FORCE")
                        forceMode="DETECT_CHARACTER_DIALOG"
                        forceCols={
                            "CHARACTER":5,
                            "DIALOG":6
                        }
                        
                    converted_file_path=convert_word_to_txt(file_path,os.path.abspath(currentOutputFolder),forceMode=forceMode,forceCols=forceCols)
                    if len(converted_file_path)==0:
                        myprint3("Conversion docx to txt failed")
                        myprint3("Failed")
                        hide_loading()
                        return 
                    file_path=converted_file_path
                    myprint3("Conversion file path :"+file_path)
                    
                if extension==".rtf":
                    myprint3("Conversion RTF to txt")
                    converted_file_path,txt_encoding=convert_rtf_to_txt(file_path,os.path.abspath(currentOutputFolder),enc)
                    if len(converted_file_path)==0:
                        myprint3("Conversion rtf to txt failed")
                        myprint3("Failed")
                        hide_loading()
                        return
                    enc=txt_encoding
                    file_path=converted_file_path
                    myprint3("Conversion file path :"+file_path)
                if extension==".pdf":
                    myprint3("Conversion PDF to txt")
                    converted_file_path,txt_encoding=convert_pdf_to_txt(file_path,os.path.abspath(currentOutputFolder),enc)
                    if len(converted_file_path)==0:
                        myprint3("Conversion pdf to txt failed")
                        myprint3("Failed")
                        hide_loading()
                        return
                    enc=txt_encoding
                    file_path=converted_file_path
                    myprint3("Conversion file path :"+file_path)

                myprint3("Opening")
                with open(file_path, 'r', encoding=enc) as file:
                    myprint3("Opened")
                    content = file.read()
                    myprint3("Read")
                    file_preview.delete(1.0, tk.END)
                    file_preview.insert(tk.END, content)
                    

                    myprint3("Process")
                    breakdown,character_scene_map,scene_characters_map,character_linecount_map,character_order_map,character_textlength_map=process_script(file_path,currentOutputFolder,name,method,enc)

                    myprint3("Processed")

                    if breakdown==None:
                        myprint3("Failed")
                        hide_loading()
                    else:
                        myprint3("OK")
                        currentBreakdown=breakdown

                        png_output_file=currentOutputFolder+name+"_timeline.png"
                        currentTimelinePath=png_output_file
                        currentResultCharacterOrderMap=character_order_map
                        currentResultEnc=enc
                        currentResultName=name
                        currentResultLinecountMap=character_linecount_map
                        currentResultSceneCharacterMap=scene_characters_map
                        postProcess(breakdown,character_order_map,enc,name,character_linecount_map,scene_characters_map,png_output_file)
                        
            else:
                myprint3(" > Not supported")
                #stats_label.config(text=f"Format {extension} not supported")
            hide_loading()

        except Exception as e:
            file_preview.delete(1.0, tk.END)
            file_preview.insert(tk.END, f"Error opening file: {e} tried with encoding{ enc}")
            hide_loading()

def postProcess(breakdown,character_order_map,enc,name,character_linecount_map,scene_characters_map,png_output_file):

    global currentResultCharacterOrderMap
    myprint3("Post process")
    fill_breakdown_table(breakdown)
    fill_character_list_table(character_order_map,breakdown)
    fill_character_stats_table(character_order_map,breakdown,enc)
    fill_stats_table(breakdown)
    fill_character_table(character_order_map, breakdown,character_linecount_map,scene_characters_map)
    save_dialog_csv(breakdown,enc)
    
def save_dialog_csv(breakdown,enc):
    global currentDialogPath
    totalcsvpath=currentOutputFolder+"/"+currentScriptFilename+"-dialog.csv"
    currentDialogPath=totalcsvpath
    
    data=[]
    for item in breakdown:
        line_idx=item['line_idx']
        type_=item['type']
        if(type_=="SPEECH"):
            speech=item['speech']
            character=item['character']
            character_raw=item['character_raw']
            
            filtered_speech=filter_speech(speech)
            datarow=[str(line_idx),character,filtered_speech];
            for m in countingMethods: 
                        #myprint4("add"+str(m))
                        le=compute_length_by_method(filtered_speech,m)
                        datarow.append(le)
                
            data.append(datarow)
    else:
        myprint3("save_dialog_csv skip"+character)
    
            
    myprint3("Saving dialog csv encoding="+enc)
    #myprint3("data"+str(data))
    with open(totalcsvpath, mode='w', newline='',encoding=enc) as file:
        writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        
        # Write data to the CSV file
        for row in data:
            #myprint3("Write "+str(row))
            writer.writerow(row)
    xlsxpath=totalcsvpath.replace(".csv",".xlsx")
    myprint3("xlsx"+xlsxpath)
    convert_dialog_csv_to_xlsx2(totalcsvpath,xlsxpath,enc)
def on_folder_select(event):
    global currentScriptFilename
    global currentOutputFolder
    logging.debug("on_folder_select")
    #myprint3("FOLDER SELECT")
    selected_item = folders.selection()[0]
    file_path = folders.item(selected_item, 'values')[0]
    if os.path.isfile(file_path):
        logging.debug(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
        currentScriptFilename=file_path
        loading_label_txt.config(text="Analyse de "+file_path)
        show_loading()
        app.update_idletasks()  # Force the UI to update
        #for item in disabled_character_list_table.get_children():
         #   disabled_character_list_table.delete(item)
    
        threading.Thread(target=runJob,args=(file_path,countingMethod)).start()
#        runJob(file_path,countingMethod)


# Function to remove all items
def remove_all_tree_items():
    for item in folders.get_children():
        folders.delete(item)


def open_folder():
    myprint3("Change folder")
    remove_all_tree_items()
    directory = filedialog.askdirectory(initialdir=os.getcwd())
    if directory:
        update_ini_settings_file("SCRIPT_FOLDER",directory)
        load_tree("",directory)

def open_script():
    myprint3("Open script")
    file_path = filedialog.askopenfilename()
    if file_path:
        print(f"Selected file: {file_path}")
        runJob(file_path,"ALL")

def open_folder_firsttime():
    myprint3("Change folder")
    directory = filedialog.askdirectory(initialdir=os.getcwd(),title="Choose your scripts folder")
    if directory:
        update_ini_settings_file("SCRIPT_FOLDER",directory)


def exit_app():
    app.quit()


def center_window():
    app.update_idletasks()
    width = app.winfo_width()
    height = app.winfo_height()
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
     # Calculate width and height as 80% of screen dimensions
    width = int(screen_width * 0.9)
    height = int(screen_height * 0.8)
    x = int((screen_width - width) / 2)
    y = 20#int((screen_height - height) / 2)
    app.geometry(f'{width}x{height}+{x}+{y}')



def reduce_width():
    current_width = app.winfo_width()
    current_height = app.winfo_height()
    # Reduce width by 1 pixel
    new_width = current_width - 1
    # Set the new geometry
    app.geometry(f"{new_width}x{current_height}")


def stats_per_character(breakdown,character_name):
    line_count=0
    word_count=0
    character_count=0
    
    for item in breakdown:
        if item['type']=="SPEECH":
            if item['character']==character_name:
                t=item['speech']
                filtered_speech=filter_speech(t)
                line_count=line_count+1
                le=compute_length_by_method(filtered_speech,"ALL")
                character_count=character_count+le
                word_count=word_count+len(t.split(" "))
    #myprint3(character_name,character_count)
            
    replica_count=math.ceil(character_count/50)
    return line_count,word_count,character_count,replica_count

def fill_character_table(character_order_map, breakdown,character_linecount_map,scene_characters_map):
    order_idx=0
    for item in character_order_map:
        lines=character_linecount_map[item]
        line_count,word_count,character_count,replica_count=stats_per_character(breakdown,item)
        scenes=scene_characters_map[item]
        scenes=', '.join(scenes)
        status="VISIBLE"
        if item in currentDisabledCharacters:
            status="HIDDEN"
            character_table.insert('','end',values=(" - ",item,status,str(character_count),str(math.ceil(character_count/50)),scenes),tags=("hidden"))
        elif item in currentMergedCharacters:
            status="MERGED (into "+str(currentMergedCharacters[item])+")"
            character_table.insert('','end',values=(" - ",item,status,str(character_count),str(math.ceil(character_count/50)),scenes),tags=("hidden"))
        else:
            order_idx=order_idx+1
            character_table.insert('','end',values=(str(order_idx),item,status,str(character_count),str(math.ceil(character_count/50)),scenes))
        
#        character_table.insert('','end',values=(str(character_order_map[item]),item,str(line_count),str(character_count),str(word_count),str(math.ceil(character_count)/50),str(math.ceil(character_count)/40),scenes))
        

def compute_length(method,line):
    if method=="ALL":
        return len(line);
    return len(line);

def fill_character_list_table(character_order_map, breakdown):
    #myprint3("fill_character_list_table")

    for character_name in character_order_map:
        #myprint3("CHAR add"+character_name)
        if (not character_name in currentDisabledCharacters) or (not character_name in currentMergedCharacters):
            character_named = character_name 
            #myprint4("CHAR add"+character_named)
            character_list_table.insert('','end',values=(character_named,))
        

def fill_character_stats_table(character_order_map, breakdown,encoding_used):
    myprint4("fill_character_stats_table")

    total_by_character_by_method={}
    for character_name in character_order_map:
        #myprint4("CHAR add"+character_name)

        character_named = character_name 
        #myprint4("CHAR add"+character_named)
    #    character_list_table.insert('','end',values=(character_named,))
        
        
        character_order=character_order_map[character_name]
        #myprint4("CHAR"+str(character_name))
        rowtotal=("-",character_name,"-","TOTAL")
        total_by_method={}
        for m in countingMethods:
            total_by_method[m]=0
        
        for item in breakdown:
            line_idx=item['line_idx']
            type_=item['type']
            if(type_=="SPEECH"):
                speech=item['speech']
                character=item['character']
                character_raw=item['character_raw']
                
                filtered_speech=filter_speech(speech)

                if character==character_name:
                    #myprint4("    MATCH"+str(speech))

                    row=(str(line_idx),character,character_raw, speech)
                    for m in countingMethods: 
                        #myprint4("add"+str(m))
                        le=compute_length_by_method(filtered_speech,m)
                        row=row+(str(le),)
                        total_by_method[m]=total_by_method[m]+le
                    #myprint4("add"+str(row))
                    character_stats_table.insert('','end',values=row)
        
        #round for BLOCKS
        for m in countingMethods:
            if m.startswith("BLOCKS"):
                total_by_method[m]=math.ceil(total_by_method[m])

        for m in countingMethods:
            rowtotal=rowtotal+(total_by_method[m],)
        character_stats_table.insert('','end',values=rowtotal,tags=['total'])
        
        total_by_character_by_method[character_name]=total_by_method
#        total_by_character_by_method[str(character_order)+" - "+character_name]=total_by_method
    totalcsvpath=currentOutputFolder+"/"+currentScriptFilename+"-comptage.csv"
    generate_total_csv(total_by_character_by_method,totalcsvpath,encoding_used,character_order_map)
    for item in character_stats_table.get_children():
        character_stats_table.delete(item)



def save_string_to_file(text, filename):
        """Saves a given string `text` to a file named `filename`."""
        myprint4(" > Write to "+filename)
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)
def get_excel_column_name(column_index):
    """Convert a 1-based column index to an Excel column name (e.g., 1 -> A, 27 -> AA)."""
    column_name = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        column_name = chr(65 + remainder) + column_name
    return column_name
def convert_csv_to_xlsx2(csv_file_path, xlsx_file_path, n,encoding_used):
    myprint4("convert_csv_to_xlsx2 "+csv_file_path+ " to "+xlsx_file_path+" "+encoding_used)

    # Read the CSV file
    df = pd.read_csv(csv_file_path,header=None,encoding=encoding_used)
    myprint4("convert_csv_to_xlsx2 4")
    # Convert columns explicitly to numeric where appropriate
    for col in df.columns[1:]:  # Skip the first column if it's non-numeric (e.g., names)
        df[col] = pd.to_numeric(df[col], errors='ignore')

    # Write the DataFrame to an Excel file
    myprint4(" > Write to "+xlsx_file_path)

    header1=["SCRIPT_NAME"]
    header2=["Role"]
    for i in countingMethods:
        if i != "ALL":
            header1.append("?")
            txt=countingMethodNames[i]
            if i=="BLOCKS_50":
                txt="Lines"
            header2.append(txt)
        
    header_rows = pd.DataFrame([
        header1,
        header2
    ])
    myprint4("convert_csv_to_xlsx2 3")
    # Concatenate the header rows and the original data
    # The ignore_index=True option reindexes the new DataFrame
    df = pd.concat([header_rows, df], ignore_index=True)

    # Write the DataFrame to an Excel file
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Load the workbook and sheet for modification
        workbook = writer.book
        sheet = workbook['Sheet1']

        # Merge cells in the first and second new rows
        # Assuming you want to merge from the first to the last column
        col=get_excel_column_name(n)
        sheet.merge_cells("A1:"+col+"1")  # Modify range according to your number of columns
        sheet.merge_cells("A2:"+col+"2")  # Modify range according to your number of columns
 #       sheet.merge_cells('A2:D2')  # Modify this as needed
        sheet['A1'] = currentScriptFilename
        sheet['A2'] = "Length: "
        sheet.column_dimensions['A'].width = 50 
 # Load the workbook and get the active sheet
    myprint4("convert_csv_to_xlsx2 4")
  
def convert_dialog_csv_to_xlsx2(csv_file_path, xlsx_file_path, encoding_used):
    myprint4("convert_dialog_csv_to_xlsx2 "+csv_file_path+ " to "+xlsx_file_path+" "+encoding_used)

    # Read the CSV file
    df = pd.read_csv(csv_file_path,header=None,encoding=encoding_used)
    # Convert columns explicitly to numeric where appropriate
    for col in df.columns[1:]:  # Skip the first column if it's non-numeric (e.g., names)
        df[col] = pd.to_numeric(df[col], errors='ignore')

    # Write the DataFrame to an Excel file
    myprint4(" > Write to "+xlsx_file_path)

    header1=["Line","Character","Dialog"]
    header2=[]
    for i in countingMethods:
            txt=countingMethodNames[i]
            if i=="BLOCKS_50":
                txt="Lines"
            header1.append(txt)
        
    header_rows = pd.DataFrame([
        header1,
    ])
    myprint4("convert_csv_to_xlsx2 3")
    # Concatenate the header rows and the original data
    # The ignore_index=True option reindexes the new DataFrame
    df = pd.concat([header_rows, df], ignore_index=True)

    # Write the DataFrame to an Excel file
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

  



def generate_total_csv(total,csv_path,encoding_used,character_order_map):
    global currentXlsxPath
    #myprint4("Total csv path          : "+csv_path)
    #header

    s=""
    showHeader=False
    if showHeader:
        s="Role,"
        for m in countingMethods:
            if m!="ALL":                   
                s=s+str(m)+","
        s=s[0:len(s)-1]
        s=s+"\n"
    #myprint4("Total csv path          : 1")
    
    data = []
    order_idx=0


    

    for character in total:
        if (not character in currentDisabledCharacters) and (not character in currentMergedCharacters):
            order_idx=order_idx+1
            datarow=[str(order_idx)+" - " +str(character)];
            for method in total[character]:
                if method!="ALL":
                    #check merged characters and add eventually
                    #myprint4("char="+str(character))
                    if character in currentMergedCharactersTo:
                        #myprint4("in merged"+str(currentMergedCharactersTo))
                        mergedwith=currentMergedCharactersTo[character]
                        for k in mergedwith:
                            #myprint4("in merged add "+str(k))
                            total[character][method]=total[character][method]+total[k][method]
                    #else:
                        #myprint4("not merged")
                    #myprint4(str(character)+": Add method "+method+" = "+str(total[character][method]))
                    datarow.append(str(total[character][method]))

            data.append(datarow)
        else:
            myprint4("generate_total_csv skip"+character)
    
            
    #myprint4("data"+str(data))
    with open(csv_path, mode='w', newline='',encoding=encoding_used) as file:
        writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        
        # Write data to the CSV file
        for row in data:
            #myprint4("Write "+str(row))
            writer.writerow(row)

    
    xlsx_path=csv_path.replace(".csv",".xlsx")
    currentXlsxPath=xlsx_path
    n=len(countingMethods)+1
    #myprint4("Total xlsx path          : "+xlsx_path)
    convert_csv_to_xlsx2(csv_path,xlsx_path,n,encoding_used)

def on_button_click():
    myprint4("Button clicked!")
def export_csv():
    myprint4("Export")

def fill_breakdown_table(breakdown):
    for item in breakdown:
        type_=item['type']
        line_idx=item['line_idx']
        if(type_=="SCENE_SEP"):
            scene_id=item['scene_id']
            breakdown_table.insert('','end',values=("","","",""), tags=('border'))
            breakdown_table.insert('','end',values=(str(line_idx),"New scene",scene_id,""), tags=('scene','bold'))
        elif(type_=="SPEECH"):
            speech=item['speech']
            character=item['character']
            if not character in currentDisabledCharacters:
                breakdown_table.insert('','end',values=(str(line_idx),"Speech",character,speech))
        elif(type_=="NONSPEECH"):         
            text=item['text']
            breakdown_table.insert('','end',values=(str(line_idx),"Other",text,""), tags=('nonspeech',))
    #myprint4("NB ROWS = "+str(len(breakdown_table.get_children())))
    breakdown_table.update_idletasks()

def fill_stats_table(breakdown):
    for item in breakdown:
        type_=item['type']
        line_idx=item['line_idx']
        if(type_=="SPEECH"):
            speech=item['speech']
            filtered_speech=filter_speech(speech)
            character=item['character']
            tout=len(filtered_speech)
            if not character in currentDisabledCharacters:
                stats_table.insert('','end',values=(str(line_idx),character,speech,str(tout)))
    myprint4("NB ROWS = "+str(len(breakdown_table.get_children())))
    breakdown_table.update_idletasks()


def clear_table(treeview):
    """
    Clears all rows from the given Treeview table.
    
    Args:
        treeview (ttk.Treeview): The Treeview widget instance.
    """
    for item in treeview.get_children():
        treeview.delete(item)


###################################################################################################
## .INI FILE
def check_settings_ini_exists():
    # Get the absolute path of the directory where the script is located
    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = os.path.join(script_folder, 'settings.ini')
    
    # Check if the settings.ini file exists
    if os.path.isfile(ini_file_path):
        myprint4(f"settings.ini file exists at: {ini_file_path}")
        return True
    else:
        myprint4(f"settings.ini file does not exist in the directory: {script_folder}")
        return False


def write_settings_ini():
    # Get the absolute path of the directory where the script is located
    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the content to write to the settings.ini file
    content = f"SCRIPT_FOLDER = {script_folder}"
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = os.path.join(script_folder, 'settings.ini')
    
    # Write the content to the settings.ini file
    with open(ini_file_path, 'w') as ini_file:
        ini_file.write(content)
    
    myprint4(f"settings.ini file created at: {ini_file_path}")


def read_settings_ini():
    # Get the absolute path of the directory where the script is located
    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = os.path.join(script_folder, 'settings.ini')
    
    # Check if the settings.ini file exists
    if not os.path.isfile(ini_file_path):
        raise FileNotFoundError(f"settings.ini file does not exist in the directory: {script_folder}")
    
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
    script_folder = os.path.abspath(os.path.dirname(__file__))
    
    # Define the path to the settings.ini file in the same directory as the script
    ini_file_path = os.path.join(script_folder, 'settings.ini')
    
    # Check if the settings.ini file exists
    if not os.path.isfile(ini_file_path):
        raise FileNotFoundError(f"settings.ini file does not exist in the directory: {script_folder}")
    
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
    
    myprint4(f"settings.ini file updated with SCRIPT_FOLDER = {new_folder}")

###################################################################################################
## MAIN
logging.debug("Checking settings ini file")

settings_ini_exists = check_settings_ini_exists()
if settings_ini_exists == False:
    logging.debug("Writing settings ini file")
    write_settings_ini()
    script_folder = os.path.abspath(os.path.dirname(__file__))
    ini_file_path = os.path.join(script_folder, 'settings.ini')
    
    update_ini_settings_file("SCRIPT_FOLDER",script_folder+"/examples")

settings = read_settings_ini()
app_dir = os.path.dirname(os.path.abspath(__file__))
icons_dir =app_dir+"/icons/"
myprint4("App dir           :"+app_dir)
app = tk.Tk()
app.title('Scripti')
app.iconbitmap(icons_dir+'app_icon.ico') 
logging.debug("Creating app")

def on_resize(event):
        return

#app.bind('<Configure>', on_resize)

# Menu bar
menu_bar = Menu(app)
app.config(menu=menu_bar)

folder_icon = tk.PhotoImage(file=icons_dir+"folder_icon.png")  # Adjust path to your icon file
txt_icon = tk.PhotoImage(file=icons_dir+"txt_icon.png")  # Adjust path to your icon file
docx_icon = tk.PhotoImage(file=icons_dir+"docx_icon.png")  # Adjust path to your icon file
original_icon = tk.PhotoImage(file=icons_dir+"textd_icon.png")  # Adjust path to your icon file
char_icon = tk.PhotoImage(file=icons_dir+"character_icon.png")  # Adjust path to your icon file
timeline_icon = tk.PhotoImage(file=icons_dir+"timeline_icon.png")  # Adjust path to your icon file
scene_icon = tk.PhotoImage(file=icons_dir+"scenes_icon.png")  # Adjust path to your icon file
export_icon = tk.PhotoImage(file=icons_dir+"export_icon.png")  # Adjust path to your icon file
chat_icon = tk.PhotoImage(file=icons_dir+"chat_icon.png")  # Adjust path to your icon file


# File menu
file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Fichier", menu=file_menu)
file_menu.add_command(label="Ouvrir un dossier de travail...", command=open_folder)
file_menu.add_command(label="Ouvrir un fichier de script...", command=open_script)
#file_menu.add_command(label="Export csv...", command=export_csv)
file_menu.add_separator()
file_menu.add_command(label="Quitter", command=exit_app)

def on_folder_open(event):
    myprint4("on_folder_open")
    # Find the node that was opened
    oid = folders.focus()  # Get the ID of the focused item
    values = folders.item(oid, "values")
    myprint4("on_folder_open 1")

    if len(values) > 0 and values[0] == "dummy":
        myprint4("values>0 ignore")
        # Ignore the dummy nodes (if the first value in the tuple is "dummy")
        return
    myprint5("on_folder_open 2")

    # Check if the node has the dummy child indicating it hasn't been loaded
    children = folders.get_children(oid)
    if len(children) == 1 and folders.item(children[0], "values")[0] == "dummy":
        myprint5("on_folder_open 3")

        # Remove the dummy and load actual content
        folders.delete(children[0])
        myprint5("Load "+str(folders.item(oid, "values")))
        load_tree(oid, folders.item(oid, "values")[0])
    else:
        myprint5("on_folder_open 4 oid="+str(oid)+" val="+str( folders.item(oid, "values")))
        load_tree(oid, folders.item(oid, "values")[0])

def toggle_folder(event):
    myprint5("Toggle folder")
# Identify the item on which the click occurred
    x, y, widget = event.x, event.y, event.widget
    row_id = widget.identify_row(y)
    if not row_id:
        return  # Exit if the click didn't happen on a row
    
    # Toggle the open state of the node
    if widget.tag_has('folder', row_id):  # Check if the item has the 'folder' tag
        
        if widget.item(row_id, 'open'):  # If the folder is open, close it
            myprint5("opened, close")
            widget.item(row_id, open=False)

        else:  # If the folder is closed, open it
            myprint5("not opened, open")
            widget.item(row_id, open=True)
#            on_folder_open(event)
            children = folders.get_children(row_id)
            if len(children) == 1 and folders.item(children[0], "values")[0] == "dummy":
                myprint5("on_folder_open 3")

                # Remove the dummy and load actual content
                folders.delete(children[0])
            load_tree(row_id, folders.item(row_id, "values")[0])


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
    
def open_xlsx_recap():
    global currentXlsxPath
    os_=get_os()
    myprint5("Open"+currentXlsxPath)
    if os_=="Windows":
        currentOutputFolderAbs = os.path.abspath(currentXlsxPath)
        myprint5("Absolute path          : "+currentOutputFolderAbs)

    # Check if the folder exists
        if not os.path.exists(currentOutputFolderAbs):
            myprint5(f"Folder does not exist: {currentOutputFolderAbs}")
            return
        try:
            os.startfile(currentOutputFolderAbs)
        except Exception as e:
            myprint5(f"Failed to open file: {e}")
    else:
        try:
            subprocess.run(['open', currentXlsxPath], check=True)
        except subprocess.CalledProcessError as e:
            myprint5(f"Failed to open file: {e}")
def open_dialog_recap():
    global currentDialogPath
    os_=get_os()
    myprint5("Open"+currentDialogPath)
    if os_=="Windows":
        currentOutputFolderAbs = os.path.abspath(currentDialogPath)
        myprint5("Absolute path          : "+currentOutputFolderAbs)

    # Check if the folder exists
        if not os.path.exists(currentOutputFolderAbs):
            myprint5(f"Folder does not exist: {currentOutputFolderAbs}")
            return
        try:
            os.startfile(currentOutputFolderAbs)
        except Exception as e:
            myprint5(f"Failed to open file: {e}")
    else:
        try:
            subprocess.run(['open', currentDialogPath], check=True)
        except subprocess.CalledProcessError as e:
            myprint5(f"Failed to open file: {e}")

def open_file_in_system():
    global currentRightclickRowId
    myprint5("open_file_in_system" )
    selected_item = folders.focus() 
    file_path = folders.item(currentRightclickRowId, 'values')[0]
    myprint5(file_path)
    if os.path.isfile(file_path):
        
        myprint5(f"Opening file for item: {file_path}")
        abs_file_path=file_path#folders.item(selected_item)
        os_=get_os()
        myprint5("Open"+file_path)
        if os_=="Windows":
         
        # Check if the folder exists
            if not os.path.exists(abs_file_path):
                myprint5(f"Folder does not exist: {abs_file_path}")
                return
            try:
                os.startfile(abs_file_path)
            except Exception as e:
                myprint5(f"Failed to open file: {e}")
        else:
            try:
                subprocess.run(['open', abs_file_path], check=True)
            except subprocess.CalledProcessError as e:
                myprint5(f"Failed to open file: {e}")

def open_timeline():
    global currentTimelinePath
    os_=get_os()
    myprint5("Open"+currentTimelinePath)
    if os_=="Windows":
        currentOutputFolderAbs = os.path.abspath(currentTimelinePath)
        myprint5("Absolute path          : "+currentOutputFolderAbs)

    # Check if the folder exists
        if not os.path.exists(currentOutputFolderAbs):
            myprint5(f"Folder does not exist: {currentOutputFolderAbs}")
            return
        try:
            os.startfile(currentOutputFolderAbs)
        except Exception as e:
            myprint5(f"Failed to open file: {e}")
    else:
        try:
            subprocess.run(['open', currentTimelinePath], check=True)
        except subprocess.CalledProcessError as e:
            myprint5(f"Failed to open file: {e}")

def set_counting_method(i):
    myprint5("set method "+i)
    global countingMethod
    countingMethod=i
def set_line_size(i):
    print("a")
def show_popup_line_size():
    popup = tk.Toplevel()
    popup.title("Popup") 

    # Create a label
    label = ttk.Label(popup, text="Enter your text:")
    label.pack(pady=10)

    # Create a single-line text entry widget
    entry = ttk.Entry(popup, width=30)
    entry.pack(pady=10)
    button = ttk.Button(popup, text="Appliquer", command=set_line_size(i))
    button.pack(side=tk.TOP, fill=tk.X)

def show_popup_counting_method():
    popup = tk.Toplevel()
    popup.title("Popup") 

    for i in countingMethods:
        button = ttk.Button(popup, text=i, command=set_counting_method(i))
        button.pack(side=tk.TOP, fill=tk.X)

    dropdown = ttk.Combobox(popup, values=countingMethods)
    dropdown.pack(pady=20)
    dropdown.current(0)
 #   dropdown.bind('<<ComboboxSelected>>', on_value_change)

def resizechart(self, event=None):
        # Resize the figure to match the dimensions of its container
        width, height = event.width, event.height
        if width > 0 and height > 0:  # Check to prevent initial null-dimension error
            dpi = self.fig.get_dpi()
            self.fig.set_size_inches(width / dpi, height / dpi)
            self.canvas.draw()

def myprint5_frame_size(fr):
    # This function myprint5s the size of the frame
    # Ensure the frame has been rendered by Tkinter before calling this
    myprint5("Frame width:"+str( fr.winfo_width()))
    myprint5("Frame height:"+str( fr.winfo_height()))

def show_loading():
        #myprint5("SHOW LOADING")
        loading_label.pack()
        paned_window.pack_forget()
def hide_loading():
        #myprint5("HIDE LOADING")
        paned_window.pack(fill='both', expand=True)
        loading_label.pack_forget()

def restore_characters():
    myprint2("restore_characters")
    global currentDisabledCharacters
    currentDisabledCharacters=[]
    reset_tables()
    postProcess(currentBreakdown,currentResultCharacterOrderMap,currentResultEnc,currentResultName,currentResultLinecountMap,currentResultSceneCharacterMap,currentTimelinePath)


def clear_chart():
    global currentFig
    myprint2("CLEAR CHART")
    currentFig.clf()
    currentFig.canvas.draw()
settings_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Parametres", menu=settings_menu)
#settings_menu.add_command(label="Changer la methode de comptage counting method...", command=show_popup_counting_method)
settings_menu.add_command(label="Change la taille des repliques...", command=show_popup_line_size)
#settings_menu.add_command(label="Set block length...", command=open_folder)
settings_menu.add_command(label="Reafficher les personnages masques...", command=restore_characters)


loading_label = ttk.Frame(app)
#loading_label.pack(side=tk.TOP, fill=tk.X,expand=True)
# Load folder button
#load_button = ttk.Button(loading_label, text="Hide ", command=hide_loading)
#load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)

loading_label_txt = ttk.Label(loading_label, text="Analyzing", font=('Arial', 12),padding="20 100 20 20")
loading_label_txt.pack(side=tk.TOP, fill=tk.X,expand=True)


# Create a PanedWindow widget
paned_window = ttk.PanedWindow(app, orient=tk.HORIZONTAL)
paned_window.pack(fill=tk.BOTH, expand=True,padx=0,pady=0)


# Create two frames for the left and right panels
left_frame = ttk.Frame(paned_window, width=200, height=400, relief=tk.FLAT)
right_frame = ttk.Frame(paned_window, width=400, height=400, relief=tk.FLAT)

# Add frames to the PanedWindow
paned_window.add(left_frame, weight=1)  # The weight determines how additional space is distributed
paned_window.add(right_frame, weight=2)

menu = tk.Menu(app, tearoff=0)
menu.add_command(label="Open File", command=open_file_in_system)
def merge_characters():
    myprint2("merge_characters")
    global currentCharacterMergeFromName
    name = character_table.item(currentCharacterSelectRowId, 'values')[1]
    currentCharacterMergeFromName=name
    myprint2("Merge "+name)
    create_popup(currentResultCharacterOrderMap)

def hide_character():
    myprint2("hide_character")
    global currentDisabledCharacters
    name = character_table.item(currentCharacterSelectRowId, 'values')[1]
    currentDisabledCharacters.append(name)
    #disabled_character_list_table.insert('','end',values=(name,))

    myprint2("Hide "+name)
    reset_tables()
    postProcess(currentBreakdown,currentResultCharacterOrderMap,currentResultEnc,currentResultName,currentResultLinecountMap,currentResultSceneCharacterMap,currentTimelinePath)
 #   show_loading()
#    app.update_idletasks()  # Force the UI to update
#    threading.Thread(target=runJob,args=(currentScriptFilename,"ALL",)).start()



char_menu = tk.Menu(app, tearoff=0)
char_menu.add_command(label="Merge with...", command=merge_characters)
char_menu.add_command(label="Hide", command=hide_character)

#######################################################################################
# Folder tree
folders = ttk.Treeview(left_frame, columns=("Path","Extension",))
folders.heading("#0", text="Nom")
folders.heading("Extension", text="Type")
folders.heading("Path", text="Path")
folders.column("#0", width=240)  # Adjust as needed
folders.column("Path", width=0, stretch=tk.NO)
folders.column("Extension", width=60, stretch=tk.NO)

folders.tag_configure('not_supported', foreground='#cccccc')
folders.tag_configure('supported', foreground='#444444')
#folders.tag_configure('folder', foreground='#6666cc')
bold_font = tkFont.Font( weight="bold")
folders.tag_configure('bold', font=bold_font)

# Default tag with normal background
folders.tag_configure('normal', background='white')

folders.tag_configure('hover', background='#f4f4f4')
style = ttk.Style()
#style.configure('TNotebook.Tab', padding=[10,10,10,10])  # Adjust these values as needed

style.configure("Treeview", rowheight=30)  # Increase the row height
style.configure("Treeview.Item", padding=(3, 4, 3, 4))  # Top and bottom padding
#bold_font = ('Arial', 10, 'bold')
#style.configure("Treeview", font=bold_font)

 
folders.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
folders.bind('<<TreeviewSelect>>', on_folder_select)
folders.bind("<<TreeviewOpen>>", on_folder_open)
# Bind motion event
folders.bind('<Motion>', on_motion)
folders.bind('<Leave>', on_leave)
folders.bind('<Button-1>', toggle_folder)


def on_right_click(event):
    global currentRightclickRowId
    # Identify the row clicked
    try:
        row_id = folders.identify_row(event.y)
        if row_id:
            # Select the row under cursor
            #folders.selection_set(row_id)
            currentRightclickRowId=row_id
            menu.post(event.x_root, event.y_root)  # Show the context menu
    except Exception as e:
        myprint2(e)

def on_character_right_click(event):
    global currentCharacterSelectRowId
    # Identify the row clicked
    try:
        row_id = character_table.identify_row(event.y)
        if row_id:
            # Select the row under cursor
            #folders.selection_set(row_id)
            currentCharacterSelectRowId=row_id
            char_menu.post(event.x_root, event.y_root)  # Show the context menu
    except Exception as e:
        myprint2(e)
folders.bind('<Button-3>', on_right_click)  # Right click on Windows/Linux
folders.bind('<Button-2>', on_right_click) 

# Notebook (tabbed interface)
notebook = ttk.Notebook(right_frame)
notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)



#char_label = ttk.Label(recap_tab, text="Characters", font=('Arial', 30))
#char_label.pack(side=tk.TOP, fill=tk.X)
#recap_tab.bind('<Configure>', resizechart)


# Configure the style of the tab
fontsize=14
if platform.system() == 'Windows':
    style.configure('TNotebook.Tab', background='#f0f0f0', padding=(5, 3), font=('Helvetica', fontsize))
#style.configure('TNotebook.Tab', background='#f0f0f0', padding=(5, 3), font=('Helvetica', fontsize))
# Configure the tab area (optional, for better Windows look)
style.configure('TNotebook', tabposition='nw', background='#f0f0f0')
style.configure('TNotebook', padding=0)  # Removes padding around the tab area

#def on_tab_selected(event):
    #myprint2("Tab selected:", event.widget.select())


#notebook.bind("<<NotebookTabChanged>>", on_tab_selected)
# File preview tab
preview_tab = ttk.Frame(notebook)
notebook.add(preview_tab, text='Texte',image=original_icon, compound=tk.LEFT)
file_preview = Text(preview_tab)
file_preview.pack(fill=tk.BOTH, expand=True)

# Statistics tab
character_tab = ttk.Frame(notebook)

# Create a Treeview widget within the stats_frame for the table

character_table = ttk.Treeview(character_tab, columns=('Order', 'Character','Status', 'Characters','Lines','Scenes'), show='headings')
#character_table = ttk.Treeview(character_tab, columns=('Order', 'Character', 'Lines','Characters','Words','Blocks (50)','Scenes'), show='headings')
# Define the column headings
character_table.heading('Order', text='Order')
character_table.heading('Character', text='Personnage')
character_table.heading('Status', text='Status')
#character_table.heading('Lines', text='Lines')
character_table.heading('Characters', text='Caracteres')
#character_table.heading('Words', text='Words')
character_table.heading('Lines', text='Lignes')
#character_table.heading('Blocks (40)', text='Blocks (40)')
character_table.heading('Scenes', text='Scenes')

# Define the column width and alignment
character_table.column('Order', width=25, anchor='center')
character_table.column('Character', width=200, anchor='w')
character_table.column('Status', width=50, anchor='w')
#character_table.column('Lines', width=50, anchor='w')
character_table.column('Characters', width=50, anchor='w')
#character_table.column('Words', width=50, anchor='w')
character_table.column('Lines', width=50, anchor='w')
character_table.column('Scenes', width=50, anchor='w')

# Pack the Treeview widget with enough space
character_table.pack(fill='both', expand=True)
notebook.add(character_tab, text='Personages',image=char_icon, compound=tk.LEFT)

character_table.tag_configure('hidden', foreground='#999999')
character_table.bind('<Button-3>', on_character_right_click)  # Right click on Windows/Linux
character_table.bind('<Button-2>', on_character_right_click) 












breakdown_tab = ttk.Frame(notebook)
# Create a Treeview widget within the stats_frame for the table
breakdown_table = ttk.Treeview(breakdown_tab, columns=('Line', 'Type', 'Character','Text'), show='headings')
# Define the column headings
breakdown_table.heading('Line', text='Ligne')
breakdown_table.heading('Type', text='Type')
breakdown_table.heading('Character', text='Personnage')
breakdown_table.heading('Text', text='Dialogue')

# Define the column width and alignment
breakdown_table.column('Line', width=25, anchor='w')
breakdown_table.column('Type', width=25, anchor='w')
breakdown_table.column('Character', width=50, anchor='w')
breakdown_table.column('Text', width=200, anchor='w')
# Pack the Treeview widget with enough space
breakdown_table.pack(fill='both', expand=True)
# Configure the tag to change the background color
breakdown_table.tag_configure('nonspeech', background='#fafafa')
breakdown_table.tag_configure('scene', background='#fffec8')
bold_font = tkFont.Font( weight="bold")
breakdown_table.tag_configure('border', background='#444444')  # A lighter shade to simulate space
breakdown_table.tag_configure('bold', font=bold_font)
#notebook.add(breakdown_tab, text='Scenes',image=scene_icon, compound=tk.LEFT)
        












def open_result_folder():
    # Open a folder in Finder using the `open` command
    myprint2("Opening "+currentOutputFolder)
    currentOutputFolderAbs = os.path.abspath(currentOutputFolder)
    myprint2("Absolute path          : "+currentOutputFolderAbs)

  # Check if the folder exists
    if not os.path.exists(currentOutputFolderAbs):
        myprint2(f"Folder does not exist: {currentOutputFolderAbs}")
        return
    try:
        if sys.platform.startswith('darwin'):
            subprocess.run(['open', currentOutputFolderAbs], check=True)
        elif sys.platform.startswith('win32'):
            # Correct approach for Windows
            subprocess.run(['explorer', currentOutputFolderAbs], check=True)
        elif sys.platform.startswith('linux'):
            subprocess.run(['xdg-open', currentOutputFolderAbs], check=True)
    except Exception as e:
        myprint2(f"Error opening folder: {e}")


stats_tab = ttk.Frame(notebook)
# Create a Treeview widget within the stats_frame for the table
stats_table = ttk.Treeview(stats_tab, columns=('Line number',  'Character','Text','Characters'), show='headings')
# Define the column headings
stats_table.heading('Line number', text='Ligne')
stats_table.heading('Character', text='Personnage')
stats_table.heading('Text', text='Dialogue')
stats_table.heading('Characters', text='Caracteres')

# Define the column width and alignment
stats_table.column('Line number', width=25, anchor='center')
stats_table.column('Character', width=100, anchor='w')
stats_table.column('Text', width=200, anchor='w')
stats_table.column('Characters', width=50, anchor='w')

# Pack the Treeview widget with enough space
stats_table.pack(fill='both', expand=True)
# Configure the tag to change the background color
stats_table.tag_configure('nonspeech', background='#fafafa')
stats_table.tag_configure('scene', background='#fffec8')
bold_font = tkFont.Font( weight="bold")
stats_table.tag_configure('border', background='#444444')  # A lighter shade to simulate space
stats_table.tag_configure('bold', font=bold_font)

notebook.add(stats_tab, text="Dialogue dans l'ordre",image=chat_icon, compound=tk.LEFT)


# Statistics tab
character_stats_tab = ttk.Frame(notebook)

# Create a Treeview widget within the stats_frame for the table
cols=('Line #', 'Character','Character (raw)','Line')
for i in countingMethods:
    cols= cols+(countingMethodNames[i],)

style.configure('CleftPanel.TFrame', background='#fafafa')
style.configure('CrightPanel.TFrame', background='#fafafa')


# Create left and right frames (panels) inside the tab
cleft_panel = ttk.Frame(character_stats_tab, borderwidth=0, relief="flat", width=200)
cright_panel = ttk.Frame(character_stats_tab, borderwidth=0, relief="flat")

cleft_panel.configure(style='CleftPanel.TFrame')  # Apply the styled background
cright_panel.configure(style='CrightPanel.TFrame')  # Apply the styled background

# Pack the frames into the tab
#cleft_panel.pack(side='left', fill='y', padx=(0, 20))
#cright_panel.pack(side='right', fill='y', expand=True)

# Configure column weights to make right panel flexible
character_stats_tab.grid_columnconfigure(1, weight=1)
character_stats_tab.grid_rowconfigure(0, weight=1)


character_list_table = ttk.Treeview(cleft_panel, columns=('Character'), show='headings')
# Define the column headings
character_list_table.heading('Character', text='Character')

# Define the column width and alignment
character_list_table.column('Character', width=50, anchor='w')

# Pack the Treeview widget with enough space
character_list_table.pack(fill='both', expand=True)

# Grid frames with padding
cleft_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 10))  # Add padding on the right side of left panel
cright_panel.grid(row=0, column=1, sticky='nsew')  # Automatically spaced by the left panel's padding

# Configure column weights to make right panel flexible
character_stats_tab.grid_columnconfigure(0, weight=0, minsize=200)  # Set minimum size for the left panel
character_stats_tab.grid_columnconfigure(1, weight=1) 
character_stats_tab.grid_rowconfigure(0, weight=1)
def clear_character_stats():
    for item in character_stats_table.get_children():
        character_stats_table.delete(item)
def on_item_selected(event):
    tree = event.widget
    selection = tree.selection()
    item = tree.item(selection)
    record = item['values']
    # Do something with the selection, for example:
    if len(record)==0:
        return
    clear_character_stats()
    character_name=record[0]
    
    character_named = character_name 
    character_list_table.insert('','end',values=(character_named,))
   
    rowtotal=("",character_name,"","TOTAL")       
    total_by_method={}
    for m in countingMethods:
        total_by_method[m]=0

    for item in currentBreakdown:
        line_idx=item['line_idx']
        type_=item['type']
        if(type_=="SPEECH"):

            speech=item['speech']
            character=item['character']
            character_raw=item['character_raw']
            
            filtered_speech=filter_speech(speech)

            if character==character_name:
                #myprint2("    MATCH"+str(speech))

                row=(str(line_idx),character,character_raw, speech)
                for m in countingMethods: 
                    #myprint2("add"+str(m))
                    le=compute_length_by_method(filtered_speech,m)
                    row=row+(str(le),)
                    total_by_method[m]=total_by_method[m]+le
                #myprint2("add"+str(row))
                character_stats_table.insert('','end',values=row)
    for m in countingMethods:
        if m.startswith("BLOCKS"):
            total_by_method[m]=math.ceil(total_by_method[m])

    for m in countingMethods:
        rowtotal=rowtotal+(total_by_method[m],)
    character_stats_table.insert('',0,values=rowtotal,tags=['total'])

character_list_table.bind('<ButtonRelease-1>', on_item_selected)



character_stats_table = ttk.Treeview(cright_panel, columns=cols, show='headings')
# Define the column headings
character_stats_table.heading('Line #', text='Ligne #')
character_stats_table.heading('Character', text='Personnage')
character_stats_table.heading('Character (raw)', text='Personnage (brut)')
character_stats_table.heading('Line', text='Replique')
for i in countingMethods:
    character_stats_table.heading(countingMethodNames[i], text=countingMethodNames[i])


# Define the column width and alignment
character_stats_table.column('Line #', width=25, anchor='center')
character_stats_table.column('Character',  anchor='w', width=0, stretch=tk.NO)
character_stats_table.column('Character (raw)',anchor='w', width=0, stretch=tk.NO)
character_stats_table.column('Line', width=100, anchor='w')
for i in countingMethods:
    character_stats_table.column(countingMethodNames[i], width=25, anchor='w')

# Pack the Treeview widget with enough space
character_stats_table.pack(fill='both', expand=True)
character_stats_table.tag_configure('total', background='#444444',foreground="#ffffff")

notebook.add(character_stats_tab, text='Repliques par personnage',image=chat_icon, compound=tk.LEFT)


# Statistics label
#stats_label = ttk.Label(right_frame, text="Words: 0 Characters: 0", font=('Arial', 12))
#stats_label.pack(side=tk.BOTTOM, fill=tk.X)

export_tab = ttk.Frame(notebook)
# Load folder button
load_button = ttk.Button(export_tab, text="Ouvrir le dossier de resultats...", command=open_result_folder)
load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)
load_button = ttk.Button(export_tab, text="Ouvrir le fichier conversion ...", command=open_result_folder)
load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)

#load_buttonb = ttk.Button(export_tab, text="Show loading ", command=show_loading)
#load_buttonb.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)

# Load folder button
load_button = ttk.Button(export_tab, text="Ouvrir le comptage .xlsx...", command=open_xlsx_recap)
load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)

load_button = ttk.Button(export_tab, text="Ouvrir le detail dialog...", command=open_dialog_recap)
load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)


#load_button = ttk.Button(export_tab, text="Clear chart", command=clear_chart)
#load_button.pack(side=tk.TOP, fill=tk.X,padx=20,pady=20)

#stats_label = ttk.Label(export_tab, text="Disabled characters", font=('Arial', 12))
#stats_label.pack(side=tk.TOP, fill=tk.X)

def merge_with():
    global currentMergedCharacters
    global currentMergedCharactersTo
    global currentMergePopupTable
    
    first_selected_item = currentMergePopupTable.selection()[0]
    first_cell_value = currentMergePopupTable.item(first_selected_item, "values")[0]  
    myprint2(str(first_cell_value))
    mergeto=first_cell_value
    mergefrom=currentCharacterMergeFromName

    currentMergedCharacters[mergefrom]=mergefrom
    if mergeto in currentMergedCharactersTo:
        currentMergedCharactersTo[mergeto].append(mergefrom)
    else:
        currentMergedCharactersTo[mergeto]=(mergefrom,)
    
    currentMergePopupWindow.destroy()

    reset_tables()
    postProcess(currentBreakdown,currentResultCharacterOrderMap,currentResultEnc,currentResultName,currentResultLinecountMap,currentResultSceneCharacterMap,currentTimelinePath)


# Pack the Treeview widget with enough space
#popup_character_list_table.pack(fill='both',expand=True)


def create_popup(character_map):
    global currentMergePopupWindow
    global currentMergePopupTable
    popup = Toplevel(app)
    popup.title("Merge with")
    popup.geometry("200x350")  # Size of the popup window
    currentMergePopupWindow=popup

    popup_character_list_table = ttk.Treeview(popup, columns=('Character'), show='headings')
    # Define the column headings
    popup_character_list_table.heading('Character', text='')

    # Define the column width and alignment
    popup_character_list_table.column('Character', width=50, anchor='w')
   # popup_character_list_table.bind('<Button-1>', merge_with)
    popup_character_list_table.pack(fill='both',expand=True)
    
    for k in character_map:
        popup_character_list_table.insert('', 'end', text=k, values=[k])

    currentMergePopupTable=popup_character_list_table


    # Frame to control the size of the button
    button_frame = tk.Frame(popup, height=40)  # Set the height to 40 pixels
    button_frame.pack(fill='x')  # Fill frame horizontally
    button_frame.pack_propagate(False)  # Prevent frame from resizing to fit contents

    close_btn = tk.Button(button_frame, text="Merge", command=merge_with)
    close_btn.pack(fill='x',expand=True)


#disabled_character_list_table = ttk.Treeview(export_tab, columns=('Character'), show='headings')
# Define the column headings
#disabled_character_list_table.heading('Character', text='Character')

# Define the column width and alignment
#disabled_character_list_table.column('Character', width=50, anchor='w')

# Pack the Treeview widget with enough space
#disabled_character_list_table.pack(fill='both',expand=True)


notebook.add(export_tab, text='Export',image=export_icon, compound=tk.LEFT)

logging.debug("Launch")

currentScriptFolder=settings['SCRIPT_FOLDER']
if currentScriptFolder=="":
    currentScriptFolder=os.getcwd()
load_tree("",currentScriptFolder)

center_window()  # Center the window

app.mainloop()