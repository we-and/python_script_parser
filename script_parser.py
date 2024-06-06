#from PyPDF2 import PdfReader
import re
import os
import math
import pandas as pd
from docx import Document
import csv
import platform

#if platform.system() == 'Windows':
 #   import pythoncom
  #  import win32api
   # import win32com.client

import pdfplumber
#from pyth.plugins.rtf15.reader import Rtf15Reader
#f#rom pyth.plugins.plaintext.writer import PlaintextWriter
import logging
# Initialize COM
#print("CoInit")
#if platform.system() == 'Windows':
 #   pythoncom.CoInitialize()

#script_path="scripts2/YOU CAN'T RUN FOREVER_SCRIPT_VO.txt"
#output_path="YOU CANT RUN FOREVER_SCRIPT_VOc/"
#script_name="YOU CANT RUN FOREVER_SCRIPT_VO"

#script_path="scripts2/ZERO11.txt"
#output_path="ZERO11c/"
#script_name="ZERO11b"
logging.basicConfig(filename='app-parser.log',level=logging.DEBUG,filemode='w')
logging.debug("Script starting...")

def myprint1(s):
    logging.debug(s)
    #print(s)


action_verbs = ["says", "asks", "whispers", "shouts", "murmurs", "exclaims"]

import chardet

characterSeparators=[
        "CHARACTER_SEMICOL_TAB",
        "CHARACTER_TAB",
        "CHARACTER_SPACES"
]
countMethods=[
    "ALL",
    "ALL_NOSPACE"
    "ALL_NOPUNC",
    "ALL_NOSPACE_NOPUNC",
    "ALL_NOAPOS",
    
]


#################################################################
#ENCODING
def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result

def test_encoding(script_path):
    encodings=['windows-1252', 'iso-8859-1', 'utf-16','ascii','utf-8']
    for enc in encodings:
        try:
            with open(script_path, 'r', encoding=enc) as file:
                print("  > Testing encoding  : "+enc)

                for line in file:
                    line = line.strip()  # Remove any leading/trailing whitespace
            return enc
        except UnicodeDecodeError:
            print(f"  > Failed decoding with {enc}")
    return "?"    








#################################################################
# SCENE SEPARATOR
def matches_format_parenthesis_name_timecode(line):
    """Checks if the line matches the specified timecode format."""
    # Regex pattern to match lines that start with '(', include a '-', have a timecode, and end with ')'
    pattern = re.compile(r"\([^\)]*-\s*\d{2}:\d{2}:\d{2}:\d{2}\)$")
    return bool(re.search(pattern, line))

def matches_number_parenthesis_timecode(line):
    """Checks if the given line matches the specified format of 'number (timecode)'."""
    pattern = re.compile(r"^\d+\s+\(\d{2}:\d{2}:\d{2}:\d{2}\)$")
    return bool(re.match(pattern, line))

def extract_scene_name2(line):
    """Extracts the scene name or identifier from a line, which is the part before the first parenthesis."""
    # Split the line at the first space or parenthesis
    parts = line.split(' (', 1)  # Splits the string at the first occurrence of ' ('
    if parts:
        return parts[0].strip()  # Return the first part, stripping any extra whitespace
    return None 

def is_scene_line(line):
    isSceneLine=matches_format_parenthesis_name_timecode(line) or matches_number_parenthesis_timecode(line)
    #print("IsScene    "+str(isSceneLine)+" "+line)
    return isSceneLine

def extract_scene_name1(line):
    """Extracts the scene name from a line formatted with timecode."""
    # Regex pattern to capture text between the opening parenthesis and the dash
    pattern = re.compile(r"\(([^-]*)")
    match = re.search(pattern, line)
    if match:
        # Strip any leading/trailing whitespace from the captured group
        return match.group(1).strip()
    return None  # Return None if no match is found or the format is incorrect

def extract_scene_name(line,scene_separator,current_scene_count):
    if scene_separator=="NAME_PARENTHESIS_TIMECODE"  :
        if matches_number_parenthesis_timecode(line):
            return extract_scene_name2(line)
    elif  scene_separator=="PARENTHESIS_NAME_TIMECODE":
        if matches_format_parenthesis_name_timecode(line):
            return extract_scene_name1(line)
    elif scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
        return "Scene "+str(current_scene_count)
    else:
        return "?"

def isSeparatorNameParenthesisTimecode(scene_separator):
    return scene_separator=="NAME_PARENTHESIS_TIMECODE"
def isSeparatorParenthesisNameTimecode(scene_separator):
    return scene_separator=="PARENTHESIS_NAME_TIMECODE"
def isSeparatorEmptyLinesTimecode(scene_separator):
    return scene_separator=="EMPTYLINES_SCENE_SEPARATOR"

def getCharacterSeparator(script_path,encod):
    print("getCharacterSeparator")
    best="?"
    bestVal=0.0
    nLines=0
    for sep in characterSeparators:
        nLines=0
        nMatches=0
        with open(script_path, 'r', encoding=encod) as file:
            for line in file:
                line = line.strip()
                if len(line)>0:
                    nLines=    nLines+1

                    if sep=="CHARACTER_SEMICOL_TAB":
                        is_match=matches_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)
                        if is_match:
                            nMatches=nMatches+1
                    elif sep=="CHARACTER_SPACES":
                        is_match=matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)
                        if is_match:
                            nMatches=nMatches+1
                    elif sep=="CHARACTER_TAB":
                        is_match=matches_charactername_NAME_ATLEAST1TAB_TEXT(line)
                        if is_match:
                            nMatches=nMatches+1
        pc=round(100*nMatches/nLines)
        print("  > Test character sep:"+sep+" " +str(nMatches)+"/"+str(nLines),str(pc))
        if pc>bestVal:
            bestVal=pc
            best=sep

    return best
def getSceneSeparator(script_path,encod):
    mode="?"
    # Open the file and process each line
    
    with open(script_path, 'r', encoding=encod) as file:
        for line in file:
            line = line.strip()  # Remove any leading/trailing whitespace
            if matches_format_parenthesis_name_timecode(line):
                print("PARENTHESIS_NAME_TIMECODE")
                return "PARENTHESIS_NAME_TIMECODE"
            
            elif matches_number_parenthesis_timecode(line):
                print("NAME_PARENTHESIS_TIMECODE")
                return "NAME_PARENTHESIS_TIMECODE"

    if mode=="?":
        n_sets_of_empty_lines=count_consecutive_empty_lines(script_path,2,encod)
        #print("check empty lines count"+str(n_sets_of_empty_lines))
        if n_sets_of_empty_lines>1:
            print( "Found EMPTYLINES_SCENE_SEPARATOR in "+line)
            return "EMPTYLINES_SCENE_SEPARATOR"

    return mode    


#################################################################
# CHARACTER SEPARATOR
def is_matching_character_speaking(line,character_mode):
    if character_mode=="CHARACTER_TAB":
        return matches_charactername_NAME_ATLEAST1TAB_TEXT(line)
    elif character_mode=="CHARACTER_SPACES": 
        return matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return matches_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)  
    else:
        print("ERROR wrong mode="+character_mode)


    """Checks if the line indicates a character speaking."""
    # Regex pattern to match lines that start with text, followed by a tab, then more text
    pattern = re.compile(r"^\S+\s+\t.*$")
    pattern = re.compile(r"^\S+\s*\t.*$")

    return bool(re.match(pattern, line))

def is_character_speaking(line,character_mode):
    is_match= is_matching_character_speaking(line,character_mode) 
    if is_match:
        name= extract_character_name(line,character_mode)
        return not is_didascalie(name) and not is_ambiance(name)
    else:
        return False


def extract_speech(line,character_mode,character_name):
    if character_mode=="CHARACTER_TAB":
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTER_SPACES": 
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return extract_speech_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line,character_name)  
    else:
        print("ERROR wrong mode="+str(character_mode))
        exit()

def extract_character_name(line,character_mode):
    if character_mode=="CHARACTER_TAB":
        return extract_charactername_NAME_ATLEAST1TAB_TEXT(line)
    elif character_mode=="CHARACTER_SPACES": 
        return extract_charactername_NAME_ATLEAST8SPACES_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return extract_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)  
    else:
        print("ERROR wrong mode="+str(character_mode))
        exit()

def matches_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(text):
    # Define the regex pattern
    # ^ starts the match at the beginning of the line
    # [\w\s]+ matches one or more word characters or spaces to include names with spaces
    # : matches the literal colon
    # \s* matches zero or more whitespace characters (spaces or tabs)
    # \t matches a tab
    # .+ matches one or more of any character (the text following the tab)
    # $ ensures the pattern goes to the end of the line
    pattern = r'^[\w\s]+:\s*\t.+'

    # Use re.match to check if the start of the string matches the pattern
    if re.match(pattern, text):
        return True
    else:
        return False


def matches_charactername_NAME_ATLEAST8SPACES_TEXT(text):
    # Define the regex pattern:
    # ^ starts the match at the beginning of the line
    # (.+) matches one or more of any character (the first text block), captured for potential use
    # {8,} specifies at least 8 spaces
    # (.+) matches one or more of any character following the spaces (the second text block)
    #pattern = r'^(.+)\s{8,}(.+)$'
    pattern = r'^(.+?)\s{8,}(.+)$'
    # Use re.match to check if the whole string matches the pattern
    if re.match(pattern, text):
        return True
    else:
        return False

def matches_charactername_NAME_ATLEAST1TAB_TEXT(text):
    # Define the regex pattern:
    # ^ starts the match at the beginning of the line
    # (.+) captures one or more characters as the first part of text
    # \t+ matches one or more tab characters
    # (.+) captures one or more characters as the second part of text
    # $ ensures the match extends to the end of the line
    pattern = r'^(.+)\t+(.+)$'

    # Use re.match to check if the whole string matches the pattern
    if re.match(pattern, text):
        return True
    else:
        return False
def extract_speech_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line,character_name):
    right=line.replace(character_name,"")
    if right.startswith(':'):
        return right[1:].strip()
    return right
def extract_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line):
    # Define the regex pattern:
    # ^ asserts the start of the line
    # ([\w\s]+) captures a group of word characters or spaces which will be the name
    # : matches the literal colon
    # \s* matches zero or more spaces
    # \t matches a literal tab
    # .+ matches one or more of any characters (the following text)
    pattern = r'^([\w\s]+):\s*\t.+'

    # Use re.search to find the first occurrence of the pattern
    match = re.search(pattern, line)
    if match:
        res=match.group(1).strip()
        # Return the first captured group, which is the name, stripping any extra spaces
        return res
    else:
        return None

def extract_charactername_NAME_ATLEAST8SPACES_TEXT(line):
    """
    Extracts the first part of the line which is uppercase text,
    given the line format is uppercase text followed by at least 8 spaces and more text.
    Returns the uppercase text if the pattern matches, otherwise returns None.
    """
    #match = re.match(r"^([A-Z]+)\s{8,}.*$", line)
    match = re.match(r"^([A-Z ]+)\s{8,}.*$", line)
    if match:
        return match.group(1)
    return None

def extract_charactername_NAME_ATLEAST1TAB_TEXT(line):
    """Extracts the character name from a line where the name is followed by a tab and then dialogue."""
    # Split the line at the first tab character
    parts = line.split('\t', 1)  # The '1' limits the split to the first occurrence of '\t'
    if len(parts) > 1:
        return parts[0].strip()  # Return the first part, stripping any extra whitespace
    return None  # Return None if no tab is found, indicating an improperly formatted line

#################################################################
# CHARACTER UTILS
def is_didascalie(name):
    return name=="DIDASCALIES"
def is_ambiance(name):
    return name=="AMBIANCE"
def filter_character_name(line):
    if line:
        if "(O.S)" in line:
            line=line.replace("(O.S)","")
        if "(O.S.)" in line:
            line=line.replace("(O.S.)","")
        if "(CONT'D)" in line:
            line=line.replace("(CONT'D)","")    
        if line.endswith(':'):
            line= line[:-1]
        if line.endswith(')'):
            line= line[:-1]
    return line
#    return line



#################################################################
# UTILS


def convert_csv_to_xlsx(csv_file_path, xlsx_file_path, script_name,encoding_used):
    print("convert_csv_to_xlsx > 0")
    # Read the CSV file
    df = pd.read_csv(csv_file_path,header=None,encoding=encoding_used)

    # Write the DataFrame to an Excel file
    #print("convert_csv_to_xlsx > Write to "+xlsx_file_path)

    header_rows = pd.DataFrame([
        [None, 'Header 1', None, 'Header Information Across Columns'],  # Merge cells will be across 1 & 4
        ['Role', 'Line count', 'Characters', 'Blocks']
    ])
    #print("convert_csv_to_xlsx > 1")
    
    # Concatenate the header rows and the original data
    # The ignore_index=True option reindexes the new DataFrame
    df = pd.concat([header_rows, df], ignore_index=True)
    #print("convert_csv_to_xlsx > 2")

    # Write the DataFrame to an Excel file
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        #print("convert_csv_to_xlsx > 3")

        df.to_excel(writer, index=False, sheet_name='Sheet1')
        #print("convert_csv_to_xlsx > 4")

        # Load the workbook and sheet for modification
        workbook = writer.book
        sheet = workbook['Sheet1']
        #print("convert_csv_to_xlsx > 5")

        # Merge cells in the first and second new rows
        # Assuming you want to merge from the first to the last column
        sheet.merge_cells('A1:D1')  # Modify range according to your number of columns
        sheet.merge_cells('A2:D2')  # Modify this as needed
        sheet['A1'] = script_name
        sheet['A2'] = "Length: "
        #print("convert_csv_to_xlsx > 6")

    print("convert_csv_to_xlsx > done")

#    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')

def write_character_map_to_file(character_map, filename):
    print(" > Write map to "+filename)
    """Writes the character to scene map to a specified file."""
    with open(filename, 'w', encoding='utf-8') as file:
        for character, scenes in character_map.items():
            file.write(f"{character}: {scenes}\n")

def find_first_uppercase_sequence(line):
    """Finds the first sequence of contiguous uppercase words in a line."""
    # Regex pattern to match the first sequence of contiguous uppercase words separated by spaces
    pattern = re.compile(r'\b([A-Z]+(?:\s[A-Z]+)*)\b')
    match = re.search(pattern, line)
    if match:
        return match.group(0)
    return None  # Return None if no uppercase sequence is found

def count_consecutive_empty_lines(file_path, n,encod):
    """Counts occurrences of exactly n consecutive empty lines in a file."""
    i=1
    with open(file_path, 'r', encoding=encod) as file:
        count_empty = 0
        occurrences = 0
        previous_empty = False

        for line in file:
            # Check if the current line is empty or contains only whitespace
            if line.strip() == '':
                count_empty += 1
                previous_empty = True
                #print(str(i)+"empty line")
            else:
                #print(str(i)+"not empty pre="+str(previous_empty)+" co="+str(count_empty))
                if previous_empty and count_empty >= n:
                    occurrences += 1
                    #print(str(i)+" ADD OCC")

                count_empty = 0
                previous_empty = False
            i=i+1
        # Check at the end of the file if the last lines were empty
        if count_empty >= n:
            occurrences += 1
            #print(str(i)+" ADD OCC")

    return occurrences


def sort_dict_values(d):
    sorted_dict = {}
    for key, value_set in d.items():
        try:
            # Attempt to sort assuming all values are numeric strings
            sorted_list = sorted(value_set, key=int)
        except ValueError:
            # Handle the case where values are not all numeric
            print(f"Non-numeric values found in the set for key '{key}'. Values: {value_set}")
            # Optionally sort only numeric values or handle differently
            numeric_values = [val for val in value_set if val.isdigit()]
            sorted_list = sorted(numeric_values, key=int)
        sorted_dict[key] = sorted_list
    return sorted_dict

def compute_length(line,method):
    if method=="ALL":
        return len(line)
    elif method=="ALL_NOSPACE":
        return len(line.replace(" ",""))
    else:
        return len(line)
    

from docx import Document

def read_docx(file_path):
    # Load the document
    doc = Document(file_path)
    
    # Iterate through each paragraph in the document
    for para in doc.paragraphs:
        print(para.text)
    return doc
def is_supported_extension(ext):
    ext=ext.lower()
    return ext==".txt" or ext==".docx" or ext==".doc" or ext==".rtf" or ext==".pdf"



def convert_docx_combined_continuity(file,table):
    print("Conversion mode       : COMBINED_CONTINUITY")
    current_character=""
    mode = "linear"
    titles=[]
    for row in table.rows[1:]:
        n=len(row.cells)
        if n>7:
            title=row.cells[n-1].text
            dur=row.cells[n-2].text
            is_title= len(dur)>0 and len(title)>0
            is_speech= len(dur)==0 and len(title)>0


            s=""
            if is_title:
                if "/" in title:
                    mode="split"
                    titles=title.split("/")
                else:
                    mode="linear"
                    title=clean_character_name(title)
                    current_character=title
            if is_speech:
                if mode=="split":
                    speeches=title.split("\n")

                    idx=0
  #                  print("split mode speech"+title+str(len(speeches)))

                    for i in speeches:
                        current_character=titles[idx]
                        speech=speeches[idx].strip()
                        current_character=clean_character_name(current_character)
                        if speech.startswith("-"):
                            speech=speech[1:].strip()
                        s=current_character+"\t"+speech+"\n"
                        file.write(s)
                        idx=idx+1
                    mode="linear"

                elif mode=="linear":
                    current_character=clean_character_name(current_character)
                    s=current_character+"\t"+title+"\n"
                    #print("linear mode add "+s)
                        
                    file.write(s)
def match_uppercase_semicolon(text):
    # Regex pattern to match text in uppercase ending with a semicolon
    pattern = r'^[A-Z\d\s]+:$'
    
    # Use re.match to check if the text matches the pattern
    if re.match(pattern, text):
        return True
    else:
        return False
def remove_semicolon(s):
    """
    Remove a semicolon at the end of a string if it exists.
    
    Args:
    s (str): The string from which to remove the semicolon if present at the end.

    Returns:
    str: The string without the trailing semicolon.
    """
    # Check if the string ends with a semicolon and remove it
    if s.endswith(':'):
        return s[:-1]  # Slice off the last character
    return s  # Return the original string if there's no semicolon at the end

def get_text_without_parentheses(input_string):
    pattern = r'\([^()]*\)'
    # Use re.sub() to replace the occurrences with an empty string
    result_string = re.sub(pattern, '', input_string)
    return result_string
def remove_text_in_brackets(text):
    """
    Remove all text within square brackets from the provided string.

    Args:
    text (str): The string from which to remove the text in brackets.

    Returns:
    str: The string with text in brackets removed.
    """
    # Regular expression to find text in square brackets
    pattern = r'\[.*?\]'
    # Use re.sub() to replace all occurrences of the pattern with an empty string
    cleaned_text = re.sub(pattern, '', text)
    return cleaned_text
def filter_speech(input):
    s=get_text_without_parentheses(input)
    s=remove_text_in_brackets(s)
    s=s.replace("â€™","'")
    s=s.replace("â€¦ ",".")
    return s

def filter_speech_keepbrackets(input):
    s=get_text_without_parentheses(input)
    s=s.replace("â€™","'")
    s=s.replace("â€¦ ",".")
    return s

def convert_docx_dialogwithspeakerid(file,table,dialogwithspeakerid):
    idx=1
    row_idx=1
    current_character=""
    cumulated_speech=""
    ndisp=300
    
    for row in table.rows[1:]:
            show=True#row_idx<200 and row_idx>160
            if show:
                myprint1("---------------------------------------------")
                myprint1("Row "+str(row_idx))
                myprint1("Idx "+str(idx))
                
            row_idx=row_idx+1
            scenedesc=row.cells[dialogwithspeakerid].text.strip()        
            if show:
                myprint1(scenedesc)
            parts=scenedesc.split("\n")
            upperpart=""
            nonupperpart=""
            for k in parts:

                if remove_text_in_brackets(k).isupper():
                    if show:
                        myprint1('+UP '+str(k))
                    upperpart=upperpart+" "+k
                else:
                    if show:
                        myprint1('+LO '+str(k))

                    nonupperpart=nonupperpart+" "+k
            part_idx=1
            upperpart=upperpart.strip()
            nonupperpart=nonupperpart.strip()
            if show:
                myprint1("upperpart:  "+upperpart)
                myprint1("lowerpart: " +nonupperpart)
            
            if is_character_name_valid(upperpart):
                if show:
                        myprint1("valid character candidate"+upperpart)
                upperpart=remove_text_in_brackets(upperpart).strip()
                if match_uppercase_semicolon(upperpart) or upperpart.isupper():
                        # if  :
                            if show:
                                    myprint1("IS CHARACTER YES")
                            #flush cunulated speech
                            #if len(cumulated_speech)>0:
                            #   s=current_character+"\t"+cumulated_speech+"\n"
                            #  file.write(s)
                            # idx=idx+1
                            if len(current_character)>0 and len(cumulated_speech)>0:
                                s=current_character+"\t"+cumulated_speech+"\n"
                                file.write(s)
                                myprint1(">>>>>>>>>>> "+s)
                                myprint1("reset cum")
                                cumulated_speech=""
                            current_character=remove_semicolon(upperpart) 
                            current_character=clean_character_name(current_character)
                            if show:
                                myprint1("new character "+current_character)
                else:
                    if show:
                        myprint1("Not a character")

                if len(nonupperpart)>0:
                    if show:
                        myprint1("has noupper"+nonupperpart)
                # if  idx<90:
                    #    print("part "+str(part_idx)+":"+part)
                    part=filter_speech_keepbrackets(nonupperpart)
                    speech=part.strip() 
                    myprint1("Add to cummulated"+speech)
                    if len(current_character)>0:
                        cumulated_speech=cumulated_speech+" "+speech

                    if show:
                        myprint1("final"+nonupperpart)
            else:
                    if show:
                        myprint1("Not a valid character"+upperpart)
                    
    if len(cumulated_speech)>0:
            s=current_character+"\t"+cumulated_speech+"\n"
            file.write(s)
            idx=idx+1
            myprint1(">>>>>>>>>>> "+s)

def convert_docx_scenedescription(file,table,sceneDescriptionIdx,titlesIdx):
    print("Conversion mode       : SCENEDESCRIPTION")
    current_character=""
    idx=1
    cumulated_speech=""
    row_idx=1

    for row in table.rows[1:]:
        if row_idx<35:
                print("---------------------------------------------")
                print("Row "+str(row_idx))
                scenedesc=row.cells[titlesIdx].text.strip()        
                print("Row content"+str(scenedesc))
        row_idx=row_idx+1
    row_idx=1
    for row in table.rows[1:]:
            if idx<10:
                print("---------------------------------------------")
                print("Row "+str(row_idx))
            row_idx=row_idx+1
            scenedesc=row.cells[titlesIdx].text.strip()        
            if idx<100:
                print(scenedesc)
            parts=scenedesc.split("\n")
            part_idx=1
            for part in parts:
                if idx<10:
                    print("part "+str(part_idx)+":"+part)
                
                part=part.strip()
                part_idx=part_idx+1
                if len(part)>0:
#                    if idx<10:
 #                       print("part "+str(part_idx)+":"+part)
                    part=filter_speech(part)
                    if match_uppercase_semicolon(part):
                        #flush cunulated speech
                        if len(cumulated_speech)>0:
                            s=current_character+"\t"+cumulated_speech+"\n"
                            file.write(s)
                            if idx<10:
                                print(">>>>>>>>> "+current_character+"\t"+cumulated_speech)
                            idx=idx+1
                            cumulated_speech=""
                        current_character=remove_semicolon(part) 
                        if idx<10:
                            print("new character "+current_character)
                    else:
                        if not part.isupper():                    
                            speech=part.strip()
                            current_character=clean_character_name(current_character)
                            if len(current_character)>0 and len(speech)>0:
                                cumulated_speech=cumulated_speech+" "+speech
                                #s=current_character+"\t"+speech+"\n"
                                #file.write(s)
    if len(cumulated_speech)>0:
        s=current_character+"\t"+cumulated_speech+"\n"
        file.write(s)
                                                    
def clean_character_name(title):
    title = title.replace("to herself","")
    title = title.replace(" said","")
    return title.strip().upper()

def convert_rtf_to_txt(file_path,currentOutputFolder,encoding):
    if platform.system() != 'Windows':
        return 
    pythoncom.CoInitialize()

    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        doc.Activate()

        # Extract text
        text = doc.Content.Text
         # Rename path with .docx
        new_file_abs = os.path.abspath(file_path)
        base=(os.path.basename(new_file_abs).replace(".rtf", ".converted.txt"))
        new_file_abs = currentOutputFolder+"\\"+ base

        # Save and Close
#        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=16)  # FileFormat=16 for .docx, not .doc
        doc.Close(False)
        txt_encoding="utf-8"
        with open(new_file_abs, 'w', encoding=txt_encoding) as file:
            file.write(text)

        return new_file_abs,txt_encoding

    finally:
        # Make sure to uninitialize COM
        pythoncom.CoUninitialize()
    return ""

def convert_docx_plain_text(file,doc):
        for para in doc.paragraphs:
            # Write the text of each paragraph to the file followed by a newline
            file.write(para.text + '\n')

def convert_docx_characterid_and_dialogue(file,table,dialogueCol,characterIdCol):
    print("Conversion mode       : CHARACTERID_AND_DIALOGUE")
    # Iterate through each row in the table
    current_character=""
    for row in table.rows[1:]:
        cell_texts = [cell.text for cell in row.cells]
            # Join all cell text into a single string
        row_text = ' | '.join(cell_texts)
        
        
        dialogue=row.cells[dialogueCol].text.strip()
        character=row.cells[characterIdCol].text.strip()

        dialogue=dialogue.replace("\n","")
        print("----------")
        print("row"+row_text)
        print("dialogue"+dialogue)
        print("character"+character)
        if "(O.S)" in character:
            character=character.replace("(O.S)","")
        if "(O.S.)" in character:
            character=character.replace("(O.S.)","")

        if len(character)>0:
            current_character=character
        if len(dialogue)>0:  
            is_didascalie=dialogue.startswith("(")
            if not is_didascalie: 
                dialogue=filter_text(dialogue)             
                if current_character=="":
                    current_character="__VOICEOVER"
                s=current_character+"\t"+dialogue+"\n"  # New line after each row
                print("Add "+ current_character + " "+ dialogue)
                file.write(s)


def convert_doc_to_docx(doc_path,currentOutputFolder):
    """Converts a .doc file to .docx"""
    
    if platform.system() != 'Windows':
        return 

    pythoncom.CoInitialize()

    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_path)
        doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(doc_path)
        base=(os.path.basename(new_file_abs).replace(".doc", ".docx"))
        new_file_abs = currentOutputFolder+"\\"+ base

        # Save and Close
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=16)  # FileFormat=16 for .docx, not .doc
        doc.Close(False)

        return new_file_abs

    finally:
        # Make sure to uninitialize COM
        pythoncom.CoUninitialize()
    return ""

                      
def convert_pdf_to_txt(file_path,absCurrentOutputFolder,encoding):
    print("convert_pdf_to_txt")
    print("currentOutputFolder             :"+absCurrentOutputFolder)
    print("Input             :"+file_path)
    converted_file_path=""
    if ".pdf" in file_path.lower() :
        converted_file_path=absCurrentOutputFolder+"\\"+ (os.path.basename(file_path).lower().replace(".pdf",".converted.txt"))
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        with pdfplumber.open(file_path) as pdf:
          page_idx=1  
          for page in pdf.pages[1:]:

            page_idx=page_idx+1
            if page_idx<40000:
                # Extract tables from the page
                print("########################################################")
                print("convert_pdf_to_txt page"+str(page_idx))
                crop_coords = [0.68,0.20,0.95,0.90]
                my_width = page.width
                my_height = page.height
                my_bbox = (crop_coords[0]*float(my_width), crop_coords[1]*float(my_height), crop_coords[2]*float(my_width), crop_coords[3]*float(my_height))
                page = page.crop(bbox=my_bbox)
                text = str(page.extract_text())
                print(text)
                textlines=text.split("\n")
                current_character=""
                current_characters_split=[]
                speech=""
                mode="linear"
                for line in textlines:
                    if line.isupper():
                        if " THEN " in line or "," in line:
                            mode="split"
                            if " THEN " in line:
                                current_characters_split=line.split(" THEN ")   
                            elif "," in line: 
                                current_characters_split=line.split(",")    
                            
                            current_character=line
                        else:
                            current_character=line
                            mode="linear"
                    else:
                        speech=line
                        if mode=="split":
                            speeches=speech.split("\n")
                            charidx=0
                            for sp in speeches:
                                sp=sp.replace("- ","")
                                current_character=current_characters_split[charidx]
                                s=current_character+"\t"+sp+"\n"  # New line after each row
                                print("Add "+ current_character + " "+ speech)
                                file.write(s)
                                charidx=charidx+1

                        if mode=="linear":
                            if current_character=="":
                                current_character="__VOICEOVER"
                            s=current_character+"\t"+speech+"\n"  # New line after each row
                            print("Add "+ current_character + " "+ speech)
                            file.write(s)

                if False:
                    tables = page.extract_tables()
        #            tables = page.extract_tables(table_settings={"vertical_strategy": "text",    "horizontal_strategy": "text"})
                    pdf_mode="?"
                    print("convert_pdf_to_txt tables = "+str(len(tables)))
                    
                    for table in tables:
                        # Add a table to the Word document
                        if table:  # Check if the table is not empty

                            print("convert_pdf_to_txt testheader")
                            print(table)
                            headerSuccess=False
                            for i in range(3):
                                print("test row"+str(i))
                                header=table[i]
                                success=test_pdf_header(file,table,header)
                                if success:
                                    headerSuccess=True
                                    break
                            if not headerSuccess:
                                return ""
    print("Converted")                       
    print(converted_file_path)
    return converted_file_path,encoding



def convert_pdf_title(file,table,titleIdx):
    print("ROW LEN="+str(len(table[0]))+str(table[0]))
    for row in table[1:]:  # Skip header
        lastIdx=-1
        for i in range(len(row)):
            if row[i]:
                print("has data"+str(row[i]))
                lastIdx=i
        print("---------------------------------"+str(lastIdx))
        title = row[lastIdx]
        print("item "+str(i)+" "+str(row[i]))
        print("ROW LEN="+str(len(row))+str(row))
        print(str(row))
        print(">>>>>> TITLE"+str(title))
        if title:
            parts=title.split("\n")
            character=""
            speech=""
            hascharacter=False
            for part in parts:
                if part.isupper():
                    character=character+" " +part
                else:
                    hascharacter=True
                    speech=speech+speech

            file.write(f"{character}\t{speech}\n")

def test_pdf_header(file,table,header):
            print("test_pdf_headers"+str(header))
            dialogueCol=-1
            characterIdCol=-1
            titlesCol=-1
    
            print("headers"+str(header))
            if 'CHARACTER' in header and 'DIALOGUE' in header:
                pdf_mode="PDF_CHARACTER_DIALOGUE"
            if 'Title' in header:
                pdf_mode="PDF_TITLE"
            if "CHARACTER" in header:
                characterIdCol = header.index('CHARACTER')
            if "Title" in header:
                titlesCol = header.index('Title')
            if "DIALOGUE" in header:
                dialogueCol = header.index('DIALOGUE')

            pdf_mode_title=titlesCol>-1
            if pdf_mode_title:
                print("test_pdf_header pdf_mode_title")
                convert_pdf_title(file,table,titlesCol)
                return True




def test_word_header(file,table,header,forceMode="",forceCols={}):
            print("test word header")
            print("forceMode"+str(forceMode))
            print("forceCols"+str(forceCols))
            dialogueCol=-1
            characterIdCol=-1
            combinedContinuityCol=-1
            titlesCol=-1
            dialogWithSpeakerId=-1
            sceneDescriptionCol=-1
            idx=0

            docx_mode_dialogue_characterid= False
            docx_mode_combined_continuity= False
            docx_mode_scenedescription= False
            docx_mode_dialogwithspeakerid=False


            if len(forceMode)>0:
                if forceMode=="DETECT_CHARACTER_DIALOG":
                    dialogueCol=forceCols['DIALOG']
                    characterIdCol=forceCols['CHARACTER']
                    docx_mode_dialogue_characterid=True
                    print("CHARACTER "+str(characterIdCol))
                    print("DIALOG "+str(dialogueCol))
                else:
                    print("UNKNOWN FORCE MODE")
            else:
                for cell in header.cells:
                    t=cell.text.strip()
                    print("Header cell"+str(t))
                    if t=="CHARACTER ID":
                        characterIdCol=idx
                    if t=="CHARACTER":
                        characterIdCol=idx
                    elif t=="ROLE":
                        characterIdCol=idx
                    elif t=="DIALOGUE":
                        dialogueCol=idx
                    elif t=="Scene Description":
                        sceneDescriptionCol=idx
                    elif t=="Titles":
                        titlesCol=idx
                    elif t=="Dialog With \nSpeaker Id":
                        dialogWithSpeakerId=idx
                    elif t=="COMBINED CONTINUITY":
                        combinedContinuityCol=idx
                    idx=idx+1

                docx_mode_dialogue_characterid= dialogueCol>-1 and characterIdCol>-1
                docx_mode_combined_continuity= combinedContinuityCol>-1
                docx_mode_scenedescription= sceneDescriptionCol>-1 and titlesCol>-1
                docx_mode_dialogwithspeakerid=dialogWithSpeakerId>-1
                if docx_mode_dialogue_characterid or docx_mode_combined_continuity or docx_mode_scenedescription or docx_mode_dialogwithspeakerid:
                    print("Headers found")
                else:
                    print("Headers not found")
                    print(" CharacterId"+str(characterIdCol))
                    print(" sceneDescriptionCol"+str(sceneDescriptionCol))
                    print(" dialogueCol"+str(dialogueCol))
                    print(" titlesCol"+str(titlesCol))
                    print(" dialogWithSpeakerId"+str(dialogWithSpeakerId))
                    print(" combinedContinuityCol"+str(combinedContinuityCol))
                    return False
                
            if docx_mode_dialogue_characterid:
                convert_docx_characterid_and_dialogue(file,table,dialogueCol,characterIdCol)
                return True
            elif docx_mode_scenedescription:
                convert_docx_scenedescription(file,table,sceneDescriptionCol,titlesCol)
                return True
            elif docx_mode_dialogwithspeakerid:
                convert_docx_dialogwithspeakerid(file,table,dialogWithSpeakerId)
                return True
            elif docx_mode_combined_continuity:
                convert_docx_combined_continuity(file,table)
                return True
            else:
                print("Tables but no ")
                return False


def convert_word_to_txt(file_path,absCurrentOutputFolder,forceMode="",forceCols={}):
    print("convert_docx_to_txt")
    print("currentOutputFolder             :"+absCurrentOutputFolder)
    print("Input             :"+file_path)
    converted_file_path=""
    if ".docx" in file_path:
        converted_file_path=absCurrentOutputFolder+"\\"+ (os.path.basename(file_path).replace(".docx",".converted.txt"))
    elif ".doc" in file_path:
        print("Convert doc to docx  ")
        docx_file_path = convert_doc_to_docx(file_path,absCurrentOutputFolder,forceMode=forceMode,forceCols=forceCols)
        print("Output         :"+os.path.abspath(docx_file_path))
        converted_file_path=os.path.abspath(docx_file_path).replace(".docx",".converted.txt")
        file_path=docx_file_path
    print("Converted file path : "+converted_file_path)    
    print("Doc opening"+file_path)
    doc = Document(file_path)
    print("Doc opened")

    with open(converted_file_path, 'w', encoding='utf-8') as file:
        # Check if there are any tables in the document
        if len(doc.tables) > 0:
            # Get the first table
            table = doc.tables[0]
            
            row_idx=0
            headerSuccess=False
            for i in range(3):
                header=table.rows[i]
                success=test_word_header(file,table,header,forceMode=forceMode,forceCols=forceCols)
                if success:
                    headerSuccess=True
                    break
            if not headerSuccess:
                return ""

        else:
            print("No tables")
            convert_docx_plain_text(file,doc)
    return converted_file_path

def get_all_characters(breakdown):
    #print("get_all_characters")
    all_characters=[]
    for item in breakdown:
        if item["type"]=="SPEECH":
            character=item["character"]
            if character==None:
                print("ERR")
                exit()

            if not character in all_characters:
                all_characters.append(character)
    return all_characters

splitables=[" TALK TO "," TO "]
def hasSplitable(character):
    for i in splitables:
        if i in character:
            return i
    return ""

def has_duplicates(lst):
    return len(lst) != len(set(lst))

def indices_of_duplicates(lst):
    from collections import defaultdict
    index_map = defaultdict(list)  # Stores list of indices for each item
    for index, item in enumerate(lst):
        index_map[item].append(index)
    
    # Filter items that appear more than once and collect their indices
    return {item: indices for item, indices in index_map.items() if len(indices) > 1}

def map_semi_duplicates(names):
    normalized_map = {}  # Maps normalized names to their first occurrence
    duplicates = {}  # Stores mappings of semi-duplicate entries

    for name in names:
        normalized = name.replace(" ", "")  # Remove spaces to normalize
        if normalized in normalized_map:
            # Map current name to the first occurrence of this normalized form
            duplicates[name] = normalized_map[normalized]
        else:
            # Store the first occurrence of this normalized form
            normalized_map[normalized] = name

    return duplicates

def is_action_verb_charactername(charactername):
    # List of action verbs
   
  
    # Regex to match "<NAME> <ACTION>"
    pattern = r"(\b\w+\b) (" + '|'.join(action_verbs) + r")\b"

    # Search for matches
    match = re.search(pattern, charactername)

    if match:
        print("Match found:", match.groups())
    else:
        print("No match found.")

def extract_character_and_action(charactername, action_verbs):
    # Create a regex pattern that matches a word followed by any of the action verbs
    pattern = r"(\b\w+\b) (" + '|'.join(action_verbs) + r")\b"

    # Search for matches in the provided line
    match = re.search(pattern, charactername)
    if match:
        # Returns the character's name and the action verb
        return match.groups()
    else:
        return None

def merge_breakdown_character_by_replacelist(breakdown,replace_list):
    print("merge_breakdown_character_by_replacelist")
    checkIfAlreadyNamed=False
    for item in breakdown:
        if item["type"]=="SPEECH":
            character=item["character"]
            if character in replace_list:
                firstchar=replace_list[character]
                item['character']=firstchar                   

    return breakdown

def merge_breakdown_character_talking_to(breakdown,all_characters):
    print("merge_breakdown_character_talking_to")
    replaceList={}
    checkIfAlreadyNamed=False
    for item in breakdown:
        #sprint(str(item))
        if item["type"]=="SPEECH":
            character=item["character"]
            if character!=character.strip():
                character=character.strip()

            splitable=hasSplitable(character)
            if splitable!="":
                #print(" has to"+character)
                characters=character.split(splitable)
                #print(" split"+str(characters))

                if checkIfAlreadyNamed:
                    are_parts_characternames=True
                    for k in characters:
                        if k in all_characters:
                            are_parts_characternames=True
                        else:
                            are_parts_characternames=False
                            break
                    if are_parts_characternames:
                        firstchar=characters[0].strip()
                        #print("REPLACE "+character+" with "+str(firstchar))
                        replaceList[character]=firstchar
                        item['character']=firstchar                   
                else:
                    firstchar=characters[0].strip()
                    #print("REPLACE"+character+" with "+str(firstchar))
                    replaceList[character]=firstchar
                    item['character']=firstchar                   

    return breakdown,replaceList
def filter_text(s):
    res=s.replace("♪","").replace("Â§","").replace("§","")
    #filter songs
    return res

def is_character_name_valid(char):
    isNote= "NOTE D'AUTEUR" in char
    isEnd = ("END CREDITS" in char )
    isNar=("NARRATIVE TITLE" in char)
    isOST =("ON-SCREEN TEXT" in char )
    isMain =("MAIN TITLE" in char )
    isOpen=("OPENING CREDITS" in char)
    return (not isNote) and (not isEnd) and (not isNar) and (not isOST) and (not isOpen) and (not isMain)

#################################################################
# PROCESS
def process_script(script_path,output_path,script_name,countingMethod,encoding,forceMode="",forceCols={}):
    print("  > -----------------------------------")
    print("  > SCRIPT PARSER version 1.3")
    print("  > Script path       : "+script_path)
    print("  > Output folder     : "+output_path)
    print("  > Script name       : "+script_name) 
    print("  > Forced encoding   : "+encoding)
    print("  > Counting method   : "+countingMethod)

    if not os.path.exists(output_path):
        os.mkdir(output_path)

    file_name = os.path.basename(script_path)
    name, extension = os.path.splitext(file_name)
    print("  > File name         : "+file_name)
    print("  > Extension         : "+extension)
    if not is_supported_extension(extension):
        print("  > File type "+extension+" not supported.")
        return

    if extension.lower()==".docx":

        print(" !!!!!!!!!! CHECK FILENAME")
        if "CLEAR CUT" in file_name:
            print(" !!!!!!!!!! FORCE")
            forceMode="DETECT_CHARACTER_DIALOG"
            forceCols={
                "CHARACTER":5,
                "DIALOG":6
            }
        converted_file_path=convert_word_to_txt(script_path,forceMode=forceMode,forceCols=forceCols)
        if len(converted_file_path)==0:
            print("  > Conversion failed 1")
            return None,None,None,None,None,None
        return process_script(converted_file_path,output_path,script_name,countingMethod,forceMode=forceMode,forceCols=forceCols)

    if extension.lower()==".pdf":
        converted_file_path,encoding=convert_pdf_to_txt(script_path)
        print("process 1")
        if len(converted_file_path)==0:
            print("  > Conversion failed 1")
            return None,None,None,None,None,None
        print("process")
        return process_script(converted_file_path,output_path,script_name,countingMethod,forceMode=forceMode,forceCols=forceCols)

    uppercase_lines=[]
    current_scene_id=""
    wasEmptyLine=False
    scene_characters_map={}
    character_linecount_map={}
    character_order_map={}
    character_order={}
    character_count=1
    character_textlength_map={}
    character_scene_map={}
    current_scene_count=1
    breakdown=[]

    is_verbose=False
    encoding_info = detect_file_encoding(script_path)
    print("  > Encoding          : "+encoding_info['encoding'])
    encoding_used=encoding_info['encoding']
    print("  > Encoding used     : "+encoding_used)

    encoding_tested=test_encoding(script_path)
    encoding_used=encoding_tested
    if not (encoding==""):
        print("  > Force encoding    :"+encoding)
        encoding_used=encoding
    scene_separator=getSceneSeparator(script_path,encoding_used)
    print("  > Scene separator   : "+scene_separator)

    character_mode=getCharacterSeparator(script_path,encoding_used)
    print("  > Character mode    : "+str(character_mode))
    
    if scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
        current_scene_id="Scene 1"
    # Open the file and process each line
    line_idx=1
    isEmptyLine=False
    with open(script_path, 'r', encoding=encoding_used) as file:
        for line in file:
            line = line.strip()  # Remove any leading/trailing whitespace
            trimmed_line = line.strip()  # Remove any leading/trailing whitespace
    
            isNewEmptyLine=len(trimmed_line)==0
            if scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
                if (not isNewEmptyLine) and  (isEmptyLine and wasEmptyLine):
                    current_scene_count=current_scene_count+1
                    current_scene_id = extract_scene_name(line,scene_separator,current_scene_count)
                    if is_verbose:
                        print("  > ---------------------------------------")
                    #print(f"Scene Line: {line}")
    
    
    
            isEmptyLine=len(trimmed_line)==0
            if is_verbose:
                print("  > Line "+str(line_idx))
            if len(trimmed_line)>0:
                if is_scene_line(line) or (isEmptyLine and wasEmptyLine):
                    current_scene_count=current_scene_count+1
                                    
                    current_scene_id = extract_scene_name(line,scene_separator,current_scene_count)
                    breakdown.append({"line_idx":line_idx,"scene_id":current_scene_id,"type":"SCENE_SEP" })    
                    if is_verbose:
                        print("  > --------------------------------------")
                    print(f"  > Scene Line: {current_scene_id}")
                else:
                        if True:#current_scene_id!=1:
                            is_speaking=is_character_speaking(trimmed_line,character_mode)
                            if is_verbose:
                                print("    IsSpeaking "+str(is_speaking)+" "+trimmed_line)
                            if is_speaking:
                                character_name=extract_character_name(trimmed_line,character_mode)
                                character_name=filter_character_name(character_name)
                                if is_verbose:
                                    print("   name="+str(character_name))
                                if not character_name == None:
                                    if is_character_name_valid(character_name):
                                        #remove character name for stats
                                        spoken_text=extract_speech(trimmed_line,character_mode,character_name)
                                        spoken_text=filter_text(spoken_text)
                                        
                                    
                                        breakdown.append({"scene_id":current_scene_id,
                                                        "character_raw":character_name,
                                                        "line_idx":line_idx,"speech":spoken_text,"type":"SPEECH", "character":character_name })    
                                        if is_verbose:
                                            print("   text="+str(spoken_text))
                                    
                            else:
                                breakdown.append({"line_idx":line_idx,"text":trimmed_line,"type":"NONSPEECH" })    
                                    
                                        
            wasEmptyLine=isEmptyLine
            line_idx=line_idx+1

    all_characters=get_all_characters(breakdown)
    breakdown,replaceList=merge_breakdown_character_talking_to(breakdown,all_characters)
    all_characters=get_all_characters(breakdown)
    #print("all_characters"+str(all_characters))
    #print("replacelist"+str(replaceList))

    replace_map=map_semi_duplicates(all_characters)
   # print("replace_map"+str(replace_map))

    breakdown=merge_breakdown_character_by_replacelist(breakdown,replace_map)
    all_characters=get_all_characters(breakdown)
 #   print("all_characters"+str(all_characters))
  #  print("replacelist"+str(replaceList))

    for item in breakdown:
        if item['type']=="SPEECH":
            character_name=item['character']
            #add to character order if new character
            if character_name not in character_order_map:
                character_order_map[character_name]=character_count
                character_count=character_count+1


            if character_name not in character_linecount_map:
                character_linecount_map[character_name]=1
            else:
                character_linecount_map[character_name]=character_linecount_map[character_name]+1

            spoken_text=item['speech']
            le=compute_length(spoken_text,countingMethod)
            if character_name not in character_textlength_map:
                character_textlength_map[character_name]=le
            else:
                character_textlength_map[character_name]=character_textlength_map[character_name]+le


            scene_id=item['scene_id']
            if character_name not in scene_characters_map:
                scene_characters_map[character_name] = set()
            scene_characters_map[character_name].add(scene_id)
            
            #add character to scene if not existing
            if scene_id not in character_scene_map:
                character_scene_map[scene_id] = set()
            character_scene_map[scene_id].add(character_name)
                                    

#    print(str(character_order_map))

    #character_scene_presence=sort_dict_values(character_scene_presence)
    #scene_characters_presence=sort_dict_values(scene_characters_presence)
    #write_character_map_to_file(character_scene_map, output_path+"character_by_scenes.txt")
    #write_character_map_to_file(scene_characters_map, output_path+"scenes_by_character.txt")
    #write_character_map_to_file(character_linecount_map, output_path+"character_linecount.txt")
    #write_character_map_to_file(character_order_map, output_path+"character_order.txt")
    #write_character_map_to_file(character_textlength_map, output_path+"character_textlength.txt")

    def save_string_to_file(text, filename):
        """Saves a given string `text` to a file named `filename`."""
        print(" > Write to "+filename)
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)

    #print(character_order_map)
#    s="Role,Lignes,Nb charactères,Répliques\n"
    #recap




    csv_file_path =output_path+script_name+"-recap-detailed.csv"
    data = [
    ]
    for key in character_order_map:
        data.append([
            str(character_order_map[key])+" - "+str(key),str(character_linecount_map[key]),
            str(character_textlength_map[key]),
            str(math.ceil(character_textlength_map[key]/40))])

    print("  > Convert to csv.")
    with open(csv_file_path, mode='w', newline='',encoding=encoding_used) as file:
        writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        
        # Write data to the CSV file
        for row in data:
            writer.writerow(row)


    #s=""
    #for key in character_order_map:
    #    s=s+str(character_order_map[key])+" - "+str(key)+","+str(character_linecount_map[key])+","+str((character_textlength_map[key]))+","+str(math.ceil(character_textlength_map[key]/40))+"\n"
    #save_string_to_file(s, output_path+script_name+"-recap.csv")
    print("  > Convert to xslx.")
    convert_csv_to_xlsx(output_path+script_name+"-recap-detailed.csv",output_path+script_name+"-recap-detailed.xlsx", script_name,encoding_used)
    print("  > Parsing done.")
    return breakdown, character_scene_map,scene_characters_map,character_linecount_map,character_order_map,character_textlength_map


#process_script("scripts/COMPTAGE/20012020.txt","20012020/","20012020","ALL")    
#process_script("scripts/examples/YOU CAN'T RUN FOREVER_SCRIPT_VO.txt","YOUCANRUNFOREVER_SCRIPT_VOe/","YOU CAN'T RUN FOREVER_SCRIPT_VO","ALL")    
#process_script("190421-1.txt","190421-1/","190421-1.txt")
#process_script("scripts/examples/LATENCY.docx","LATENCY/","LATENCY","ALL")    
#process_script("scripts/examples/LATENCY.docx","LATENCY/","LATENCY","ALL")    
#process_script("scripts/examples/Gods of the deep - CCSL.docx","GODS/","GODS","ALL")    
#process_script("scripts/examples/Blackwater Lane.docx","Blackwater Lane/","Blackwater Lane","ALL")    

            