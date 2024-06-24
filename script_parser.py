#from PyPDF2 import PdfReader
import re
import os
import math
import sys
import io
import pandas as pd
from docx import Document
import csv
import platform
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
from pdfminer.layout import LTPage
from utils import get_log_file_path
import pdfplumber
import logging
import chardet

#script_path="scripts2/YOU CAN'T RUN FOREVER_SCRIPT_VO.txt"
#output_path="YOU CANT RUN FOREVER_SCRIPT_VOc/"
#script_name="YOU CANT RUN FOREVER_SCRIPT_VO"

#script_path="scripts2/ZERO11.txt"
#output_path="ZERO11c/"
#script_name="ZERO11b"



app_log_path=get_log_file_path()
logging.basicConfig(filename=app_log_path,level=logging.DEBUG,encoding='utf-8',filemode='w')
logging.debug("Script starting...")

if False:
    class UTFStreamHandler(object):
        def __init__(self, stream):
            self.stream = stream

        def write(self, data):
            if not isinstance(data, str):
                data = str(data)
            self.stream.buffer.write(data.encode('utf-8'))
            self.stream.buffer.flush()

        def flush(self):
            self.stream.flush()

    # Only replace sys.stdout if it's not already wrapped
    if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout = UTFStreamHandler(sys.stdout)

#sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def myprint1(s):
    logging.debug(s)
    #myprint1(s)


action_verbs = ["says", "asks", "whispers", "shouts", "murmurs", "exclaims"]


characterSeparators=[
        "CHARACTER_SEMICOL_TAB",
        "CHARACTER_TAB",
        "CHARACTER_SPACES"
]
multilineCharacterSeparators=[
        "CHARACTER_NEWLINE_DIALOG_NEWLINE_NEWLINE"
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
    myprint1("detect encoding of file "+str(file_path))
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result

def test_encoding(script_path):
    encodings=['windows-1252', 'iso-8859-1', 'utf-16','ascii','utf-8']
    for enc in encodings:
        try:
            with open(script_path, 'r', encoding=enc) as file:
                myprint1("  > Testing encoding  : "+enc)

                for line in file:
                    line = line.strip()  # Remove any leading/trailing whitespace
            return enc
        except UnicodeDecodeError:
            myprint1(f"  > Failed decoding with {enc}")
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
    #myprint1("IsScene    "+str(isSceneLine)+" "+line)
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

def count_nonempty_lines_in_file(script_path,encod):
    nLines=0
    with open(script_path, 'r', encoding=encod) as file:
        for line in file:
            line = line.strip()
            if len(line)>0:
                nLines=nLines+1
    return nLines

def extract_matches(file_path,encod):
    # Define the pattern to match
    pattern = re.compile(r'^([A-Z\s]+)$\n((?:[^\n]+\n)*?)\n', re.MULTILINE)

    # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Extract the matched groups into a list of dictionaries
    extracted_data = []
    for match in matches:
        uppercase_text = match[0]
        lines_of_text = match[1].strip()  # Remove trailing newlines
        extracted_data.append({
            'character': uppercase_text,
            'dialog': lines_of_text
        })

    return extracted_data
def count_matches_charactername_NAME_NEWLINE_DIALOG_NEWLINE_NEWLINE(file_path,encod):
    # Define the pattern to match
    pattern = re.compile(r'^[A-Z\s]+$\n([^\n]+\n)*\n', re.MULTILINE)

    # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Return the number of matches
    return len(matches)
def getCharacterSepType(character_mode):
    if character_mode in characterSeparators:
        return "CHARACTER_MODE_SINGLELINE"
    if character_mode in multilineCharacterSeparators:
        return "CHARACTER_MODE_MULTILINE"
def detectCharacterSeparator(script_path,encod):
    myprint1("detectCharacterSeparator")
    best="?"
    bestVal=0.0
    nLines=count_nonempty_lines_in_file(script_path,encod)
    myprint1(f"detectCharacterSeparator nlines={nLines}")
    for sep in characterSeparators:
        nMatches=0
        with open(script_path, 'r', encoding=encod) as file:
            for line in file:
                line = line.strip()
                if len(line)>0:
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
        myprint1("  > Test character sep:"+sep+" " +str(nMatches)+"/"+str(nLines)+" "+str(pc))
        if pc>bestVal:
            bestVal=pc
            best=sep

        if best=="?":
            myprint1("  > Test character sep:"+sep+" " +str(nMatches)+"/"+str(nLines)+" "+str(pc))
            for sep in multilineCharacterSeparators:
                if sep == "CHARACTER_NEWLINE_DIALOG_NEWLINE_NEWLINE":
                    nMatches=count_matches_charactername_NAME_NEWLINE_DIALOG_NEWLINE_NEWLINE(script_path,encod)
                    
                if nMatches>0:
                    best=sep

    myprint1(f"detectCharacterSeparator best= {best}")
    return best
def getSceneSeparator(script_path,encod):
    mode="?"
    # Open the file and process each line
    
    with open(script_path, 'r', encoding=encod) as file:
        for line in file:
            line = line.strip()  # Remove any leading/trailing whitespace
            if matches_format_parenthesis_name_timecode(line):
                myprint1("PARENTHESIS_NAME_TIMECODE")
                return "PARENTHESIS_NAME_TIMECODE"
            
            elif matches_number_parenthesis_timecode(line):
                myprint1("NAME_PARENTHESIS_TIMECODE")
                return "NAME_PARENTHESIS_TIMECODE"

    if mode=="?":
        n_sets_of_empty_lines=count_consecutive_empty_lines(script_path,2,encod)
        #myprint1("check empty lines count"+str(n_sets_of_empty_lines))
        if n_sets_of_empty_lines>1:
            myprint1( "Found EMPTYLINES_SCENE_SEPARATOR in "+line)
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
        myprint1("ERROR wrong mode="+character_mode)


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
        myprint1("ERROR wrong mode="+str(character_mode))
        exit()

def extract_character_name(line,character_mode):
    if character_mode=="CHARACTER_TAB":
        return extract_charactername_NAME_ATLEAST1TAB_TEXT(line)
    elif character_mode=="CHARACTER_SPACES": 
        myprint1("debug")       
        is_match=matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)
        myprint1("is_match?"+str(is_match))
        myprint1("extract "+str(character_mode)+" "+str(line))
        return extract_charactername_NAME_ATLEAST8SPACES_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return extract_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)  
    else:
        myprint1("ERROR wrong mode="+str(character_mode))
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
    #myprint1("test match"+str(text))
    # Define the regex pattern:
    # ^ starts the match at the beginning of the line
    # (.+) matches one or more of any character (the first text block), captured for potential use
    # {8,} specifies at least 8 spaces
    # (.+) matches one or more of any character following the spaces (the second text block)
    #pattern = r'^(.+)\s{8,}(.+)$'
    pattern = r'^(.+?)\s{8,}(.+)$'
    # Use re.match to check if the whole string matches the pattern
    if re.match(pattern, text):
       # myprint1("is match")
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
    myprint1("extract name"+str(line))
    pattern = r'^(.+?)\s{8,}(.+)$'
    match = re.match(pattern, line)
    
#    match = re.match(r"^([A-Z ]+)\s{8,}.*$", line)
    if match:
        myprint1("is match")
        return match.group(1).strip()
    myprint1("no match")
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
        
        if "(V.O)" in line:
            line=line.replace("(V.O)","")
        if "(V.O.)" in line:
            line=line.replace("(V.O.)","")
        if "(V.O" in line:
            line=line.replace("(V.O","")
        
        if "(O.S.)" in line:
            line=line.replace("(O.S.)","")
        
        if "(CON'T)" in line:
            line=line.replace("(CON'T)","")    
        if "(CONT.)" in line:
            line=line.replace("(CONT.)","")    
        if "(CONT." in line:
            line=line.replace("(CONT.","")    
        if "(CON’T)" in line:
            line=line.replace("(CON’T)","")    
        if "(CON’T)" in line:
            line=line.replace("(CON’T)","")    
        if "(CONT'D)" in line:
            line=line.replace("(CONT'D)","")    
        if "(CONT.)" in line:
            line=line.replace("(CONT')","")    
        if "(CONT'D" in line:
            line=line.replace("(CONT'D","")    
        if "(CONT’D" in line:
            line=line.replace("(CONT’D","")    
        line=line.replace("\ufeff","")
        if line.endswith(':'):
            line= line[:-1]
        if line.endswith(')'):
            line= line[:-1]
    return line.strip()
#    return line



#################################################################
# UTILS


def convert_csv_to_xlsx(csv_file_path, xlsx_file_path, script_name,encoding_used):
    myprint1("convert_csv_to_xlsx > 0")
    # Read the CSV file
    df = pd.read_csv(csv_file_path,header=None,encoding=encoding_used)

    # Write the DataFrame to an Excel file
    #myprint1("convert_csv_to_xlsx > Write to "+xlsx_file_path)

    header_rows = pd.DataFrame([
        [None, 'Header 1', None, 'Header Information Across Columns'],  # Merge cells will be across 1 & 4
        ['Role', 'Prises de parole', 'Caractères', 'Lignes']
    ])
    #myprint1("convert_csv_to_xlsx > 1")
    
    # Concatenate the header rows and the original data
    # The ignore_index=True option reindexes the new DataFrame
    df = pd.concat([header_rows, df], ignore_index=True)
    #myprint1("convert_csv_to_xlsx > 2")

    # Write the DataFrame to an Excel file
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        #myprint1("convert_csv_to_xlsx > 3")

        df.to_excel(writer, index=False, sheet_name='Sheet1')
        #myprint1("convert_csv_to_xlsx > 4")

        # Load the workbook and sheet for modification
        workbook = writer.book
        sheet = workbook['Sheet1']
        #myprint1("convert_csv_to_xlsx > 5")

        # Merge cells in the first and second new rows
        # Assuming you want to merge from the first to the last column
        sheet.merge_cells('A1:D1')  # Modify range according to your number of columns
        sheet.merge_cells('A2:D2')  # Modify this as needed
        sheet['A1'] = script_name
        sheet['A2'] = "Length: "
        #myprint1("convert_csv_to_xlsx > 6")

    myprint1("convert_csv_to_xlsx > done")

#    df.to_excel(xlsx_file_path, index=False, engine='openpyxl')

def write_character_map_to_file(character_map, filename):
    myprint1(" > Write map to "+filename)
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
                #myprint1(str(i)+"empty line")
            else:
                #myprint1(str(i)+"not empty pre="+str(previous_empty)+" co="+str(count_empty))
                if previous_empty and count_empty >= n:
                    occurrences += 1
                    #myprint1(str(i)+" ADD OCC")

                count_empty = 0
                previous_empty = False
            i=i+1
        # Check at the end of the file if the last lines were empty
        if count_empty >= n:
            occurrences += 1
            #myprint1(str(i)+" ADD OCC")

    return occurrences


def sort_dict_values(d):
    sorted_dict = {}
    for key, value_set in d.items():
        try:
            # Attempt to sort assuming all values are numeric strings
            sorted_list = sorted(value_set, key=int)
        except ValueError:
            # Handle the case where values are not all numeric
            myprint1(f"Non-numeric values found in the set for key '{key}'. Values: {value_set}")
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
        myprint1(para.text)
    return doc

def is_supported_extension(ext):
    ext=ext.lower()
    return ext==".txt" or ext==".docx" or ext==".doc" or ext==".rtf" or ext==".pdf" or ext==".xlsx"



def convert_docx_combined_continuity(file,table):
    myprint1("Conversion mode       : COMBINED_CONTINUITY")
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
  #                  myprint1("split mode speech"+title+str(len(speeches)))

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
                    #myprint1("linear mode add "+s)
                        
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
    s=s.replace("♪","").replace("Â§","").replace("§","")
    s=s.replace("â€¦ ",".")
    if s.startswith("- "):
        s=s.lstrip("- ")
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
                    #    myprint1("part "+str(part_idx)+":"+part)
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
    myprint1("Conversion mode       : SCENEDESCRIPTION")
    current_character=""
    idx=1
    cumulated_speech=""
    row_idx=1

    for row in table.rows[1:]:
        if row_idx<35:
                myprint1("---------------------------------------------")
                myprint1("Row "+str(row_idx))
                scenedesc=row.cells[titlesIdx].text.strip()        
                myprint1("Row content"+str(scenedesc))
        row_idx=row_idx+1
    row_idx=1
    for row in table.rows[1:]:
            if idx<10:
                myprint1("---------------------------------------------")
                myprint1("Row "+str(row_idx))
            row_idx=row_idx+1
            scenedesc=row.cells[titlesIdx].text.strip()        
            if idx<100:
                myprint1(scenedesc)
            parts=scenedesc.split("\n")
            part_idx=1
            for part in parts:
                if idx<10:
                    myprint1("part "+str(part_idx)+":"+part)
                
                part=part.strip()
                part_idx=part_idx+1
                if len(part)>0:
#                    if idx<10:
 #                       myprint1("part "+str(part_idx)+":"+part)
                    part=filter_speech(part)
                    if match_uppercase_semicolon(part):
                        #flush cunulated speech
                        if len(cumulated_speech)>0:
                            s=current_character+"\t"+cumulated_speech+"\n"
                            file.write(s)
                            if idx<10:
                                myprint1(">>>>>>>>> "+current_character+"\t"+cumulated_speech)
                            idx=idx+1
                            cumulated_speech=""
                        current_character=remove_semicolon(part) 
                        if idx<10:
                            myprint1("new character "+current_character)
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
        new_file_abs = os.path.join(currentOutputFolder, base)

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
    myprint1("convert_docx_plain_text")
    myprint1("convert_docx_plain_text "+str(len(doc.paragraphs))+" paragraphs")
    for para in doc.paragraphs:
        # Write the text of each paragraph to the file followed by a newline
        myprint1("convert_docx_plain_text write "+para.text)
        myprint1("                        isupper="+str(para.text.isupper()))
        myprint1("                        para= "+str(para))
        file.write(para.text + '\n')

def convert_docx_indented_plaintext(file,doc):
    myprint1("convert_docx_indented_plaintext")
    myprint1("convert_docx_indented_plaintext "+str(len(doc.paragraphs))+" paragraphs")
    for para in doc.paragraphs:
        if len(para.text)>0:
            indent = para.paragraph_format.left_indent
            myprint1("convert_docx_indented_plaintext text="+para.text)
            myprint1("convert_docx_indented_plaintext indent="+str(indent))
            
            if indent is not None:
                # Write the text of each paragraph to the file followed by a newline
                myprint1("                        write "+str(para.text.isupper()))
                myprint1("                        write "+str(para))
                file.write(para.text + '\n')

def extract_speakersNO(conversation, character_list):
    # Helper function to validate if a name is in the character list
    def is_valid_character(name):
        return name in character_list

    # Split the string by " TO " and process parts
    parts = conversation.split(" TO ")
    names = []
    seen = set()
    
    i = 0
    while i < len(parts) - 1:
        speaker = parts[i]
        listener = parts[i + 1]
        
        # Add speaker if it's a valid character and not already seen
        if is_valid_character(speaker) and speaker not in seen:
            names.append(speaker)
            seen.add(speaker)
        
        # Add listener if it's a valid character and not already seen
        if is_valid_character(listener) and listener not in seen:
            names.append(listener)
            seen.add(listener)
        
        i += 2
    
    # Handle the case where conversation does not end in a proper " TO " sequence
    if len(parts) % 2 == 1:
        last_part = parts[-1]
        if is_valid_character(last_part) and last_part not in seen:
            names.append(last_part)
    
    # Handle case where there's only one speaker or no valid conversations
    if not names and parts:
        if is_valid_character(parts[0]):
            return [parts[0]]
    
    return names
def normalize_spaces(text):
    return re.sub(r'\s+', ' ', text).strip()
def extract_speakers(conversation,character_list):
    # Split the string by " TO "
    parts = conversation.split(" TO ")
    nb_counts=conversation.count(" TO ")
    if nb_counts==0:
        return [conversation]
    if nb_counts==1:
        return [parts[0]]
    # Check if there are exactly two parts
    if nb_counts == 2:
        # If there is a reciprocal conversation, both parts should have the same speaker at the end
        first_speaker=parts[0]
        last_dest=parts[2]
        mid=parts[1]
        mid=normalize_spaces(mid)
        myprint1("case first= "+str(first_speaker)+" mid="+str(mid)+" last "+str(last_dest) )

        nb_words=mid.count(" ")
        midspl=mid.split(" ")
        myprint1("case mid= "+str(midspl) )
        if len(midspl)==2:
            second_speaker=midspl[1]
            return [first_speaker,second_speaker]
        else:
            myprint1("error splitting characters within "+str(conversation) )
    else:
        # For more complex conversations, extract unique names
        seen = set()
        names = []
        for name in parts:
            if name not in seen:
                names.append(name)
                seen.add(name)
        return names




    
def extract_speakers1(conversation):
    # Split the string by " TO "
    parts = conversation.split(" TO ")

    # Extract names while preserving order
    seen = set()
    names = []
    for name in parts:
        if name not in seen:
            names.append(name)
            seen.add(name)
    
    # If there is only one unique name, return just that name
    if len(names) == 1:
        return names[0]
    
    # Otherwise, return the list of unique names
    return names

def convert_docx_characterid_and_dialogue(file,table,dialogueCol,characterIdCol):
    myprint1("Conversion mode       : CHARACTERID_AND_DIALOGUE")
    # Iterate through each row in the table
    current_character=""
    is_song_sung_by_character=False
    character_list=set()
    for row in table.rows[1:]:
        cell_texts = [cell.text for cell in row.cells]
            # Join all cell text into a single string
        row_text = ' | '.join(cell_texts)
        
        
        dialogue=row.cells[dialogueCol].text.strip()
        character=row.cells[characterIdCol].text.strip()

        mode="LINEAR"
        dialogue=dialogue.replace("\n","")
        myprint1("----------")
        myprint1("row"+row_text)
        myprint1("dialogue"+dialogue)
        myprint1("character"+character)
        character=filter_character_name(character)
        if not character in character_list:
            character_list.add(character)
        myprint1("character filtered"+character)
        ##if "(O.S)" in character:
          #  character=character.replace("(O.S)","")
        #if "(O.S.)" in character:
         #   character=character.replace("(O.S.)","")

        myprint1("extract speakers for"+character)
        speakers=extract_speakers(character,character_list)

        #spl=character.split("\n")
        myprint1("speakers="+str(speakers))
        nb_lines= len(speakers)
        myprint1("Mode "+mode)
        if nb_lines>1:
            mode="SPLIT"

        

        if mode=="LINEAR" :   
        
            if len(character)>0:
                current_character=character
            if len(dialogue)>0:  
                is_didascalie=dialogue.startswith("(")
                is_song=dialogue.startswith("♪") 
                is_song_sung_by_character=is_song and len(character)>0 #song starts but has character name on it
                if is_song:
                    dialogue=filter_speech(dialogue)             
                    force_current_character="__SONG"
                    if is_song_sung_by_character:
                         force_current_character=current_character
                    s=force_current_character+"\t"+dialogue+"\n"  # New line after each row
                    myprint1("Add "+ force_current_character + " "+ dialogue)
                    file.write(s)
                else:
                    is_song_sung_by_character=False
                if not is_didascalie and not is_song: 
                    dialogue=filter_speech(dialogue)             
                    if current_character=="":
                        current_character="__VOICEOVER"
                    s=current_character+"\t"+dialogue+"\n"  # New line after each row
                    myprint1("Add "+ current_character + " "+ dialogue)
                    file.write(s)
        elif mode=="SPLIT":
            myprint1("MODE SPLIT")
            dialogue=filter_speech(dialogue)    
            myprint1("characters"+str(speakers))
            dialoguespl=dialogue.split("- ")
            filtered_array = [element for element in dialoguespl if element]
            myprint1("dialogue"+str(filtered_array))
            for index,k in enumerate(speakers):
                myprint1("index"+str(index)+" k="+str(k))

                current_character=k
                speech=filtered_array[index]
                s=current_character+"\t"+speech+"\n"  # New line after each row
                myprint1("Add "+ current_character + " "+ dialogue)
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
        new_file_abs = os.path.join(currentOutputFolder, base)

        # Save and Close
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=16)  # FileFormat=16 for .docx, not .doc
        doc.Close(False)

        return new_file_abs

    finally:
        # Make sure to uninitialize COM
        pythoncom.CoUninitialize()
    return ""


def is_in_left_margin(block_left,  threshold):
    return block_left<threshold
def is_in_top_margin(block_left,  threshold):
    return block_left>threshold
def is_in_bottom_margin(block_left,  threshold):
    return block_left<threshold
def is_in_right_margin(block_left, min_left):
    return block_left>min_left

def is_centered(block_left, min_left, max_left, threshold):
    center = (min_left + max_left) / 2
    return abs(block_left - center) <= threshold
       
def get_pdf_text_elements(file_path,page_idx, page_start,page_end,progress_bar):
    return get_pdf_text_elements_alt(file_path,page_idx,page_start,page_end,progress_bar)

def get_pdf_text_elements_alt(file_path,page_idx, page_start,page_end,progress_bar):
        myprint1("---------- get_pdf_page_blocks -----------------")
        text_elements=[]
        minboxleft=100000
        if page_idx<0:
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    progress_bar['value'] = page_num
                    progress_bar.update_idletasks()

                    myprint1(f"\nPage {page_num}")
                    if page_num<page_start:
                            continue
                    if page_num>page_end:
                            break
                                    # Get page dimensions
                    page_width = page.width
                    page_height = page.height
                    page_elements=[]
                    # Text
                    text_settings = {
                    "x_tolerance": 3,  # Adjust this value to fine-tune line detection
                    "y_tolerance": 3,  # Adjust this value to fine-tune line detection
                    }
                    page_layout = page.extract_text_lines(text_settings)
                    for element in page_layout:
                                myprint1(f"  Word: '{element['text']}', bbox: {element['x0']}, {element['top']}, {element['x1']}, {element['bottom']}")
                                tt=element['text'].strip().replace("\n","")
                                bbox =   [element['x0'],page_height -  element['top'], element['x1'], page_height - element['bottom']]                  
                                stripped_text=tt
                                if len(stripped_text)>0:
                                    # Print the text block and its bounding box
                                    myprint1(f"{bbox} {tt}")
                                    if bbox[0] < minboxleft:
                                        minboxleft=bbox[0]
                                    page_elements.append({
                                        'text':stripped_text,
                                        'bbox':bbox,
                                    })
                    myprint1("NON EMPTY ELEMENTS: "+str(len(page_elements)))
                
                    page_elements = sorted(page_elements, key=lambda x: x['bbox'][1], reverse=True)
                    for k in page_elements:
                        text_elements.append(k)
        else:        
            with pdfplumber.open(file_path) as pdf:
                page = pdf.pages[page_idx]
                page_layout=page.extract_words()
                
                # Get page dimensions
                page_width = page.width
                page_height = page.height
                text_settings = {
                    "x_tolerance": 3,  # Adjust this value to fine-tune line detection
                    "y_tolerance": 3,  # Adjust this value to fine-tune line detection
                }
                page_layout = page.extract_text_lines(text_settings)

                myprint1("ELEMENTS: "+str(len(page_layout)))
                for element in page_layout:
                        
                        # Loop over the text blocks within the text container
                        text_block = element['text']
                        bbox =   [element['x0'], page_height - element['top'], element['x1'],page_height -  element['bottom']]                  
                        
                        tt=text_block.strip().replace("\n","")
                        stripped_text=tt
                        if len(stripped_text)>0:
                            # Print the text block and its bounding box
                            myprint1(f"{bbox} {stripped_text}")
                            if bbox[0] < minboxleft:
                                minboxleft=bbox[0]
                            text_elements.append({
                                'text':stripped_text,
                                'bbox':bbox,
                            })
                myprint1("NON EMPTY ELEMENTS: "+str(len(text_elements)))
                
                text_elements = sorted(text_elements, key=lambda x: x['bbox'][1], reverse=True)

    
        return text_elements
def get_pdf_text_elements_pdfminer(file_path,page_idx, page_start,page_end):
        myprint1("---------- get_pdf_page_blocks -----------------")
        text_elements=[]
        minboxleft=100000

        if page_idx<0:
            for page_layout in extract_pages(file_path):
                    myprint1("----------------------------")
                    myprint1(f"Page number: {page_layout.pageid}")
                    if page_layout.pageid<page_start:
                        continue
                    if page_layout.pageid>page_end:
                        break

                    # Get the page bounding box coordinates and dimensions
                    if isinstance(page_layout, LTPage):
                        page_bbox = page_layout.bbox
                        page_x0, page_y0, page_x1, page_y1 = page_bbox
#                        page_width = page_x1 - page_x0
 #                       page_height = page_y1 - page_y0
                    else:
                        # If page bounding box is not available, skip processing the page
                        continue
                    # Loop over the elements in the page
                    myprint1(str(page_layout))
                    myprint1("ELEMENTS: "+str(len(page_layout)))
                    page_elements=[]
                    for element in page_layout:
                        
                        # Check if the element is a text container
                        if isinstance(element, LTTextContainer):
                            # Loop over the text blocks within the text container
                            text_block = element.get_text()
                            bbox = element.bbox  # (x0, y0, x1, y1)                        
                            is_inside=    bbox[0] >= page_x0 and bbox[1] >= page_y0 and                bbox[2] <= page_x1 and bbox[3] <= page_y1
                            if is_inside:
                                tt=text_block.strip().replace("\n","")
                                stripped_text=tt
                                if len(stripped_text)>0:
                                # Print the text block and its bounding box
                                    myprint1(f"{bbox} {tt}")
                                    if bbox[0] < minboxleft:
                                        minboxleft=bbox[0]
                                    page_elements.append({
                                        'text':text_block.strip(),
                                        'bbox':bbox,
                                    })
                    myprint1("NON EMPTY ELEMENTS: "+str(len(page_elements)))
                
                    page_elements = sorted(page_elements, key=lambda x: x['bbox'][1], reverse=True)
                    for k in page_elements:
                        text_elements.append(k)
        else:        
            for page_layout in extract_pages(file_path, page_numbers=[page_idx ]):
                myprint1("----------------------------")
                myprint1(f"Page number: {page_layout.pageid}")
                myprint1("-- PAGE "+str(page_idx))
            
                
                # Get the page bounding box coordinates and dimensions
                if isinstance(page_layout, LTPage):
                    page_bbox = page_layout.bbox
                    page_x0, page_y0, page_x1, page_y1 = page_bbox
                    page_width = page_x1 - page_x0
                    page_height = page_y1 - page_y0
                else:
                    # If page bounding box is not available, skip processing the page
                    continue
                # Loop over the elements in the page
                myprint1(str(page_layout))
                myprint1("ELEMENTS: "+str(len(page_layout)))
                for element in page_layout:
                    
                    # Check if the element is a text container
                    if isinstance(element, LTTextContainer):
                        # Loop over the text blocks within the text container
                        text_block = element.get_text()
                        bbox = element.bbox  # (x0, y0, x1, y1)                        
                        is_inside=    bbox[0] >= page_x0 and bbox[1] >= page_y0 and                bbox[2] <= page_x1 and bbox[3] <= page_y1
                        if is_inside:
                            tt=text_block.strip().replace("\n","")
                            stripped_text=tt
                            if len(stripped_text)>0:
                                # Print the text block and its bounding box
                                myprint1(f"{bbox} {stripped_text}")
                                if bbox[0] < minboxleft:
                                    minboxleft=bbox[0]
                                text_elements.append({
                                    'text':stripped_text,
                                    'bbox':bbox,
                                })
                myprint1("NON EMPTY ELEMENTS: "+str(len(text_elements)))
                
                text_elements = sorted(text_elements, key=lambda x: x['bbox'][1], reverse=True)

    
        return text_elements
def get_pdf_page_blocks(file_path,page_idx):
    myprint1("---------- get_pdf_page_blocks -----------------")
    text_elements=[]
    minboxleft=100000
    for page_layout in extract_pages(file_path, page_numbers=[page_idx ]):
                myprint1("----------------------------")
                myprint1(f"Page number: {page_layout.pageid}")
                myprint1("-- PAGE "+str(page_idx))
            
                # Get the page bounding box coordinates and dimensions
                if isinstance(page_layout, LTPage):
                    page_bbox = page_layout.bbox
                    page_x0, page_y0, page_x1, page_y1 = page_bbox
                    page_width = page_x1 - page_x0
                    page_height = page_y1 - page_y0
                else:
                    # If page bounding box is not available, skip processing the page
                    continue
                # Loop over the elements in the page
                myprint1(str(page_layout))

                for element in page_layout:
                    
                    # Check if the element is a text container
                    if isinstance(element, LTTextContainer):
                        # Loop over the text blocks within the text container
                        text_block = element.get_text()
                        #for text_line in element:
                        #   text_block = ""
                        #  text_position = []

                        bbox = element.bbox  # (x0, y0, x1, y1)
                        
                        is_inside=    bbox[0] >= page_x0 and bbox[1] >= page_y0 and                bbox[2] <= page_x1 and bbox[3] <= page_y1
                        #myprint1(                bbox[0] >= page_x0) 
                        #myprint1(         bbox[1] >= page_y0 )
                        #myprint1(         bbox[2] <= page_x1 )
                        #myprint1(          bbox[3] <= page_y1)
         
                        if is_inside:
                            # Print the text block and its bounding box
                            myprint1(f"{bbox} {text_block.strip()}")
                            if bbox[0] < minboxleft:
                                minboxleft=bbox[0]
                            text_elements.append({
                                'text':text_block.strip(),
                                'bbox':bbox,
                            })
    xvalues={}
    for k,el in enumerate(text_elements):
        left=el['bbox'][0]
        if left in xvalues:
            xvalues[left]=xvalues[left]+1
        else:
            xvalues[left]=1

    myprint1("XVALUES"+str(xvalues))

    xval=list(xvalues.keys())
    # Find the minimum and maximum left positions
    min_left = min(xval)
    max_left = max(xval)
    myprint1("MIN_LEFT"+str(min_left))
    myprint1("MAX_LEFT"+str(max_left))


    left_margin=110
    top_margin=740
    right_margin=500
    # Categorize blocks as left-aligned or centered
    top_aligned_blocks = []
    left_aligned_blocks = []
    right_aligned_blocks = []
    centered_blocks = []

    myprint1("-------------------------------")
    myprint1("TEST")
    for k,el in enumerate(text_elements):
        left_pos=el['bbox'][0]
        bottom_pos=el['bbox'][1]
        myprint1("test ["+str(left_pos) +"]"+str(el['text']))
        if is_in_top_margin(bottom_pos, top_margin):
            myprint1("test left")
            top_aligned_blocks.append(el)
        elif is_in_left_margin(left_pos, left_margin):
            myprint1("test left")
            left_aligned_blocks.append(el)
        elif is_in_right_margin(left_pos,right_margin):
            myprint1("test right")
            right_aligned_blocks.append(el)
        else:#if is_centered(left_pos, min_left, max_left, threshold):
            myprint1("test centre")
            centered_blocks.append(el)
    return {
        'left':left_aligned_blocks,
        'right':right_aligned_blocks,
        'top':top_aligned_blocks,
        'center':centered_blocks
    }
def split_elements(text_elements,left_margin,top_margin,right_margin,bottom_margin):
    top_blocks = []
    left_blocks = []
    right_blocks = []
    bottom_blocks = []
    centered_blocks = []
    myprint1("SPLIT")
    if text_elements != None:
        myprint1("NB EL"+str(len(text_elements)))
        myprint1(f" > elements {len(text_elements)}")
    myprint1(f" > margins left={left_margin} top={top_margin} right={right_margin} bottom={bottom_margin}")
    #myprint1("TEST")
    if text_elements!=None:
        for k,el in enumerate(text_elements):
            if len(el['text'])>0:
                left_pos=el['bbox'][0]
                bottom_pos=el['bbox'][1]
            #   myprint1("test ["+str(left_pos) +"]"+str(el['text']))
                if is_in_top_margin(bottom_pos, top_margin):
                    #myprint1(f"    - is_top bottompos={bottom_pos} topmargin={top_margin}")
                    top_blocks.append(el)
                elif is_in_left_margin(left_pos, left_margin):
            #     myprint1("test left")
                    left_blocks.append(el)
                elif is_in_right_margin(left_pos,right_margin):
                #    myprint1("test right")
                    right_blocks.append(el)
                elif is_in_bottom_margin(bottom_pos,bottom_margin):
                #    myprint1("test right")
                    bottom_blocks.append(el)
                else:#if is_centered(left_pos, min_left, max_left, threshold):
                #   myprint1("test centre")
                    centered_blocks.append(el)
        myprint1(f" > groups left={len(left_blocks)} top={len(top_blocks)} right={len(right_blocks)} center={len(centered_blocks)} bottom={len(bottom_blocks)}")

    return {
        'bottom':bottom_blocks,
        'left':left_blocks,
        'right':right_blocks,
        'top':top_blocks,
        'center':centered_blocks
    }

def convert_pdf_to_txt(file_path,absCurrentOutputFolder,encoding):
    myprint1("convert_pdf_to_txt")
    myprint1("currentOutputFolder             :"+absCurrentOutputFolder)
    myprint1("Input              :"+file_path)
    converted_file_path=""
    if ".pdf" in file_path.lower() :
        converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".converted.txt"))
    if ".pdf" in file_path.lower() :
        converted_raw_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".raw.txt"))
    minboxleft=100000
    npagesmax=10000
    firstpage=3
    page_idx=0
    
    with open(converted_raw_file_path, 'w', encoding='utf-8') as fileraw:              
        
        text_elements=[]
        for page_layout in extract_pages(file_path):
                page_idx=page_idx+1
                if page_idx<firstpage:
                    continue
                if page_idx>npagesmax:
                    break
                myprint1("----------------------------")
                myprint1(f"Page number: {page_layout.pageid}")
                myprint1("-- PAGE "+str(page_idx))
            
                # Get the page bounding box coordinates and dimensions
                if isinstance(page_layout, LTPage):
                    page_bbox = page_layout.bbox
                    page_x0, page_y0, page_x1, page_y1 = page_bbox
                    page_width = page_x1 - page_x0
                    page_height = page_y1 - page_y0
                else:
                    # If page bounding box is not available, skip processing the page
                    continue
                # Loop over the elements in the page
                for element in page_layout:
                    # Check if the element is a text container
                    if isinstance(element, LTTextContainer):
                        # Loop over the text blocks within the text container
                        text_block = element.get_text()
                        #for text_line in element:
                        #   text_block = ""
                        #  text_position = []

                        bbox = element.bbox  # (x0, y0, x1, y1)
                        
                        is_inside=    bbox[0] >= page_x0 and bbox[1] >= page_y0 and                bbox[2] <= page_x1 and bbox[3] <= page_y1
                        #myprint1(                bbox[0] >= page_x0) 
                        #myprint1(         bbox[1] >= page_y0 )
                        #myprint1(         bbox[2] <= page_x1 )
                        #myprint1(          bbox[3] <= page_y1)
         
                        if is_inside:
                            # Print the text block and its bounding box
                            if len(text_block.strip())>0:
                                myprint1(f"{bbox} {text_block.strip()}")
                            if bbox[0] < minboxleft:
                                minboxleft=bbox[0]
                            text_elements.append({
                                'text':text_block.strip(),
                                'bbox':bbox,
                            })
    xvalues={}
    for k,el in enumerate(text_elements):
        left=el['bbox'][0]
        if left in xvalues:
            xvalues[left]=xvalues[left]+1
        else:
            xvalues[left]=1

    myprint1("XVALUES"+str(xvalues))

    xval=list(xvalues.keys())
    # Find the minimum and maximum left positions
    min_left = min(xval)
    max_left = max(xval)
    myprint1("MIN_LEFT"+str(min_left))
    myprint1("MAX_LEFT"+str(max_left))

    threshold = 50
    # Categorize blocks as left-aligned or centered
    top_aligned_blocks = []
    left_aligned_blocks = []
    right_aligned_blocks = []
    centered_blocks = []

    myprint1("-------------------------------")
    myprint1("TEST")
    for k,el in enumerate(text_elements):
        left_pos=el['bbox'][0]
        bottom_pos=el['bbox'][1]
        #myprint1("test ["+str(left_pos) +"]"+str(el['text']))
        if is_in_top_margin(bottom_pos,  threshold):
         #   myprint1("test left")
            top_aligned_blocks.append(el)
        elif is_in_left_margin(left_pos, min_left, threshold):
          #  myprint1("test left")
            left_aligned_blocks.append(el)
        elif is_in_right_margin(left_pos, max_left, threshold):
           # myprint1("test right")
            right_aligned_blocks.append(el)
        else:#if is_centered(left_pos, min_left, max_left, threshold):
            #myprint1("test centre")
            centered_blocks.append(el)

    
    myprint1("-------------------------------")
    myprint1("LEFT")
    for k,el in enumerate(left_aligned_blocks):
        left_pos=el['bbox'][0]
        text=el['text']
        myprint1(str(text)+" "+str(left_pos))
    myprint1("-------------------------------")
    myprint1("RIGHT")
    for k,el in enumerate(right_aligned_blocks):
        left_pos=el['bbox'][0]
        text=el['text']
        myprint1(str(text)+" "+str(left_pos))
    myprint1("-------------------------------")
    myprint1("CENTER")
    for k,el in enumerate(centered_blocks):
        left_pos=el['bbox'][0]
        text=el['text']
        myprint1(str(text)+" "+str(left_pos))

    myprint1("-------------------------------")
    myprint1("OUTPUT")
    current_character=None
    is_after_character=False
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        for k,el in enumerate(centered_blocks):
            left_pos=el['bbox'][0]
            text=el['text']
            parts=text.split("\n")
            Nparts=len(parts)
            if Nparts==1:    
                if text.isupper() and not is_after_character and not text.startswith("¡"):
                    myprint1("OUT  ["+str(left_pos)+"]   "+text+"    -->  CHAR   ")
                    current_character=filter_character_name(text)
                    is_after_character=True
                else:
                    myprint1("OUT  ["+str(left_pos)+"]   "+text+"    -->  DIALOG   ")
                    dialog=filter_speech(text)
                    is_after_character=False
                    if current_character!=None:
                        s=current_character+"\t"+dialog+"\n"  # New line after each row
                        file.write(s)
            else:
                for part in parts:
                    if part.isupper() and not is_after_character and not text.startswith("¡"):
                        myprint1("OUT  ["+str(left_pos)+"]   "+part+"    -->  CHAR   ")
                        current_character=filter_character_name(part)
                        is_after_character=True
                    else:
                        myprint1("OUT  ["+str(left_pos)+"]   "+part+"    -->  DIALOG   ")
                        dialog=filter_speech(part)
                        is_after_character=False
                        if current_character!=None:
                            s=current_character+"\t"+dialog+"\n"  # New line after each row
                            file.write(s)


    myprint1("-------------------------------")
    myprint1("FINISHED")
    myprint1("Converted")                       
    myprint1(converted_file_path)
    return converted_file_path,encoding

def run_convert_pdf_to_txt(file_path,absCurrentOutputFolder,centered_blocks,encoding):
    converted_file_path=""
    if ".pdf" in file_path.lower() :
        converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".converted.txt"))
    if ".pdf" in file_path.lower() :
        converted_raw_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".raw.txt"))

    current_character=None
    is_after_character=False
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        for k,el in enumerate(centered_blocks):
            left_pos=el['bbox'][0]
            text=el['text']
            parts=text.split("\n")
            Nparts=len(parts)
            if Nparts==1:    
                if text.isupper() and not is_after_character and not text.startswith("¡"):
                    myprint1("OUT1  ["+str(left_pos)+"]   "+text+"        --> CHAR   ")
                    current_character=filter_character_name(text)
                    is_after_character=True
                else:
                    myprint1("OUT1  ["+str(left_pos)+"]   "+text+"        --> DIALOG   ")
                    dialog=filter_speech(text)
                    is_after_character=False
                    if current_character!=None:
                        s=current_character+"\t"+dialog+"\n"  # New line after each row
                        file.write(s)
            else:
                for part in parts:
                    if part.isupper() and not is_after_character and not text.startswith("¡"):
                        myprint1("OUT1  ["+str(left_pos)+"]   "+part+"        --> CHAR   ")
                        current_character=filter_character_name(part)
                        is_after_character=True
                    else:
                        myprint1("OUT1  ["+str(left_pos)+"]   "+part+"        --> DIALOG   ")
                        dialog=filter_speech(part)
                        is_after_character=False
                        if current_character!=None:
                            s=current_character+"\t"+dialog+"\n"  # New line after each row
                            file.write(s)


    myprint1("-------------------------------")
    myprint1("FINISHED")
    myprint1("Converted")                       
    myprint1(converted_file_path)
    return converted_file_path,encoding
            
def convert_pdf_to_txt_pdfplum(file_path,absCurrentOutputFolder,encoding):
    myprint1("convert_pdf_to_txt")
    myprint1("currentOutputFolder             :"+absCurrentOutputFolder)
    myprint1("Input              :"+file_path)
    converted_file_path=""
    if ".pdf" in file_path.lower() :
        converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".converted.txt"))
    if ".pdf" in file_path.lower() :
        converted_raw_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".raw.txt"))
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        with open(converted_raw_file_path, 'w', encoding='utf-8') as fileraw:
                pdf = PdfReader(file_path)
                page_idx=1  
                for page_num in range(len(pdf.pages)):
                    page = pdf.pages[page_num]
                    page_idx=page_idx+1
                    if page_idx<40000:
                        # Extract tables from the page
                        myprint1("########################################################")
                        myprint1("convert_pdf_to_txt page"+str(page_idx))
                        #crop_coords = [0.68,0.20,0.95,0.90]
                        #my_width = page.width
                        #my_height = page.height
                        #my_bbox = (crop_coords[0]*float(my_width), crop_coords[1]*float(my_height), crop_coords[2]*float(my_width), crop_coords[3]*float(my_height))
                        #page = page.crop(bbox=my_bbox)
                        myprint1("extract")
                        extr=page.extract_text(TextAlignment.NONE).split("\n\n")
                        myprint1(extr)
                        text = str(extr)
                        myprint1(text)
                        myprint1("write")
                        fileraw.write(text)
                        myprint1("split")
                        text_blocks = extr.split('\n')

                        for block_num, block in enumerate(text_blocks, start=1):
                            # Skip empty blocks
                            if block.strip() == "":
                                continue
                            
                            # Print the page number, block number, and block text
                            myprint1(f"Page: {page_num + 1}")
                            myprint1(f"Block: {block_num}")
                            myprint1(f"Text: {block.strip()}")
                            myprint1("---")
                            block_position = page.get_text_block_position(block)
                            if block_position is not None:
                                # Extract the coordinates of the block position
                                x0, y0, x1, y1 = block_position

                                myprint1(f"Position: ({x0}, {y0}) - ({x1}, {y1})")
                            else:

                                myprint1(f"Position: none")

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
                                        myprint1("Add "+ current_character + " "+ speech)
                                        file.write(s)
                                        charidx=charidx+1

                                if mode=="linear":
                                    if current_character=="":
                                        current_character="__VOICEOVER"
                                    s=current_character+"\t"+speech+"\n"  # New line after each row
                                    myprint1("Add "+ current_character + " "+ speech)
                                    file.write(s)

    myprint1("Converted")                       
    myprint1(converted_file_path)
    return converted_file_path,encoding

def convert_pdf_to_txt_pdfplumber(file_path,absCurrentOutputFolder,encoding):
    myprint1("convert_pdf_to_txt")
    myprint1("currentOutputFolder             :"+absCurrentOutputFolder)
    myprint1("Input              :"+file_path)
    converted_file_path=""
    if ".pdf" in file_path.lower() :
        converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".converted.txt"))
    if ".pdf" in file_path.lower() :
        converted_raw_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).lower().replace(".pdf",".raw.txt"))
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        with open(converted_raw_file_path, 'w', encoding='utf-8') as fileraw:
            with pdfplumber.open(file_path) as pdf:
                page_idx=1  
                for page in pdf.pages[1:]:

                    page_idx=page_idx+1
                    if page_idx<40000:
                        # Extract tables from the page
                        myprint1("########################################################")
                        myprint1("convert_pdf_to_txt page"+str(page_idx))
                        #crop_coords = [0.68,0.20,0.95,0.90]
                        #my_width = page.width
                        #my_height = page.height
                        #my_bbox = (crop_coords[0]*float(my_width), crop_coords[1]*float(my_height), crop_coords[2]*float(my_width), crop_coords[3]*float(my_height))
                        #page = page.crop(bbox=my_bbox)
                        myprint1("extract")
                        extr=page.extract_text()
                        myprint1(extr)
                        text = str(extr)
                        myprint1(text)
                        myprint1("write")
                        fileraw.write(text)
                        myprint1("split")

                        text_blocks = extr.split('\n')
                        myprint1("Blocks"+str(len(text_blocks)))
                        # Iterate over each text block
                        for block in text_blocks:
                            myprint1("block"+str(block))
                            # Find the position and dimensions of the text block
                            bbox = page.bbox_for_text(block)
                            myprint1("bbox")
                            if bbox:
                                x0, top, x1, bottom = bbox
                                width = x1 - x0
                                
                                # Print the text block information
                                myprint1(f"Page: {page_idx}")
                                myprint1(f"Text: '{block}'")
                                myprint1(f"Position: ({x0}, {top}) - ({x1}, {bottom})")
                                myprint1(f"Width: {width}")
                                myprint1("---")
                            else:
                                myprint1("no bbox")

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
                                        myprint1("Add "+ current_character + " "+ speech)
                                        file.write(s)
                                        charidx=charidx+1

                                if mode=="linear":
                                    if current_character=="":
                                        current_character="__VOICEOVER"
                                    s=current_character+"\t"+speech+"\n"  # New line after each row
                                    myprint1("Add "+ current_character + " "+ speech)
                                    file.write(s)

                       
                       
    myprint1("Converted")                       
    myprint1(converted_file_path)
    return converted_file_path,encoding



def convert_pdf_title(file,table,titleIdx):
    myprint1("ROW LEN="+str(len(table[0]))+str(table[0]))
    for row in table[1:]:  # Skip header
        lastIdx=-1
        for i in range(len(row)):
            if row[i]:
                myprint1("has data"+str(row[i]))
                lastIdx=i
        myprint1("---------------------------------"+str(lastIdx))
        title = row[lastIdx]
        myprint1("item "+str(i)+" "+str(row[i]))
        myprint1("ROW LEN="+str(len(row))+str(row))
        myprint1(str(row))
        myprint1(">>>>>> TITLE"+str(title))
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
            myprint1("test_pdf_headers"+str(header))
            dialogueCol=-1
            characterIdCol=-1
            titlesCol=-1
    
            myprint1("headers"+str(header))
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
                myprint1("test_pdf_header pdf_mode_title")
                convert_pdf_title(file,table,titlesCol)
                return True


def detect_word_table(table,forceMode="",forceCols={}):
    myprint1("detect_word_table")
    for i in range(3):
        myprint1("try header "+str(i))
        header=table.rows[i]
        myprint1("header read ")        
        success, mode, character,dialog,map_= detect_word_header(header,forceMode,forceCols)
        myprint1("header success= "+str(success)+" "+str(map_))        
        if success:
            return success, mode, character,dialog,map_
    return False,"",-1,-1,{}


def detect_word_header(header,forceMode="",forceCols={}):
    myprint1("-------------- test word header -----------------")
    myprint1("forceMode"+str(forceMode))
    myprint1("forceCols"+str(forceCols))
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
            myprint1("CHARACTER "+str(characterIdCol))
            myprint1("DIALOG "+str(dialogueCol))
        else:
            myprint1("UNKNOWN FORCE MODE")
    else:
        myprint1("Header:")
        for cell in header.cells:
            t=cell.text.strip()
            myprint1("   * "+str(t))
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
            elif t=="Titles" or t=="Title":
                titlesCol=idx
            elif t=="Dialog With \nSpeaker Id":
                dialogWithSpeakerId=idx
            elif t=="COMBINED CONTINUITY":
                combinedContinuityCol=idx
            idx=idx+1

    myprint1(f"dialogCol {dialogueCol}")
    myprint1(f"characterIdCol {characterIdCol}")
    myprint1(f"combinedContinuityCol {combinedContinuityCol}")
    myprint1(f"titlesCol {titlesCol}")
    myprint1(f"sceneDescriptionCol {sceneDescriptionCol}")
    myprint1(f"dialogWithSpeakerId {dialogWithSpeakerId}")
    myprint1("detect_word_header assigned")
    docx_mode_dialogue_characterid= dialogueCol>-1 and characterIdCol>-1
    #docx_mode_combined_continuity= combinedContinuityCol>-1
    docx_mode_scenedescription= sceneDescriptionCol>-1 and titlesCol>-1
    docx_mode_dialogwithspeakerid=dialogWithSpeakerId>-1
    myprint1("detect_word_header assigned mode")

    mode=None
    character=None
    dialog=None
    map_={}
    idx=0
    myprint1("detect_word_headermap")

    if dialogueCol>-1 and characterIdCol>-1:
        mode="SPLIT"
        character=characterIdCol
        dialog=dialogueCol
    if sceneDescriptionCol>-1 and titlesCol>-1:
        mode="COMBINED"
        character=titlesCol
        dialog=titlesCol
    if combinedContinuityCol>-1 or dialogWithSpeakerId>-1:
        mode="COMBINED"
        character=dialogWithSpeakerId
        dialog=dialogWithSpeakerId

    myprint1(f"detect_word_header gen map d={dialog} c={character}")
    for cell in header.cells:
        myprint1(f"  test idx={idx} d={dialog}")
        t=cell.text.strip()
        if idx==dialog and idx==character:
            t='BOTH'
        elif idx==dialog:
            t='DIALOG'
        elif idx==character:
            t='CHARACTER'
        else:
            t='NONE'
        
        map_[idx]={'type':t}
        idx=idx+1

    myprint1("detect_word_header done")
    myprint1("map_"+str(map_))
    myprint1("mode"+str(mode))
    myprint1("dialog"+str(dialog))
    myprint1("character"+str(character))
    

    if docx_mode_dialogue_characterid:
        return True, mode, character,dialog,map_
    elif docx_mode_scenedescription:
        return True, mode, character,dialog,map_
    elif docx_mode_dialogwithspeakerid:
        return True, mode, character,dialog,map_
   # elif docx_mode_combined_continuity:
    #    return True, mode, character,dialog,map_
    else:
        myprint1("Tables but no ")
        return False, mode, character,dialog,map_


def test_word_header_and_convert(file,table,header,forceMode="",forceCols={}):
            myprint1("test word header")
            myprint1("forceMode"+str(forceMode))
            myprint1("forceCols"+str(forceCols))
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
                    myprint1("CHARACTER "+str(characterIdCol))
                    myprint1("DIALOG "+str(dialogueCol))
                else:
                    myprint1("UNKNOWN FORCE MODE")
            else:
                for cell in header.cells:
                    t=cell.text.strip()
                    myprint1("Header cell"+str(t))
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
                    myprint1("Headers found")
                else:
                    myprint1("Headers not found")
                    myprint1(" CharacterId"+str(characterIdCol))
                    myprint1(" sceneDescriptionCol"+str(sceneDescriptionCol))
                    myprint1(" dialogueCol"+str(dialogueCol))
                    myprint1(" titlesCol"+str(titlesCol))
                    myprint1(" dialogWithSpeakerId"+str(dialogWithSpeakerId))
                    myprint1(" combinedContinuityCol"+str(combinedContinuityCol))
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
                myprint1("Tables but no ")
                return False

def convert_xlsx_to_txt(file_path,absCurrentOutputFolder,forceMode="",forceCols={}):
    myprint1("convert_xlsx_to_txt")
    myprint1("currentOutputFolder             :"+absCurrentOutputFolder)
    myprint1("Input             :"+file_path)
    df = pd.read_excel(file_path)
    # Load the Excel file
    
    # Extract the columns "Character" and "English"
    character_column = df["Character"]
    english_column = df["English"]
    
    # Prepare the text file content
    lines = []
    for char, eng in zip(character_column, english_column):
        if pd.notna(char) and pd.notna(eng):
            eng=eng.replace("\n"," ")
            myprint1("Add line"+str(eng))
#            if False:
 #           nblines=eng.count("\n")

  #          if nblines==1:
            myprint1("Add line linear")
            if eng.startswith("- "):
                    eng=eng.lstrip("- ")

            if eng.count("- ")>0:
                spl=eng.split("- ")
                chars=char.split("-")
                if len(spl) == len(chars):
                    for index,k in enumerate(spl):
                        lines.append(f"{chars[index].upper()}\t{k}")        
                else:
                    lines.append(f"{char.upper()}\t{eng}")                            
            elif eng.count("-")>0:
                spl=eng.split("-")
                chars=char.split("-")
                if len(spl) == len(chars):
                    for index,k in enumerate(spl):
                        lines.append(f"{chars[index].upper()}\t{k}")        
                else:
                    lines.append(f"{char.upper()}\t{eng}")                            


            else:      
                lines.append(f"{char.upper()}\t{eng}")
   #         else:
    #            myprint1("Add line split")
     #           spl=eng.split("\n")
      #          for k in spl:
       #             lines.append(f"{char.upper()}\t{k}")


    converted_file_path= os.path.join(absCurrentOutputFolder,os.path.basename(file_path).replace(".xlsx",".converted.txt"))

    # Write to a text file
    with open(converted_file_path, 'w') as f:
        for line in lines:
            f.write(line + '\n')
    myprint1("Done")
    return converted_file_path

def get_paragraph_indentation(paragraph):
    """
    Returns the effective left indentation of the paragraph, considering the style hierarchy.
    """
    indent = paragraph.paragraph_format.left_indent
    style = paragraph.style
    while indent is None and style is not None:
        indent = style.paragraph_format.left_indent
        style = style.base_style
    return indent

def get_paragraph_style_hierarchy(paragraph):
    """
    Returns the style hierarchy of a paragraph.
    """
    styles = []
    style = paragraph.style
    while style is not None:
        styles.append(style.name)
        style = style.base_style
    return styles

def inspect_paragraphs_with_style_hierarchy(para):
        text = para.text.strip()
        indent = get_paragraph_indentation(para)
        styles = get_paragraph_style_hierarchy(para)
        myprint1(str({
            'text': text,
            'indent': indent,
            'style_hierarchy': ' > '.join(styles)
        }))

def word_has_paragraph_style_dialog(para):
    styles=get_paragraph_style_hierarchy(para)
    styles = ' > '.join(styles)
    return "DIALOG" in styles

def word_has_paragraph_style_character(para):
    styles=get_paragraph_style_hierarchy(para)
    styles = ' > '.join(styles)
    return "CHARACTER" in styles

def convert_word_withstyles_to_plaintext(doc,file_path,absCurrentOutputFolder):
    converted_file_path=get_docx_to_txt_converted_filepath(file_path,absCurrentOutputFolder)
    current_character=None
    current_dialog=None
    with open(converted_file_path, 'w', encoding='utf-8') as file:
        for para in doc.paragraphs: 
            if word_has_paragraph_style_character(para):
                char=filter_character_name(para.text.upper())
                current_character=char
            if word_has_paragraph_style_dialog(para):
                speech=para.text
                current_dialog=filter_speech(speech)
            if current_character != None and current_dialog!=None:
                s=current_character+"\t"+current_dialog+"\n"  # New line after each row
                myprint1("Add "+ current_character + " "+ current_dialog)
                file.write(s)
                current_dialog=None
    return converted_file_path
def detect_word_styles_character_dialog(doc):
    has_dialog_style=False
    has_character_style=False
    for para in doc.paragraphs:
        if word_has_paragraph_style_dialog(para):
            has_dialog_style=True
        if word_has_paragraph_style_character(para):
            has_character_style=True
    if has_dialog_style and has_character_style:
        return True
    return False

def detect_plaintext_indented(doc):
    isIndented=False
    nParagraphs=0
    nIndented=0
    # Iterate over paragraphs
    for para in doc.paragraphs:
        myprint1('-------')
        if len(para.text)>0:
            myprint1(str(para.text)+" ")
            nParagraphs=nParagraphs+1
            indent = para.paragraph_format.left_indent
            rightindent=para.paragraph_format.right_indent
            first= para.paragraph_format.first_line_indent
            for  run in para.runs:
            
                myprint1("RUN UPPER"+str(run.text)+" "+ "allcaps="+str(run.font.all_caps))
                if run.font.all_caps:
                    myprint1("UPPER")
            myprint1("leftindent="+str(indent)+" rightindent="+str(rightindent)+" first="+str(first))
            inspect_paragraphs_with_style_hierarchy(para)
            if indent is not None:
                myprint1("INDENTED "+str(para.text))
                nIndented=nIndented+1
        
    indentedPc=nIndented/nParagraphs
    myprint1("Plaintext detect indented")
    myprint1("Plaintext "+str(nIndented)+" / "+str(nParagraphs)+" paragraphs= "+str(indentedPc)+" indented")
    if indentedPc>0.1:
        return True
    else:
        return False        
            
def get_docx_to_txt_converted_filepath(file_path,absCurrentOutputFolder):
    converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).replace(".docx",".converted.txt"))
    return converted_file_path
def get_doc_to_txt_converted_filepath(file_path,absCurrentOutputFolder):
    converted_file_path=os.path.join(absCurrentOutputFolder, os.path.basename(file_path).replace(".doc",".converted.txt"))
    return converted_file_path
def convert_word_to_txt(file_path,absCurrentOutputFolder,forceMode="",forceCols={}):
    myprint1("convert_docx_to_txt")
    myprint1("currentOutputFolder : "+absCurrentOutputFolder)
    myprint1("Input               : "+file_path)
    converted_file_path=""
    if ".docx" in file_path:
        myprint1(".docx")
        converted_file_path=get_docx_to_txt_converted_filepath(file_path,absCurrentOutputFolder)
    elif ".doc" in file_path:
        myprint1(".doc")
        converted_file_path=get_docx_to_txt_converted_filepath(file_path,absCurrentOutputFolder)

#        myprint1("Convert doc to docx  ")
 #       docx_file_path = convert_doc_to_docx(file_path,absCurrentOutputFolder)
  #      myprint1("Output         :"+os.path.abspath(docx_file_path))
   #     converted_file_path=os.path.abspath(docx_file_path).replace(".docx",".converted.txt")
    #    file_path=docx_file_path
    else:
        myprint1("other extension")

    myprint1("Converted file path : "+converted_file_path)    
    myprint1("Doc opening         : "+file_path)
    doc = Document(file_path)
    myprint1("Doc opened")

    with open(converted_file_path, 'w', encoding='utf-8') as file:
        # Check if there are any tables in the document
        myprint1("Table count             : "+str(len(doc.tables)))

        if len(doc.tables) > 0:
            # Get the first table
            table = doc.tables[0]
            
            headerSuccess=False
            for i in range(3):
                header=table.rows[i]
                success=test_word_header_and_convert(file,table,header,forceMode=forceMode,forceCols=forceCols)
                if success:
                    headerSuccess=True
                    break
            if not headerSuccess:
                return ""

        else:
            hasWordStyles=detect_word_styles_character_dialog(doc)
            if hasWordStyles:
                convert_word_withstyles_to_plaintext(doc,file_path,absCurrentOutputFolder)
            else:
                myprint1("No tables, convert to plain text")
                isIndented=detect_plaintext_indented(doc)
                convert_docx_plain_text(file,doc)
                myprint1("Converted file path : "+converted_file_path)  
                myprint1("converted to plain text")
        
    return converted_file_path

def get_all_characters(breakdown):
    #myprint1("get_all_characters")
    all_characters=[]
    for item in breakdown:
        if item["type"]=="SPEECH":
            character=item["character"]
            if character==None:
                myprint1("ERR")
                exit()
            #myprint1("test  CHARACTER"+str(character)+" "+str(all_characters))

            if not character in all_characters:
                myprint1("ADD NEW CHARACTER"+character)
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
        myprint1("Match found:", match.groups())
    else:
        myprint1("No match found.")

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
    myprint1("merge_breakdown_character_by_replacelist")
    checkIfAlreadyNamed=False
    for item in breakdown:
        if item["type"]=="SPEECH":
            character=item["character"]
            if character in replace_list:
                firstchar=replace_list[character]
                item['character']=firstchar                   

    return breakdown

def is_multiple_character(char):
    return " AND " in char

def split_AND_character(breakdown):
    myprint1("merge_breakdown_character_by_replacelist")
    for item in breakdown:
        if item["type"]=="SPEECH":
            character=item["character"]
            if is_multiple_character(character):
                characters=character.split(" AND ")
                first=characters[0]
                chidx=0
                for ch in characters:
                    if chidx==0:
                        item['character']=ch
                    else:
                        item2={
                            'character':ch,
                            'type':item['type'],
                            'scene_id':item['scene_id'],
                            'line_idx':item['line_idx'],
                            'character_raw':character,
                            'speech':item['speech']
                        }
                        breakdown.append(item2)

                    chidx=chidx+1

    return breakdown

def merge_breakdown_character_talking_to(breakdown,all_characters):
    myprint1("merge_breakdown_character_talking_to")
    replaceList={}
    checkIfAlreadyNamed=False
    for item in breakdown:
        #smyprint1(str(item))
        if item["type"]=="SPEECH":
            character=item["character"]
            if character!=character.strip():
                character=character.strip()

            splitable=hasSplitable(character)
            if splitable!="":
                #myprint1(" has to"+character)
                characters=character.split(splitable)
                #myprint1(" split"+str(characters))

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
                        #myprint1("REPLACE "+character+" with "+str(firstchar))
                        replaceList[character]=firstchar
                        item['character']=firstchar                   
                else:
                    firstchar=characters[0].strip()
                    #myprint1("REPLACE"+character+" with "+str(firstchar))
                    replaceList[character]=firstchar
                    item['character']=firstchar                   

    return breakdown,replaceList
def filter_speech2(s):
    res=s.replace("♪","").replace("Â§","").replace("§","")
    #filter songs
    return res

def save_string_to_file(text, filename):
    """Saves a given string `text` to a file named `filename`."""
    myprint1(" > Write to "+filename)
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(text)


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
    myprint1("  > -----------------------------------")
    myprint1("  > SCRIPT PARSER version 1.3")
    myprint1("  > Script path       : "+script_path)
    myprint1("  > Output folder     : "+output_path)
    myprint1("  > Script name       : "+script_name) 
    myprint1("  > Forced encoding   : "+encoding)
    myprint1("  > Counting method   : "+countingMethod)

    if not os.path.exists(output_path):
        os.mkdir(output_path)

    file_name = os.path.basename(script_path)
    name, extension = os.path.splitext(file_name)
    myprint1("  > File name         : "+file_name)
    myprint1("  > Extension         : "+extension)
    if not is_supported_extension(extension):
        myprint1("  > File type "+extension+" not supported.")
        return

    if extension.lower()==".docx":

        if "CLEAR CUT" in file_name:
            myprint1(" !!!!!!!!!! FORCE")
            forceMode="DETECT_CHARACTER_DIALOG"
            forceCols={
                "CHARACTER":5,
                "DIALOG":6
            }
        converted_file_path=convert_word_to_txt(script_path,forceMode=forceMode,forceCols=forceCols)
        if len(converted_file_path)==0:
            myprint1("  > Conversion failed 1")
            return None,None,None,None,None,None
        return process_script(converted_file_path,output_path,script_name,countingMethod,forceMode=forceMode,forceCols=forceCols)




    if extension.lower()==".xlsx":

        converted_file_path=convert_xlsx_to_txt(script_path,forceMode=forceMode,forceCols=forceCols)
        if len(converted_file_path)==0:
            myprint1("  > Conversion failed 1")
            return None,None,None,None,None,None
        return process_script(converted_file_path,output_path,script_name,countingMethod,forceMode=forceMode,forceCols=forceCols)





    if extension.lower()==".pdf":
        converted_file_path,encoding=convert_pdf_to_txt(script_path)
        myprint1("process 1")
        if len(converted_file_path)==0:
            myprint1("  > Conversion failed 1")
            return None,None,None,None,None,None
        myprint1("process")
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
    myprint1("  > Encoding          : "+encoding_info['encoding'])
    encoding_used=encoding_info['encoding']
    myprint1("  > Encoding used     : "+encoding_used)

    encoding_tested=test_encoding(script_path)
    encoding_used=encoding_tested
    if not (encoding==""):
        myprint1("  > Force encoding    :"+encoding)
        encoding_used=encoding
    scene_separator=getSceneSeparator(script_path,encoding_used)
    myprint1("  > Scene separator   : "+scene_separator)

    character_mode=detectCharacterSeparator(script_path,encoding_used)
    myprint1("  > Character mode    : "+str(character_mode))
    
    character_sep_type=getCharacterSepType(character_mode)
    myprint1("  > Character sep type    : "+str(character_sep_type))
    
    if scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
        current_scene_id="Scene 1"
    # Open the file and process each line
    line_idx=1
    isEmptyLine=False
    multiline_current_uppercase_text = None
    multiline_current_lines_of_text = []
    multiline_in_pattern = False
    with open(script_path, 'r', encoding=encoding_used) as file:
        myprint1("  > Opened    : "+str(script_path))
        for line in file:
            myprint1("  > Opened    : "+str(line))
            line = line  # Remove any leading/trailing whitespace
            trimmed_line = line.strip()  # Remove any leading/trailing whitespace
    
            isNewEmptyLine=len(trimmed_line)==0
            myprint1("  > Sep")

            if scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
                if (not isNewEmptyLine) and  (isEmptyLine and wasEmptyLine):
                    current_scene_count=current_scene_count+1
                    current_scene_id = extract_scene_name(line,scene_separator,current_scene_count)
                    if is_verbose:
                        myprint1("  > ---------------------------------------")
                    #myprint1(f"Scene Line: {line}")
    
            isEmptyLine=len(trimmed_line)==0
            if character_sep_type=="CHARACTER_MODE_SINGLELINE":
                if is_verbose:
                    myprint1("  > Line "+str(line_idx))
                if len(trimmed_line)>0:
                    myprint1("  > trimmed")   
                    if is_scene_line(line) or (isEmptyLine and wasEmptyLine):
                        current_scene_count=current_scene_count+1
                                        
                        current_scene_id = extract_scene_name(line,scene_separator,current_scene_count)
                        breakdown.append({"line_idx":line_idx,"scene_id":current_scene_id,"type":"SCENE_SEP" })    
                        if is_verbose:
                            myprint1("  > --------------------------------------")
                        myprint1(f"  > Scene Line: {current_scene_id}")
                    else:
                            myprint1("  > not scene")   
                            if True:#current_scene_id!=1:
                                is_speaking=is_character_speaking(trimmed_line,character_mode)
                                myprint1("  > speaking")   

                                if is_verbose:
                                    myprint1("    IsSpeaking "+str(is_speaking)+" "+trimmed_line)
                                if is_speaking:
                                    myprint1("  > speaking1"+str(trimmed_line)+" "+str(character_mode))   

                                    character_name=extract_character_name(trimmed_line,character_mode)
                                    myprint1("  > speaking 2b"+str(character_name))   
                                    character_name=filter_character_name(character_name)
                                    myprint1("  > speaking 2")   
                                    if is_verbose:
                                        myprint1("   name="+str(character_name))
                                    if not character_name == None:
                                        if is_character_name_valid(character_name):
                                            #remove character name for stats
                                            spoken_text=extract_speech(trimmed_line,character_mode,character_name)
                                            spoken_text=filter_speech(spoken_text)
                                            
                                            breakdown.append({"scene_id":current_scene_id,
                                                            "character_raw":character_name,
                                                            "line_idx":line_idx,"speech":spoken_text,"type":"SPEECH", "character":character_name })    
                                            if is_verbose:
                                                myprint1("   text="+str(spoken_text))
                                        
                                else:
                                    breakdown.append({"line_idx":line_idx,"text":trimmed_line,"type":"NONSPEECH" })    
            elif character_sep_type=="CHARACTER_MODE_MULTILINE":
                    # Check for uppercase text
                if trimmed_line.isupper() and not multiline_in_pattern:
                    multiline_current_uppercase_text = trimmed_line
                    multiline_current_lines_of_text = []
                    multiline_in_pattern = True
                elif multiline_in_pattern:
                    if trimmed_line == '':  # Empty line
                        # Check if we have hit the double newline
                        if not multiline_current_lines_of_text or multiline_current_lines_of_text[-1] == '':
                            # Double newline indicates end of current pattern
                            if multiline_current_uppercase_text:
                                character_name=filter_character_name(multiline_current_uppercase_text)
                                speech='\n'.join(multiline_current_lines_of_text).strip()
                                breakdown.append({"scene_id":current_scene_id,
                                                            "character_raw":character_name,
                                                            "line_idx":line_idx,"speech":speech,"type":"SPEECH", "character":character_name })    
                                           
                                
                            multiline_current_uppercase_text = None
                            multiline_current_lines_of_text = []
                            multiline_in_pattern = False
                        else:
                            multiline_current_lines_of_text.append('')
                    else:
                        multiline_current_lines_of_text.append(trimmed_line)
                           
                                        
            wasEmptyLine=isEmptyLine
            line_idx=line_idx+1

    if character_sep_type=="CHARACTER_MODE_MULTILINE":
        # Handle the case where the file ends while still in pattern
        if multiline_current_uppercase_text and multiline_current_lines_of_text:
            character_name=filter_character_name(multiline_current_uppercase_text)
            speech='\n'.join(multiline_current_lines_of_text).strip()          
            breakdown.append({"scene_id":current_scene_id,
                                    "character_raw":character_name,
                                    "line_idx":line_idx,"speech":speech,"type":"SPEECH", "character":character_name })    
                

    myprint1("breakdown"+str(breakdown))

    all_characters=get_all_characters(breakdown)
    breakdown,replaceList=merge_breakdown_character_talking_to(breakdown,all_characters)
    all_characters=get_all_characters(breakdown)
    #myprint1("all_characters"+str(all_characters))
    #myprint1("replacelist"+str(replaceList))

    replace_map=map_semi_duplicates(all_characters)
   # myprint1("replace_map"+str(replace_map))

    breakdown=merge_breakdown_character_by_replacelist(breakdown,replace_map)
    breakdown=split_AND_character(breakdown)
    all_characters=get_all_characters(breakdown)


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

    if len(character_order_map)>0:                                    
        csv_file_path =output_path+script_name+"-comptage-detail.csv"
        data = [
        ]
        for key in character_order_map:
            data.append([
                str(character_order_map[key])+" - "+str(key),str(character_linecount_map[key]),
                str(character_textlength_map[key]),
                str(math.ceil(character_textlength_map[key]/40))])


        myprint1("  > Convert to csv.")
        with open(csv_file_path, mode='w', newline='',encoding=encoding_used) as file:
            writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            
            # Write data to the CSV file
            for row in data:
                writer.writerow(row)


        myprint1("  > Convert to xslx.")
        convert_csv_to_xlsx(output_path+script_name+"-comptage-detail.csv",output_path+script_name+"-comptage-detail.xlsx", script_name,encoding_used)
    myprint1("  > Parsing done.")
    return breakdown, character_scene_map,scene_characters_map,character_linecount_map,character_order_map,character_textlength_map


#process_script("scripts/COMPTAGE/20012020.txt","20012020/","20012020","ALL")    
#process_script("scripts/examples/YOU CAN'T RUN FOREVER_SCRIPT_VO.txt","YOUCANRUNFOREVER_SCRIPT_VOe/","YOU CAN'T RUN FOREVER_SCRIPT_VO","ALL")    
#process_script("190421-1.txt","190421-1/","190421-1.txt")
#process_script("scripts/examples/LATENCY.docx","LATENCY/","LATENCY","ALL")    
#process_script("scripts/examples/LATENCY.docx","LATENCY/","LATENCY","ALL")    
#process_script("scripts/examples/Gods of the deep - CCSL.docx","GODS/","GODS","ALL")    
#process_script("scripts/examples/Blackwater Lane.docx","Blackwater Lane/","Blackwater Lane","ALL")    

            