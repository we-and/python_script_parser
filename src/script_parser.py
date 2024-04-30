from PyPDF2 import PdfReader
import re
import os
import math
#script_path="scripts2/YOU CAN'T RUN FOREVER_SCRIPT_VO.txt"
#output_path="YOU CANT RUN FOREVER_SCRIPT_VOc/"
#script_name="YOU CANT RUN FOREVER_SCRIPT_VO"

#script_path="scripts2/ZERO11.txt"
#output_path="ZERO11c/"
#script_name="ZERO11b"


import chardet

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
                print("Testing encoding  : "+enc)

                for line in file:
                    line = line.strip()  # Remove any leading/trailing whitespace
            return enc
        except UnicodeDecodeError:
            print(f"Failed decoding with {enc}")
    return "?"    






def find_first_uppercase_sequence(line):
    """Finds the first sequence of contiguous uppercase words in a line."""
    # Regex pattern to match the first sequence of contiguous uppercase words separated by spaces
    pattern = re.compile(r'\b([A-Z]+(?:\s[A-Z]+)*)\b')
    match = re.search(pattern, line)
    if match:
        return match.group(0)
    return None  # Return None if no uppercase sequence is found

#################################################################
# OUTPUT UTILS
def write_character_map_to_file(character_map, filename):
    """Writes the character to scene map to a specified file."""
    with open(filename, 'w', encoding='utf-8') as file:
        for character, scenes in character_map.items():
            file.write(f"{character}: {scenes}\n")


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

def extract_scene_name(line,scene_separator):
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

def getSceneSeparator(script_path,encod):
    mode="?"
    # Open the file and process each line
    print("Open as          : "+encod)

    with open(script_path, 'r', encoding=encod) as file:
        print("Opened         : "+encod)

        for line in file:
            line = line.strip()  # Remove any leading/trailing whitespace
            print(line)
            if matches_format_parenthesis_name_timecode(line):
                print("PARENTHESIS_NAME_TIMECODE")
                return "PARENTHESIS_NAME_TIMECODE"
            
            elif matches_number_parenthesis_timecode(line):
                print("NAME_PARENTHESIS_TIMECODE")
                return "NAME_PARENTHESIS_TIMECODE"

    if mode=="?":
        print("Try ")
        n_sets_of_empty_lines=count_consecutive_empty_lines(script_path,2,encod)
        #print("check empty lines count"+str(n_sets_of_empty_lines))
        if n_sets_of_empty_lines>1:
            print( "Found EMPTYLINES_SCENE_SEPARATOR in "+line)
            return "EMPTYLINES_SCENE_SEPARATOR"

    print("DEFAULT")
    return mode    


#################################################################
# CHARACTER SEPARATOR
def is_matching_character_speaking(line):
    """Checks if the line indicates a character speaking."""
    # Regex pattern to match lines that start with text, followed by a tab, then more text
    pattern = re.compile(r"^\S+\s+\t.*$")
    pattern = re.compile(r"^\S+\s*\t.*$")

    return bool(re.match(pattern, line))

def is_character_speaking(line,scene_separator):
    is_match= is_matching_character_speaking 
    if is_match:
        name= extract_character_name(line,scene_separator)
        return not is_didascalie(name) and not is_ambiance(name)
    else:
        return False
    
def extract_character_name(line,scene_separator):
    if isSeparatorParenthesisNameTimecode(scene_separator) or isSeparatorNameParenthesisTimecode(scene_separator):
        return extract_character_name_timecodemode(line)
    elif isSeparatorEmptyLinesTimecode(scene_separator): 
        return extract_character_name_consecutiverows(line)  
    else:
        print("ERROR wrong mode")

def extract_character_name_consecutiverows(line):
    """
    Extracts the first part of the line which is uppercase text,
    given the line format is uppercase text followed by at least 8 spaces and more text.
    Returns the uppercase text if the pattern matches, otherwise returns None.
    """
    match = re.match(r"^([A-Z]+)\s{8,}.*$", line)
    if match:
        return match.group(1)
    return None

def extract_character_name_timecodemode(line):
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
    return line



wasEmptyLine=False
#################################################################
# UTILS
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












#################################################################
# PROCESS
def process_script(script_path,output_path,script_name):
    print("-----------------------------------")
    print("SCRIPT PARSER version 1.3")
    print("Script path       : "+script_path)
    print("Output folder     : "+output_path)
    print("Script name       : "+script_name)

    if not os.path.exists(output_path):
        os.mkdir(output_path)
    uppercase_lines=[]
    current_scene_id=""
    scene_characters_map={}
    character_linecount_map={}
    character_order_map={}
    character_order={}
    character_count=1
    character_textlength_map={}
    character_scene_map={}
    current_scene_count=1

    is_verbose=False
    encoding_info = detect_file_encoding(script_path)
    print("Encoding          : "+encoding_info['encoding'])
    encoding_used=encoding_info['encoding']
    print("Encoding used     : "+encoding_used)

    encoding_tested=test_encoding(script_path)
    encoding_used=encoding_tested
    scene_separator=getSceneSeparator(script_path,encoding_used)
    print("Scene separator   : "+scene_separator)


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
                    current_scene_id = extract_scene_name(line,scene_separator)
                    if is_verbose:
                        print("---------------------------------------")
                    print(f"Scene Line: {line}")
    
    
    
            isEmptyLine=len(trimmed_line)==0
            if is_verbose:
                print("Line "+str(line_idx))
            if len(trimmed_line)>0:
                if is_scene_line(line) or (isEmptyLine and wasEmptyLine):
                    current_scene_count=current_scene_count+1
                    current_scene_id = extract_scene_name(line,scene_separator)
                    if is_verbose:
                        print("---------------------------------------")
                    print(f"Scene Line: {current_scene_id}")
                else:
                        if True:#current_scene_id!=1:
                            is_speaking=is_character_speaking(trimmed_line,scene_separator)
                            if is_verbose:
                                print("    IsSpeaking "+str(is_speaking)+" "+trimmed_line)
                            if is_speaking:
                                character_name=extract_character_name(trimmed_line,scene_separator)
                                #character_name=filter_character_name(character_name)
                                if is_verbose:
                                    print("   name="+str(character_name))
                                if not character_name == None:
                                    text=trimmed_line.replace(character_name,"").strip()
                                    if is_verbose:
                                        print("   text="+str(text))
                                    if character_name not in scene_characters_map:
                                        scene_characters_map[character_name] = set()
                                    scene_characters_map[character_name].add(current_scene_id)
                                    
                                    if current_scene_id not in character_scene_map:
                                        character_scene_map[current_scene_id] = set()
                                    character_scene_map[current_scene_id].add(character_name)
                                    
                                    if character_name not in character_linecount_map:
                                        character_linecount_map[character_name]=1
                                    else:
                                        character_linecount_map[character_name]=character_linecount_map[character_name]+1

                                    if character_name not in character_textlength_map:
                                        character_textlength_map[character_name]=1
                                    else:
                                        character_textlength_map[character_name]=character_textlength_map[character_name]+len(text)


                                    if character_name not in character_order_map:
                                        character_order_map[character_name]=character_count
                                        character_count=character_count+1
    #                               else:
    #                                    character_order_map[character_name]=character_order_map[character_name]+1    
            wasEmptyLine=isEmptyLine
            line_idx=line_idx+1

    #character_scene_presence=sort_dict_values(character_scene_presence)
    #scene_characters_presence=sort_dict_values(scene_characters_presence)
    write_character_map_to_file(character_scene_map, output_path+"character_by_scenes.txt")
    write_character_map_to_file(scene_characters_map, output_path+"scenes_by_character.txt")
    write_character_map_to_file(character_linecount_map, output_path+"character_linecount.txt")
    write_character_map_to_file(character_order_map, output_path+"character_order.txt")
    write_character_map_to_file(character_textlength_map, output_path+"character_textlength.txt")

    def save_string_to_file(text, filename):
        """Saves a given string `text` to a file named `filename`."""
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)

    #print(character_order_map)
    s="#Role,Lignes,Nb charactères,Répliques\n"
    for key in character_order_map:
        s=s+str(character_order_map[key])+" - "+str(key)+","+str(character_linecount_map[key])+","+str((character_textlength_map[key]))+","+str(math.ceil(character_textlength_map[key]/40))+"\n"
    save_string_to_file(s, output_path+script_name+"-recap.csv")



#process_script("scripts/examples/EBDEF10.txt","EBDEF10e/","EBDEF10")    
process_script("190421-1.txt","190421-1","190421-1.txt")