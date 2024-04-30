from PyPDF2 import PdfReader
import re
import os
script_path="scripts2/ZERO11.txt"
output_path="ZERO11/"
print("-----------------------------------")
print("SCRIPT PARSER")
print("v1.3 ")
if not os.path.exists(output_path):
    os.mkdir(output_path)
uppercase_lines=[]
current_scene_id=""
scene_characters_presence={}
character_scene_presence={}

import chardet

def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result
encoding_info = detect_file_encoding(script_path)
print(encoding_info)
def find_first_uppercase_sequence(line):
    """Finds the first sequence of contiguous uppercase words in a line."""
    # Regex pattern to match the first sequence of contiguous uppercase words separated by spaces
    pattern = re.compile(r'\b([A-Z]+(?:\s[A-Z]+)*)\b')
    match = re.search(pattern, line)
    if match:
        return match.group(0)
    return None  # Return None if no uppercase sequence is found

def write_character_map_to_file(character_map, filename):
    """Writes the character to scene map to a specified file."""
    with open(filename, 'w', encoding='utf-8') as file:
        for character, scenes in character_map.items():
            file.write(f"{character}: {scenes}\n")


def matches_timecode_format(line):
    """Checks if the line matches the specified timecode format."""
    # Regex pattern to match lines that start with '(', include a '-', have a timecode, and end with ')'
    pattern = re.compile(r"\([^\)]*-\s*\d{2}:\d{2}:\d{2}:\d{2}\)$")
    return bool(re.search(pattern, line))

def matches_timecode_format2(line):
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
    isSceneLine=matches_timecode_format(line) or matches_timecode_format2(line)
    print("IsScene    "+str(isSceneLine)+" "+line)
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
    if scene_separator=="NAME_PARENTHESIS_TIMECODE" or scene_separator=="PARENTHESIS_NAME_TIMECODE":
        if matches_timecode_format(line):
            return extract_scene_name1(line)
        elif matches_timecode_format2(line):
            return extract_scene_name2(line)
    elif scene_separator=="EMPTYLINES_SCENE_SEPARATOR":
        return "Scene "+current_scene_count
    else:
        return "?"


def is_matching_character_speaking(line):
    """Checks if the line indicates a character speaking."""
    # Regex pattern to match lines that start with text, followed by a tab, then more text
    pattern = re.compile(r"^\S+\s+\t.*$")
    pattern = re.compile(r"^\S+\s*\t.*$")

    return bool(re.match(pattern, line))

def is_character_speaking(line):
    is_match= is_matching_character_speaking 
    if is_match:
        name= extract_character_name(line)
        return not is_didascalie(name) and not is_ambiance(name)
    else:
        return False
    
def extract_character_name(line):
    """Extracts the character name from a line where the name is followed by a tab and then dialogue."""
    # Split the line at the first tab character
    parts = line.split('\t', 1)  # The '1' limits the split to the first occurrence of '\t'
    if len(parts) > 1:
        return parts[0].strip()  # Return the first part, stripping any extra whitespace
    return None  # Return None if no tab is found, indicating an improperly formatted line

def is_didascalie(name):
    return name=="DIDASCALIES"
def is_ambiance(name):
    return name=="AMBIANCE"

def filter_character_name(line):
    return line

encoding="utf-8"
encoding="ISO-8859-1"

wasEmptyLine=False

def count_consecutive_empty_lines(file_path, n):
    """Counts occurrences of exactly n consecutive empty lines in a file."""
    with open(file_path, 'r', encoding='utf-8') as file:
        count_empty = 0
        occurrences = 0
        previous_empty = False

        for line in file:
            # Check if the current line is empty or contains only whitespace
            if line.strip() == '':
                count_empty += 1
                previous_empty = True
            else:
                if previous_empty and count_empty == n:
                    occurrences += 1
                count_empty = 0
                previous_empty = False

        # Check at the end of the file if the last lines were empty
        if count_empty == n:
            occurrences += 1

    return occurrences

def getSceneSeparator(script_path):
    mode=""
    # Open the file and process each line
    with open(script_path, 'r', encoding=encoding) as file:
        for line in file:
            line = line.strip()  # Remove any leading/trailing whitespace
            if matches_timecode_format(line):
                return "NAME_PARENTHESIS_TIMECODE"
            elif matches_timecode_format2(line):
                return "PARENTHESIS_NAME_TIMECODE"

    if mode=="":
        n_sets_of_empty_lines=count_consecutive_empty_lines(script_path,2)
        if n_sets_of_empty_lines>1:
            return "EMPTYLINES_SCENE_SEPARATOR"

    return mode    


current_scene_count=1
scene_separator=getSceneSeparator(script_path)
# Open the file and process each line
with open(script_path, 'r', encoding=encoding) as file:
    for line in file:
        line = line.strip()  # Remove any leading/trailing whitespace
        trimmed_line = line.strip()  # Remove any leading/trailing whitespace
        isEmptyLine=len(trimmed_line)==0
        
        if len(trimmed_line)>0:
            if is_scene_line(line) or (isEmptyLine and wasEmptyLine):
                current_scene_count=current_scene_count+1
                current_scene_id = extract_scene_name(line,scene_separator)
                print("---------------------------------------")
                print(f"Scene Line: {line}")
            else:
                    if current_scene_id!=1:
                        is_speaking=is_character_speaking(trimmed_line)
                        print("IsSpeaking "+str(is_speaking)+" "+trimmed_line)
                        if is_character_speaking(trimmed_line):
                            character_name=extract_character_name(trimmed_line)
                            character_name=filter_character_name(character_name)
                            if not character_name == None:
                                if character_name not in scene_characters_presence:
                                    scene_characters_presence[character_name] = set()
                                scene_characters_presence[character_name].add(current_scene_id)
                                
                                if current_scene_id not in character_scene_presence:
                                    character_scene_presence[current_scene_id] = set()
                                character_scene_presence[current_scene_id].add(character_name)
        wasEmptyLine=isEmptyLine

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


#character_scene_presence=sort_dict_values(character_scene_presence)
scene_characters_presence=sort_dict_values(scene_characters_presence)
write_character_map_to_file(character_scene_presence, output_path+"character_by_scenes.txt")
write_character_map_to_file(scene_characters_presence, output_path+"scenes_by_character.txt")






def matches_format(line):
    pattern = re.compile(r"^\S+\s+\t.*$")
    pattern = re.compile(r"^\S+\s*\t.*$")
    return bool(re.match(pattern, line))

# Test the function with the given line
line = "EMILIA	H... Bon"
result = matches_format(line)
print(f"Does the line match the format? {result}")