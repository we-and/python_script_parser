import re
import logging
from  constants import action_verbs,characterSeparators,countMethods,multilineCharacterSeparators
from utils_filters import is_didascalie,is_ambiance,is_music,filter_character_name
from utils_regex import is_TIMECODE_HYPHEN_TIMECODE,is_TIMECODE_SPACE_TIMECODE_SPACE_DIALOG
def myprint1(s):
    logging.debug(s)
    #myprint1(s)

def count_NUM_TIMECODE_ARROW_TIMECODE_NEWLINE_MULTILINEDIALOG(file_path, encoding):
    # Define the regex pattern
    pattern = re.compile(r'\d+ \d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}\n.+(\n.+)?', re.MULTILINE)
    
    # Read the content of the file
    with open(file_path, 'r', encoding=encoding) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)
def count_NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME(file_path, encoding):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}: +\d{2}:\d{2}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}:\d{2} \d{2}:\d{2}\n.+(\n.+)?', re.MULTILINE)
    
    # Read the content of the file
    with open(file_path, 'r', encoding=encoding) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)

def count_TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_NEWLINE_DIALOG(file_path, encoding):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2} - \d{2}:\d{2}:\d{2}:\d{2}\n\w+\n.*', re.DOTALL)
    
    # Read the content of the file
    with open(file_path, 'r', encoding=encoding) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)


def extract_character_name_TIMECODE_HYPHEN_UPPERCASECHARACTER_SPACE_SEMICOLON_DOUBLESPACE_TEXT_NEWLINE_TEXT(text):
    pattern = re.compile(r'\d{2}’\d{2}-([A-Z]+) :', re.DOTALL)
    
    match = pattern.match(text)
    if match:
        return match.group(1)
    return None
def is_characterline_TIMECODE_HYPHEN_UPPERCASECHARACTER_SPACE_SEMICOLON_DOUBLESPACE_TEXT_NEWLINE_TEXT(text):
    pattern = re.compile(r'\d{2}’\d{2}-[A-Z]+ :*', re.DOTALL)
    return bool(pattern.fullmatch(text))
def count_matches_TIMECODE_HYPHEN_UPPERCASECHARACTER_SPACE_SEMICOLON_DOUBLESPACE_TEXT_NEWLINE_TEXT(file_path,encod):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}’\d{2}-[A-Z]+ :*', re.DOTALL)
    
     # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)
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
def matches_scenestart_sceneno(line):
    return line.startswith("SCENE NO") or line.startswith("CENE NO") or line.startswith("ESCENA NO")
 
def is_scene_line(line):
    isSceneLine=matches_format_parenthesis_name_timecode(line) or matches_number_parenthesis_timecode(line) or matches_scenestart_sceneno(line)
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
    elif scene_separator=="SCENENO_INTEXT_LOCATION":
        return line
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

def count_lines_in_file(script_path,encod):
    nLines=0
    with open(script_path, 'r', encoding=encod) as file:
        for line in file:
            line = line.strip()
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

def count_matches_TIMECODE_NEWLINE_CHARACTERINBRACKETS_DIALOG_NEWLINE_NEWLINE(file_path,encod):
    """
    Counts the number of matches for the specified pattern in a file.

    The pattern matches the following sequence:
    - A timestamp in the format [HH:MM:SS.SS]
    - Followed by a newline
    - Followed by a character name in square brackets and some dialog
    - Followed by a newline
    - Followed by another timestamp in the format [HH:MM:SS.SS]

    :param file_path: The path to the file to be read.
    :param encod: The encoding of the file to be read.
    :return: The number of matches found in the file.
    """
    
    # Define the pattern to match
    pattern = re.compile(
        r'\[\d{2}:\d{2}:\d{2}\.\d{2}\]\n'
        r'\[\w+\].*\n'
        r'\[\d{2}:\d{2}:\d{2}\.\d{2}\]'
    )
    
    # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Return the number of matches
    return len(matches)
def extract_TIMECODE_ARROW_TIMECODE_NEWLINE_CHARACTER_SEMICOLON_DIALOG_NEWLINE_DIALOG(file_path,encod):
    # Define the regular expression pattern
    pattern = re.compile(
        r'(\d+)\n'  # Number followed by newline (captured as group 1)
        r'(\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3})\n'  # Two timecodes separated by arrow and newline (captured as group 2)
        r'(.*?): (.*?)\n',  # Text line with : and <i></i> tags (captured as group 3 and group 4)
        re.MULTILINE
    )

    # Read the content of the file
    with open(file_path, 'r', encoding=encod) as file:
        content = file.read()

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Extract the required fields into a list of dictionaries
    extracted_data = [
        {
            "text": match[2].strip(),
            "dialog": match[3].strip()
        }
        for match in matches
    ]

    return extracted_data

def count_matches_LINE_NEWLINE_TIMECODE_ARROW_TIMECODE_NEWLINE_TEXT_ITAG(file_path,encod):
    """
    Counts the number of matches for the specified pattern in a file.

    The pattern matches the following sequence:
    - A line containing a number followed by a newline
    - Followed by two timecodes in the format HH:MM:SS,SSS separated by " --> " and followed by a newline
    - Followed by a text line that contains a colon and may include <i></i> tags, followed by a newline

    :param file_path: The path to the file to be read.
    :param encod: The encoding of the file to be read.
    :return: The number of matches found in the file.
    """
    # Define the regular expression pattern
    pattern = re.compile(
        r'\d+\n'  # Number followed by newline
        r'\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}\n'  # Two timecodes separated by arrow and newline
        r'.*?: .*?\n',  # Text line with : and <i></i> tags
        re.MULTILINE
    )
    
    # Read the content of the file
    with open(file_path, 'r', encoding=encod) as file:
        content = file.read()

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Return the number of matches
    return len(matches)
def count_matches_TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_SEMICOLON_NEWLINE_DIALOGNEWLINE(file_path,encod):
    """
    Counts the occurrences of the specified pattern:
    1. First line: two timecodes separated by " - "
    2. Second line: text with ":"
    3. Rest of the lines: free text
    
    :param text: The text to check.
    :return: The number of occurrences of the pattern in the text.
    """
    # Define the regex pattern for matching the whole block
    pattern = (
        r'^\d{2}:\d{2}:\d{2}:\d{2} - \d{2}:\d{2}:\d{2}:\d{2}\n'  # First line: timecodes
        r'.+:\n'  # Second line: text with ":"
        r'(.|\n)*?'  # Rest of the lines: free text, non-greedy to match one block at a time
        r'(?=\d{2}:\d{2}:\d{2}:\d{2} - |\Z)'  # Lookahead to ensure we're at the end of a block or the end of text
    )
    # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()

    # Find all matches of the pattern
    matches = re.findall(pattern, content, re.MULTILINE)
    
    # Return the number of matches
    return len(matches)


def count_matches_NAME_NEWLINE_DIALOG_NEWLINE_NEWLINE(file_path,encod):
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
def extract_text_after_brackets(text):
    """
    Extracts the text immediately following the first pair of square brackets in the given string.
    
    :param text: The string to extract from.
    :return: The text immediately following the first pair of square brackets, or None if not found.
    """
    pattern = r'\[.*?\]\s*(.*)'
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip()
    return None
def extract_text_between_brackets(text):
    """
    Extracts the text between the first pair of square brackets in the given string.
    
    :param text: The string to extract from.
    :return: The text between the first pair of square brackets, or None if not found.
    """
    pattern = r'\[(.*?)\]'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None
def is_text_with_brackets_pattern(text):
    """
    Checks if a given string matches the pattern "[Text] more text".
    
    :param text: The string to check.
    :return: True if the string matches the pattern, False otherwise.
    """
    pattern = r'^\[.*?\] .*$'
    match = re.match(pattern, text)
    return bool(match)

def count_matches_TIMECODE_ARROW_TIMECODE_NEWLINE_BRACKETS_CHARACTER_DIALOG_NEWLINE_DIALOG(file_path,encod):
    """
    Counts the occurrences of the pattern:
    00:01:06,691 --> 00:01:08,943
    [Text] Text
    More Text
    
    :param text: The text to check.
    :return: The number of occurrences of the pattern in the text.
    """
    pattern = (
        r'\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}\n'  # Timestamps
        r'\[.*?\] .*?\n'  # Text in square brackets followed by more text
        r'.*?'  # Optional more text on the third line
    )
    # Read the content of the file
    with open(file_path, 'r',encoding=encod) as file:
        content = file.read()

    # Use re.findall to find all occurrences of the pattern
    matches = re.findall(pattern, content, re.MULTILINE)
    
    # Return the number of matches
    return len(matches)
def detectCharacterSeparator(script_path,encod):
    myprint1("detectCharacterSeparator")
    singlebest="?"
    singlebestMatchPercent=0.0
    nLines=count_nonempty_lines_in_file(script_path,encod)
    if nLines==0:
        return None
    myprint1(f"detectCharacterSeparator nlines={nLines}")
    myprint1(f"--------------")
    myprint1(f" > Test single line "+str(len(characterSeparators)))
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
                    if sep=="CHARACTER_SEMICOL_DIALOG":
                        is_match=matches_charactername_NAME_SEMICOLON_DIALOG(line)
                        if is_match:
                            nMatches=nMatches+1
                    if sep=="CHARACTERUPPERCASE_DIALOG":
                        is_match=matches_CHARACTERUPPERCASE_DIALOG(line)
                        # myprint1(f"line {line} {is_match}")
                        if is_match:
                            nMatches=nMatches+1
                    if sep=="CHARACTER_SPACES":
                        is_match=matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)
                        if is_match:
                            nMatches=nMatches+1
                    if sep=="CHARACTER_TAB":
                        is_match=matches_charactername_NAME_ATLEAST1TAB_TEXT(line)
                        if is_match:
                            nMatches=nMatches+1

                    if sep=="TIMECODE_SPACE_TIMECODE_SPACE_DIALOG":
                        is_match=is_TIMECODE_SPACE_TIMECODE_SPACE_DIALOG(line)
                        if is_match:
                            nMatches=nMatches+1
        pc=(nMatches/nLines)
        myprint1("  > Test sep="+sep+" matches=" +str(nMatches)+"/"+str(nLines)+" pc="+str(pc))
        if pc>=singlebestMatchPercent:
            singlebestMatchPercent=pc
            singlebest=sep

    multibestMatches=-1
    myprint1(f"--------------")
    myprint1(" > Test Multiline  "+str(len(multilineCharacterSeparators)))
    multibest="?"
    for sep in multilineCharacterSeparators:
        if sep == "CHARACTER_NEWLINE_DIALOG_NEWLINE_NEWLINE":
            nMatches=count_matches_NAME_NEWLINE_DIALOG_NEWLINE_NEWLINE(script_path,encod)
        elif sep=="TIMECODE_ARROW_TIMECODE_NEWLINE_BRACKETS_CHARACTER_DIALOG_NEWLINE_DIALOG":
            nMatches=count_matches_TIMECODE_ARROW_TIMECODE_NEWLINE_BRACKETS_CHARACTER_DIALOG_NEWLINE_DIALOG(script_path,encod)
        elif sep=="NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME":
            nMatches=count_NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME(script_path,encod)
        elif sep=="NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME":
            nMatches=count_NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME(script_path,encod)
        elif sep=="NUM_TIMECODE_ARROW_TIMECODE_NEWLINE_MULTILINEDIALOG":
            nMatches=count_NUM_TIMECODE_ARROW_TIMECODE_NEWLINE_MULTILINEDIALOG(script_path,encod)
        elif sep=="TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_NEWLINE_DIALOG":
            nMatches=count_TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_NEWLINE_DIALOG(script_path,encod)
        elif sep=="TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_SEMICOLON_NEWLINE_DIALOG_NEWLINE":
            nMatches=count_matches_TIMECODE_HYPHEN_TIMECODE_NEWLINE_CHARACTER_SEMICOLON_NEWLINE_DIALOGNEWLINE(script_path,encod)
        elif sep=="TIMECODE_NEWLINE_CHARACTERINBRACKETS_DIALOG_NEWLINE_NEWLINE":
            nMatches=count_matches_TIMECODE_NEWLINE_CHARACTERINBRACKETS_DIALOG_NEWLINE_NEWLINE(script_path,encod)
        elif sep=="LINE_NEWLINE_TIMECODE_ARROW_TIMECODE_NEWLINE_TEXT_ITAG":
            nMatches=count_matches_LINE_NEWLINE_TIMECODE_ARROW_TIMECODE_NEWLINE_TEXT_ITAG(script_path,encod)
        myprint1("  > Test character sep:"+sep+" " +str(nMatches))
        if nMatches>multibestMatches:
            multibestMatches=nMatches
            multibest=sep
    myprint1(f"detectCharacterSeparator multibestMatches= {multibestMatches}")


    myprint1(f"detectCharacterSeparator singlebest= {singlebest} pc={singlebestMatchPercent} ")
    myprint1(f"detectCharacterSeparator multibest= {multibest}")
    if singlebestMatchPercent<0.1:
        best= multibest
    else:
        best= singlebest
    myprint1(f"Best {best}")
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
            elif matches_scenestart_sceneno(line):
                return "SCENENO_INTEXT_LOCATION"
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
    elif character_mode=="CHARACTERUPPERCASE_DIALOG": 
        return matches_CHARACTERUPPERCASE_DIALOG(line)
    elif character_mode=="CHARACTER_SPACES": 
        return matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return matches_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_DIALOG": 
        return matches_charactername_NAME_SEMICOLON_DIALOG(line)  
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
        if name==None:
            return False
        if character_mode=="CHARACTERUPPERCASE_DIALOG":
            return not is_didascalie(name) and not is_ambiance(name) and not is_music(name)
        else:
            return not is_didascalie(name) and not is_ambiance(name) 
    else:
        return False

def extract_dialog_TIMECODE_SPACE_TIMECODE_SPACE_DIALOG(content):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2}\s+(?:--:--:--:--|\d{2}:\d{2}:\d{2}:\d{2})\s+(.+)')

    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the list of extracted dialogs
    return  ' '.join(matches)

def extract_speech(line,character_mode,character_name):
    if character_mode=="TIMECODE_SPACE_TIMECODE_SPACE_DIALOG":
        return extract_dialog_TIMECODE_SPACE_TIMECODE_SPACE_DIALOG(line)
    if character_mode=="CHARACTER_TAB":
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTERUPPERCASE_DIALOG":
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTER_SPACES": 
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTER_SEMICOL_DIALOG": 
        return line.replace(character_name,"").strip()
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return extract_speech_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line,character_name)  
    else:
        myprint1("ERROR wrong mode1="+str(character_mode))
        exit()

def extract_character_name(line,character_mode):
    if character_mode=="CHARACTER_TAB":
        return extract_charactername_NAME_ATLEAST1TAB_TEXT(line)
    elif character_mode=="TIMECODE_SPACE_TIMECODE_SPACE_DIALOG":
        return "PERSONNAGE"
    elif character_mode=="CHARACTERUPPERCASE_DIALOG":
        return extract_charactername_CHARACTERUPPERCASE_DIALOG(line)
    elif character_mode=="CHARACTER_SPACES": 
        myprint1("debug")       
        is_match=matches_charactername_NAME_ATLEAST8SPACES_TEXT(line)
        myprint1("is_match?"+str(is_match))
        myprint1("extract "+str(character_mode)+" "+str(line))
        return extract_charactername_NAME_ATLEAST8SPACES_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_TAB": 
        return extract_charactername_NAME_SEMICOLON_OPTSPACES_TAB_TEXT(line)  
    elif character_mode=="CHARACTER_SEMICOL_DIALOG": 
        return extract_charactername_NAME_SEMICOLON_DIALOG(line)  
    else:
        myprint1("ERROR wrong mode="+str(character_mode))
        exit()
def extract_charactername_NAME_SEMICOLON_DIALOG(text):
    """
    Extracts the first uppercase part before the first colon in the given text.
    
    :param text: The text to extract from.
    :return: The first uppercase part before the first colon, or None if not found.
    """
    pattern = r'^([A-Za-z -]+):'
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip()
    return None
def matches_charactername_NAME_SEMICOLON_DIALOG(text):
    # Define the regex pattern
    # ^ starts the match at the beginning of the line
    # [\w\s]+ matches one or more word characters or spaces to include names with spaces
    # : matches the literal colon
    # \s* matches zero or more whitespace characters (spaces or tabs)
    # \t matches a tab
    # .+ matches one or more of any character (the text following the tab)
    # $ ensures the pattern goes to the end of the line
    pattern = r'^[A-Za-z -]+:\s+.*'

    # Use re.match to check if the start of the string matches the pattern
    if re.match(pattern, text):
        return True
    else:
        return False
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
def matches_CHARACTERUPPERCASE_DIALOG(s):
    """
    Tests if the given string follows the format "TEXT IN UPPERCASE Then text normal".

    :param s: The string to test.
    :return: True if the string matches the format, False otherwise.
    """
    # Define the regular expression pattern
    pattern = r'^[A-Z\s]+[a-z].*$'
    pattern = r'^([A-Z\'\s)(\\)\.-]+)\s(?=[A-Z]|[a-z])'

    # Use the re.match function to check if the string matches the pattern
    match = re.match(pattern, s)

    # Return True if there's a match, False otherwise
    return bool(match)
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

def ensure_dialog_starts_with_uppercase(character, dialog):
    """
    Ensures that the dialog starts with an uppercase letter.
    If not, it prepends the last word of the character name to the dialog.

    :param character: The character name.
    :param dialog: The dialog text.
    :return: The modified character (without its last word) and the modified dialog.
    """
    # Split the character name into words
    words = character.split()
    
    # Check if the dialog starts with an uppercase letter
    if not re.match(r'^[A-Z]', dialog):
        if len(character.split(" "))>1:
            # Get the last word of the character name
            last_word = words[-1]
            # Prepend the last word to the dialog
            dialog = f"{last_word} {dialog}"
            # Remove the last word from the character name
            character = ' '.join(words[:-1])
    return character, dialog

def extract_charactername_CHARACTERUPPERCASE_DIALOG_regex(s):
    """
    Extracts the first part in uppercase from the given string.

    :param s: The string to extract from.
    :return: The uppercase part of the string, or None if not found.
    """
    # Define the regular expression pattern
    #pattern = r'^([A-Z\s]+)(?=\s[A-Z]*[a-z])'
    #pattern = r'^([A-Z\s]+)\s(?=[A-Z]|[a-z])'
    pattern = r'^([A-Z\'\s)(\\)\.]+)\s(?=[A-Z]|[a-z])'
    pattern = r'^([A-Z\'\s)(\\)\.-]+)\s(?=[A-Z]|[a-z])'
  
    # Use the re.match function to find the uppercase part
    match = re.match(pattern, s)

    # Return the uppercase part if found, otherwise None
    if match:
        return match.group(1).strip()
    return None

def extract_charactername_CHARACTERUPPERCASE_DIALOG(s):
    """
    Extracts the first part in uppercase from the given string.

    :param s: The string to extract from.
    :return: The uppercase part of the string, or None if not found.
    """
    words = s.split()
    result = []
    
    for word in words:
        if word.isupper() and len(word) > 1:
            result.append(word)
        else:
            break  # Stop if a non-uppercase or single-letter word is found
    
    if result:
        return ' '.join(result)
    return None

def extract_charactername_NAME_ATLEAST1TAB_TEXT(line):
    """Extracts the character name from a line where the name is followed by a tab and then dialogue."""
    # Split the line at the first tab character
    parts = line.split('\t', 1)  # The '1' limits the split to the first occurrence of '\t'
    if len(parts) > 1:
        return parts[0].strip()  # Return the first part, stripping any extra whitespace
    return None  # Return None if no tab is found, indicating an improperly formatted line
