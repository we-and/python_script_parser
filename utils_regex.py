import re

def remove_parentheses_contents(input_string):
    # Define the regular expression pattern to match contents within parentheses
    pattern = re.compile(r'\(.*?\)')
    
    # Use re.sub() to replace the matched patterns with an empty string
    result = re.sub(pattern, '', input_string)
    
    # Return the modified string
    return result

def detect_NUM_SPACE_CHARACTERUPPERCASE_SEMICOLON_DIALOG(text):
    """
    Detects occurrences of the pattern "38 CHARACTER: Free text here" in the input text.
    
    Parameters:
        text (str): The input text to search.
        
    Returns:
        list: A list of matched patterns.
    """
    pattern = r'\d{2} [A-Z]+: [^:]+'
    matches = re.findall(pattern, text)
    return matches
def detect_celllayout_CHARACTERUPPERCASE_NEWLINE_DIALOG(text):
    """
    Detects occurrences of the pattern "UPPERCASE TEXT\nNormal text" in the input text.
    
    Parameters:
        text (str): The input text to search.
        
    Returns:
        list: A list of matched patterns.
    """
    pattern = r'([A-Z ]+)\n([^A-Z\n]+)'
    matches = re.findall(pattern, text)
    return matches

def is_celllayout_CHARACTERUPPERCASE_NEWLINE_DIALOG(text):
    """
    Detects occurrences of the pattern "UPPERCASE TEXT\nNormal text" in the input text.
    
    Parameters:
        text (str): The input text to search.
        
    Returns:
        list: A list of matched patterns.
    """
    pattern = r'([A-Z ]+)\n([^A-Z\n]+)'
    matches = re.findall(pattern, text)
    return len(matches)>0
def extract_character_NUM_SPACE_CHARACTERUPPERCASE_SEMICOLON_DIALOG(matched_pattern):
    """
    Extracts the uppercase text from the matched pattern.
    
    Parameters:
        matched_pattern (str): The matched pattern string.
        
    Returns:
        str: The extracted uppercase text.
    """
    pattern = r'\d+ ([A-Z]+):'
    match = re.search(pattern, matched_pattern)
    if match:
        return match.group(1)
    return None

def extract_dialog_NUM_SPACE_CHARACTERUPPERCASE_SEMICOLON_DIALOG(matched_pattern):
    """
    Extracts the free text after the colon from the matched pattern.
    
    Parameters:
        matched_pattern (str): The matched pattern string.
        
    Returns:
        str: The extracted free text.
    """
    pattern = r'\d+ [A-Z]+: (.+)'
    match = re.search(pattern, matched_pattern)
    if match:
        return match.group(1)
    return None
def is_celllayout_NUM_SPACE_CHARACTERUPPERCASE_SEMICOLON_DIALOG(text):
    """
    Detects occurrences of the pattern "38 CHARACTER: Free text here" in the input text.
    
    Parameters:
        text (str): The input text to search.
        
    Returns:
        list: A list of matched patterns.
    """
    pattern = r'\d+ [A-Z]+: [^:]+'
    matches = re.findall(pattern, text)
    return len(matches)>0
def is_TIMECODE_SPACE_TIMECODE_SPACE_DIALOG(line):
    # Define the regex pattern
    pattern = re.compile(r'(\d{2}:\d{2}:\d{2}:\d{2}\s+--:--:--:--\s+.+)|(\d{2}:\d{2}:\d{2}:\d{2}\s+\d{2}:\d{2}:\d{2}:\d{2}\s+.+)')
    
    # Check if the line matches the pattern
    return bool(pattern.fullmatch(line))
def is_NUM_TIMECODE_ARROW_TIMECODE(line):
    # Define the regex pattern
    pattern = re.compile(r'\d+ \d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}')
    
    # Check if the line matches the pattern
    return bool(pattern.fullmatch(line))

def is_NUM_SEMICOLON_TIMECODE_SPACE_TIMECODE_SPACE_TIME(line):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}: {6}\d{2}:\d{2}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}:\d{2} \d{2}:\d{2}')
    
    # Check if the line matches the pattern
    return bool(pattern.fullmatch(line))
def is_TIMECODE_HYPHEN_TIMECODE(content):
    # Define the regex pattern
    pattern = re.compile(r'\d{2}:\d{2}:\d{2}:\d{2} - \d{2}:\d{2}:\d{2}:\d{2}', re.DOTALL)
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)>0

def is_timecode_arrow_timecode_format(text):
    """
    Detects if the given text matches the timecode format "HH:MM:SS:FF --> HH:MM:SS:FF".
    
    :param text: The text to check.
    :return: True if the text matches the timecode format, False otherwise.
    """
    pattern = r'^\d{2}:\d{2}:\d{2}:\d{2} -->\r?\n\d{2}:\d{2}:\d{2}:\d{2}$'
    match = re.match(pattern, text)
    return bool(match)

def is_timecode_simple(text):
    """
    Detects if the given text matches the timecode format "HH:MM:SS:FF --> HH:MM:SS:FF".
    
    :param text: The text to check.
    :return: True if the text matches the timecode format, False otherwise.
    """
    pattern = r'^\d{2}:\d{2}:\d{2}:\d{2}$'
    match = re.match(pattern, text)
    return bool(match)
