import re

def remove_parentheses_contents(input_string):
    # Define the regular expression pattern to match contents within parentheses
    pattern = re.compile(r'\(.*?\)')
    
    # Use re.sub() to replace the matched patterns with an empty string
    result = re.sub(pattern, '', input_string)
    
    # Return the modified string
    return result
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
