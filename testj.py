import re

def detect_pattern(text):
    """
    Detects occurrences of the pattern "2 JETT: How we doing on time, Sky?" in the input text.
    
    Parameters:
        text (str): The input text to search.
        
    Returns:
        list: A list of matched patterns.
    """
    pattern = r'\d+ [A-Z]+: [^:]+'
    matches = re.findall(pattern, text)
    return matches

def extract_uppercase_text(matched_pattern):
    """
    Extracts the uppercase text from the matched pattern.
    
    Parameters:
        matched_pattern (str): The matched pattern string.
        
    Returns:
        str: The extracted uppercase text.
    """
    pattern = r'\d{1,2} ([A-Z]+):'
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
    pattern = r'\d+ [A-Z]+: (.+)'
    matches = re.findall(pattern, text)
    return len(matches)>0

def extract_free_text(matched_pattern):
    """
    Extracts the free text after the colon from the matched pattern.
    
    Parameters:
        matched_pattern (str): The matched pattern string.
        
    Returns:
        str: The extracted free text.
    """
    pattern = r'\d{1,2,3,4} [A-Z]+: (.+)'
    match = re.search(pattern, matched_pattern)
    if match:
        return match.group(1)
    return None

# Example usage
text = "2 JETT: How we doing on time, Sky?"
iss=is_celllayout_NUM_SPACE_CHARACTERUPPERCASE_SEMICOLON_DIALOG(text)
print("Is Patterns:", str(iss))

matched_patterns = detect_pattern(text)
print("Matched Patterns:", matched_patterns)

for pattern in matched_patterns:
    uppercase_text = extract_uppercase_text(pattern)
    free_text = extract_free_text(pattern)
    print(f"Pattern: {pattern}")
    print(f"Uppercase Text: {uppercase_text}")
    print(f"Free Text: {free_text}")