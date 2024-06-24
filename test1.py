import re

def extract_text_not_in_brackets(line):
    # Define the pattern to match text in brackets
    pattern = re.compile(r'\[.*?\]')
    
    # Replace the text in brackets with an empty string
    result = pattern.sub('', line)
    
    # Strip any leading or trailing whitespace
    return result.strip()
t="[gunshot] asa sds"
print(extract_text_not_in_brackets(t))
t="[gunshot] [loud inhale] a"
print(extract_text_not_in_brackets(t))