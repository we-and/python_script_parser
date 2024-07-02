import re

def count_matches_TIMECODE_HYPHEN_UPPERCASECHARACTER_SPACE_SEMICOLON_DOUBLESPACE_TEXT_NEWLINE_TEXT(file_path, encod):
    # Define the regex pattern  
    pattern = re.compile(r'\d{2}’\d{2}-[A-Z]+ :*', re.DOTALL)
        
    # Read the content of the file
    with open(file_path, 'r', encoding=encod) as file:
        content = file.read()
    
    # Find all matches in the text
    matches = pattern.findall(content)
    
    # Return the count of matches
    return len(matches)

# Example usage:
file_path = 'example.txt'  # Path to your text file
encod = 'utf-8'  # Encoding of the text file

# Test the function
count = count_matches_TIMECODE_HYPHEN_UPPERCASECHARACTER_SPACE_SEMICOLON_DOUBLESPACE_TEXT_NEWLINE_TEXT(file_path, encod)
print(f'Number of matches: {count}')  # Output should be 2