import re

def extract_text_not_in_brackets(line):
    # Define the pattern to match text in brackets
    pattern = re.compile(r'\[.*?\]')
    
    # Replace the text in brackets with an empty string
    result = pattern.sub('', line)
    
    # Strip any leading or trailing whitespace
    return result.strip()


def extract_pattern_occurrences2(content):
    # Define the regular expression pattern
    pattern = re.compile(
    r'^(.*?): (.*)$',     re.MULTILINE
    )
    

    # Find all occurrences of the pattern
    matches = pattern.findall(content)

    # Extract the required fields into a list of dictionaries
    extracted_data = [
        {
            "character": match[0].strip(),
            "dialog": match[1].strip()
        }
        for match in matches
    ]
    
    return extracted_data


t="[gunshot] asa sds"
print(extract_text_not_in_brackets(t))
t="[gunshot] [loud inhale] a"
print(extract_text_not_in_brackets(t))
t="Elias: testsetse"
print(extract_pattern_occurrences2(t))