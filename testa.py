import re

def extract_uppercase_part(s):
    """
    Extracts the first part in uppercase from the given string.
    
    :param s: The string to extract from.
    :return: The uppercase part of the string, or None if not found.
    """
    # Define the regular expression pattern
    pattern = r'^([A-Z\'\s)(\\)\.]+)\s(?=[A-Z]|[a-z])'
    pattern = r'^([A-Z\'\s)(\\)\.-]+)\s(?=[A-Z]|[a-z])'
    # Use the re.match function to find the uppercase part
    match = re.match(pattern, s)

    # Return the uppercase part if found, otherwise None
    if match:
        return match.group(1).strip()
    return None

# Example usage
test_strings = [
    "TEXT IN UPPERCASE Then text normal",
    "BARBARA - They'll say that in your hour of need you turn to God.",
    "UPPERCASE lowercase",
  "LARA  All I ever wanted was a family,",
  "LARA (O.S) All I ever wanted was a family,"
    "DR WALSH I'm not questioning your faith, I'm just trying to facilitate what's best for the both of you. ",
    "ALL UPPERCASE",
    "DR WALSH - How did it get worse?",
    "RADIO ANNOUNCER I into you. You into I.",
    "then lowercase",
    "TEXT IN UPPERCASE THEN MORE UPPERCASE Then text normal",
    "MY FIRST PART My second part",
    "MY UPPERCASE Start with Uppercase then lowercase"
]

for test_string in test_strings:
    print(f"'{test_string}': '{extract_uppercase_part(test_string)}'")