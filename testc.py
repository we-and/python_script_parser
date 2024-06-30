import re

def matches_CHARACTERUPPERCASE_DIALOG(s):
    """
    Tests if the given string follows the format "TEXT IN UPPERCASE Then text normal".

    :param s: The string to test.
    :return: True if the string matches the format, False otherwise.
    """
    # Define the regular expression pattern
    pattern = r'^([A-Z\'\s)(\\)\.-]+)\s(?=[A-Z]|[a-z])'
  
    # Use the re.match function to check if the string matches the pattern
    match = re.match(pattern, s)

    # Return True if there's a match, False otherwise
    return bool(match)

# Example usage
test_strings = [
    "Son, Regina, Mama: ",
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
    print(f"'{test_string}': '{matches_CHARACTERUPPERCASE_DIALOG(test_string)}'")