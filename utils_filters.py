
from utils_regex import remove_parentheses_contents
from constants import filterToReplace

#################################################################
# CHARACTER UTILS
def is_didascalie(name):
    return name=="DIDASCALIES"
def is_music(name):
    return name.startswith("♪")
def is_ambiance(name):
    return name=="AMBIANCE"
def filter_character_name(line):

    if line==None:
        return "ERROR CHAR???"
    if line:
        if line.endswith(':'):
            line= line[:-1]
        if line.endswith(','):
            line= line[:-1]
        if line.startswith('-'):
            line= line[1:]
        if line.endswith('-'):
            line= line[:-1]

        for pattern in filterToReplace:
            line = line.replace(pattern, "")
        line=line.replace("\ufeff","")
        if line.endswith(')'):
            line= line[:-1]

        line=remove_parentheses_contents(line)
        if line.endswith(')'):
            line= line[:-1]
    return line.strip()
#    return line
