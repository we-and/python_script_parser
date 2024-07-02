
from utils_regex import remove_parentheses_contents

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
        if line.endswith(')'):
            line= line[:-1]
        if line.startswith('-'):
            line= line[1:]
        if line.endswith('-'):
            line= line[:-1]
        if "(O.S)" in line:
            line=line.replace("(O.S)","")
        if "(V.O)" in line:
            line=line.replace("(V.O)","")
        if "(V.O.)" in line:
            line=line.replace("(V.O.)","")
        if "(VO)" in line:
            line=line.replace("(VO)","")
        if "(V.O" in line:
            line=line.replace("(V.O","")
        
        if "(O.S.)" in line:
            line=line.replace("(O.S.)","")
        if "(OS)" in line:
            line=line.replace("(OS)","")
        if "(OS/ON)" in line:
            line=line.replace("(OS/ON)","")
        
        if "(CON'T)" in line:
            line=line.replace("(CON'T)","")    
        if "(CONT.)" in line:
            line=line.replace("(CONT.)","")    
        if "(CONT." in line:
            line=line.replace("(CONT.","")    
        if "(CON’T)" in line:
            line=line.replace("(CON’T)","")    
        if "(CON’T)" in line:
            line=line.replace("(CON’T)","")    
        if "(CONT'D)" in line:
            line=line.replace("(CONT'D)","")    

        if " CONT'D" in line:
            line=line.replace(" CONT'D","")    
        if "(CONT.)" in line:
            line=line.replace("(CONT.)","")    
        if "(CONT)" in line:
            line=line.replace("(CONT)","")    
        if "(CONT'D" in line:
            line=line.replace("(CONT'D","")    
        if "(CONT’D" in line:
            line=line.replace("(CONT’D","")    
        line=line.replace("\ufeff","")

        line=remove_parentheses_contents(line)
        if line.endswith(')'):
            line= line[:-1]
    return line.strip()
#    return line
