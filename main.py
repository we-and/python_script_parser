from PyPDF2 import PdfReader
import re
script_path="scripts/Whiplash.pdf"
txt_path="output_file.txt"
centered_txt_path='output_file_centered.txt'
print("-----------------------------------")
print("SCRIPT PARSER")
print("v1.3 ")
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        print("N Pages"+str(len(reader.pages)))
        i = 1
        for page in reader.pages:
            print("Extracting page "+str(i)+"/"+str(len(reader.pages)))
            text += page.extract_text()
            i=i+1
    return text

pdf_text = extract_text_from_pdf(script_path)
print("PDF extraction done")  
print("Extract : "+pdf_text[1:100])  

uppercase_lines=[]
current_scene_nb=-1
scene_characters_presence={}
character_scene_presence={}


import pdfplumber
def get_scene_number(line):
    """Extracts the first number from a given line, which is assumed to be the scene number."""
    match = re.search(r'\d+', line)
    if match:
        return int(match.group(0))  # Convert the matched string to an integer
    return None 

def pdf_to_text(pdf_path, output_txt_path):
    with pdfplumber.open(pdf_path) as pdf, open(output_txt_path, 'w', encoding='utf-8') as output_file:
        for page in pdf.pages:
            # Extract text from the current page
            page_text = page.extract_text()
            # If page_text is None, it means no text could be extracted from this page
            if page_text:
                # Write the extracted text to the output file
                output_file.write(page_text + '\n')  # Add a newline to separate pages

# Replace 'path_to_your_pdf.pdf' with the path to your PDF file
# Replace 'output_file.txt' with your desired output text file name
pdf_to_text(script_path, txt_path)

def find_scene_lines(text):
    pattern = re.compile(r"^(\d+)\s.*\1\.$", re.MULTILINE)
    matches = pattern.findall(text)
    return matches


def is_scene_line_only_numbers(line):
    pattern = re.compile(r"^(\d+)\s.*\s\1$")
    """Check if the given line matches the scene line format."""
    return bool(re.fullmatch(pattern, line))


def is_scene_line(line):
    # Adjust the regex to match numbers possibly followed by a single letter (e.g., 14A)
    pattern = re.compile(r"^(\d+[A-Z]?)\s.*\s\1$", re.IGNORECASE)
    """Check if the given line matches the scene line format."""
    return bool(re.fullmatch(pattern, line))

def is_enclosed_by_parentheses(line):
    """Check if the given line starts with '(' and ends with ')'."""
    # Regex pattern to match lines that start and end with parentheses
    pattern = re.compile(r"^\(.*\)$")
    return bool(re.fullmatch(pattern, line))
def filter_continued(line):
    return line.replace(" (CONT’D)","").replace(" (CONT'D)","")
def filter_os(line):
    return line.replace(" (O.S.)","")
def filter_vo(line):
    return line.replace(" (V.O.)","")
def filter_prelap(line):
    return line.replace(" (PRE-LAP)","")
def filter_stars(line):
    return line.replace(" **","")

def ends_with_exclamation(line):
    """Check if the given line ends with an exclamation mark."""
    # Regex pattern to match lines that end with '!'
    pattern = re.compile(r".*!$")
    return bool(re.fullmatch(pattern, line))

def is_instruction(line):
    pattern = re.compile(r"CUT TO|FADE OUT", re.IGNORECASE)
    return bool(re.search(pattern, line))

def filter_character_name(line):
    return filter_stars(filter_prelap(filter_vo(filter_os(filter_continued(line)))))

def is_uppercase(line):
    """Check if the given line is entirely in uppercase."""
    return line.isupper()
def is_uppercase(line):
    """Check if the given line is entirely in uppercase."""
    return line.isupper()

def is_character_candidate(line):
    return  is_uppercase(line) and not is_instruction(line) and not is_enclosed_by_parentheses(line) and not ends_with_exclamation(line)
def center_text(text, width):
    """Centers text within a specified width."""
    lines = text.split('\n')
    centered_lines = [line.center(width) for line in lines]
    return '\n'.join(centered_lines)

def write_character_map_to_file(character_map, filename):
    """Writes the character to scene map to a specified file."""
    with open(filename, 'w', encoding='utf-8') as file:
        for character, scenes in character_map.items():
            file.write(f"{character}: {scenes}\n")

def pdf_to_centered_text(pdf_path, output_txt_path, line_width):
    with pdfplumber.open(pdf_path) as pdf, open(output_txt_path, 'w', encoding='utf-8') as output_file:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                # Center the text for each page
                centered_text = center_text(page_text, line_width)
                # Write the centered text to the output file
                output_file.write(centered_text + '\n\n')  # Extra newline to separate pages

# Specify the fixed width for each line (e.g., 80 characters)
line_width = 80
pdf_to_centered_text(script_path, centered_txt_path, line_width)


# Open the file and process each line
with open(centered_txt_path, 'r', encoding='utf-8') as file:
    for line in file:
        line = line.strip()  # Remove any leading/trailing whitespace
        if is_scene_line(line):
            scene_number = get_scene_number(line)
            current_scene_nb=scene_number
            print(f"Scene Line: {line}")
        else:
            if current_scene_nb>-1:
                trimmed_line = line.strip()  # Remove any leading/trailing whitespace
                if is_character_candidate(trimmed_line):
                    uppercase_lines.append(trimmed_line) 
                    character_name=trimmed_line
                    character_name=filter_character_name(character_name)
                    if character_name not in scene_characters_presence:
                        scene_characters_presence[character_name] = set()
                    scene_characters_presence[character_name].add(current_scene_nb)
                    
                    if current_scene_nb not in character_scene_presence:
                        character_scene_presence[current_scene_nb] = set()
                    character_scene_presence[current_scene_nb].add(character_name)

print(scene_characters_presence)
print(character_scene_presence)
write_character_map_to_file(character_scene_presence, "character_by_scenes.txt")
write_character_map_to_file(scene_characters_presence, "scenes_by_character.txt")
import re

def parse_script(text):
    script_elements = {}
    # Example pattern: lines start with an uppercase name followed by dialogue
    pattern = re.compile(r"^\s*([A-Z][A-Z ]+)\s*\n(.*?)\n", re.MULTILINE | re.DOTALL)
    matches = pattern.finditer(text)
    for match in matches:
        name = match.group(1).strip()
        dialogue = match.group(2).strip()
        if name in script_elements:
            script_elements[name].append(dialogue)
        else:
            script_elements[name] = [dialogue]
    return script_elements

script_data = parse_script(pdf_text)
for name, dialogues in script_data.items():
    print(f"{name}:")
    for dialogue in dialogues:
        print(f"  {dialogue}")