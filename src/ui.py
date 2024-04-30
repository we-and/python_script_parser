import os
import tkinter as tk
from tkinter import ttk, filedialog, Text,Menu
from script_parser import process_script
import chardet

outputFolder="tmp"
if not os.path.exists(outputFolder):
    os.mkdir(outputFolder)

def load_tree(parent, root_path):
    # Clear the tree view if root_path is the starting directory
    if parent == "":
        folders.delete(*folders.get_children())
        parent = folders.insert('', 'end', text=os.path.basename(root_path), open=True, values=[root_path])

    # List all entries in the directory
    try:
        dir_entries = os.listdir(root_path)
    except PermissionError:
        return  # Skip directories without permission

    # Separate and sort directories and files
    dirs = sorted([entry for entry in dir_entries if os.path.isdir(os.path.join(root_path, entry))])
    files = sorted([entry for entry in dir_entries if not os.path.isdir(os.path.join(root_path, entry))])

    # Insert directories first
    for entry in dirs:
        entry_path = os.path.join(root_path, entry)
        dir_id = folders.insert(parent, 'end', text=entry, open=False, values=[entry_path])
        load_tree(dir_id, entry_path)  # Recursively load subdirectories

    # Insert files
    for entry in files:
        entry_path = os.path.join(root_path, entry)
        folders.insert(parent, 'end', text=entry, values=[entry_path])

def is_supported_extension(ext):
    ext=ext.lower()
    return ext==".txt"
def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:  # Open the file in binary mode
        raw_data = file.read(10000)  # Read the first 10000 bytes to guess the encoding
        result = chardet.detect(raw_data)
        return result
def get_encoding(enc):
    #print("Guess encoding from"+str(enc))
    if enc=="ascii":
        return "ISO-8859-1"
    elif enc=="ISO-8859-1":
        return "ISO-8859-1"
    elif enc=="Windows-1252":
        return "Windows-1252"       
    return "utf-8"

def reset_tables(): 
    for item in breakdown_table.get_children():
        breakdown_table.delete(item)
    for item in character_table.get_children():
        character_table.delete(item)

def on_folder_select(event):
    print("FOLDER SELECT")
    reset_tables()
    selected_item = folders.selection()[0]
    file_path = folders.item(selected_item, 'values')[0]
    # Check if the selected item is a file and display its content
    if os.path.isfile(file_path):
        try:
            file_name = os.path.basename(file_path)
            name, extension = os.path.splitext(file_name)
            if is_supported_extension(extension):
                print(" > Supported")

                encoding_info = detect_file_encoding(file_path)
                encoding=encoding_info['encoding']
                print("Encoding detected  : "+str(encoding))
                print("Encoding confidence : "+str(encoding_info['confidence']))

                enc=get_encoding(encoding)
                print("Encoding used       : "+str(enc))
    #            encodings = ['windows-1252', 'iso-8859-1', 'utf-16','utf-8']
     #           for encod in encodings:
     #               print("try encoding"+encod)
                with open(file_path, 'r', encoding=enc) as file:
                    content = file.read()

                    file_preview.delete(1.0, tk.END)
                    file_preview.insert(tk.END, content)
                    
                    breakdown,character_scene_map,scene_characters_map,character_linecount_map,character_order_map,character_textlength_map=process_script(file_path,outputFolder+"/"+name+"/",name)
                    fill_breakdown_table(breakdown)
                    fill_character_table(character_order_map, breakdown,character_linecount_map,scene_characters_map)
                    update_statistics(content)
            else:
                print(" > Not supported")
                stats_label.config(text=f"Format {extension} not supported")

        except Exception as e:
            file_preview.delete(1.0, tk.END)
            file_preview.insert(tk.END, f"Error opening file: {e}")

def update_statistics(content):
    words = len(content.split())
    chars = len(content)
    stats_label.config(text=f"Words: {words} Characters: {chars}")
  #  stats_text.insert(0,f"Words: {words} Characters: {chars}")

def open_folder():
    directory = filedialog.askdirectory(initialdir=os.getcwd())
    if directory:
        load_tree(directory)


def exit_app():
    app.quit()


def center_window():
    app.update_idletasks()
    width = app.winfo_width()
    height = app.winfo_height()
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
     # Calculate width and height as 80% of screen dimensions
    width = int(screen_width * 0.8)
    height = int(screen_height * 0.8)
    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 2)
    app.geometry(f'{width}x{height}+{x}+{y}')

def stats_per_character(breakdown,character_name):
    line_count=0
    word_count=0
    character_count=0
    
    for item in breakdown:
        if item['type']=="SPEECH":
            if item['character']==character_name:
                t=item['speech']
                line_count=line_count+1
                character_count=character_count+len(t)
                word_count=word_count+len(t.split(" "))
 
    replica_count=round(character_count/40)
    return line_count,word_count,character_count,replica_count

def fill_character_table(character_order_map, breakdown,character_linecount_map,scene_characters_map):
    for item in character_order_map:
        print(item)
        lines=character_linecount_map[item]
        line_count,word_count,character_count,replica_count=stats_per_character(breakdown,item)
        scenes=scene_characters_map[item]
        scenes=', '.join(scenes)
        character_table.insert('','end',values=(str(character_order_map[item]),item,str(line_count),str(character_count),str(word_count),str(replica_count),scenes))
        

def fill_breakdown_table(breakdown):
    for item in breakdown:
        type_=item['type']
        line_idx=item['line_idx']
        if(type_=="SCENE_SEP"):
            scene_id=item['scene_id']
            breakdown_table.insert('','end',values=(str(line_idx),type_,scene_id,""))
        elif(type_=="SPEECH"):
            speech=item['speech']
            character=item['character']
            breakdown_table.insert('','end',values=(str(line_idx),type_,character,speech))
        elif(type_=="NONSPEECH"):         
            text=item['text']
            breakdown_table.insert('','end',values=(str(line_idx),type_,text,""))
    print("NB ROWS = "+str(len(breakdown_table.get_children())))
    breakdown_table.update_idletasks()


def clear_table(treeview):
    """
    Clears all rows from the given Treeview table.
    
    Args:
        treeview (ttk.Treeview): The Treeview widget instance.
    """
    for item in treeview.get_children():
        treeview.delete(item)

app = tk.Tk()
app.title('Script Analyzer')


# Menu bar
menu_bar = Menu(app)
app.config(menu=menu_bar)


# File menu
file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open Folder...", command=open_folder)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=exit_app)

# Layout configuration
left_frame = ttk.Frame(app)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

right_frame = ttk.Frame(app)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Folder tree
folders = ttk.Treeview(left_frame, columns=())
folders.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
folders.bind('<<TreeviewSelect>>', on_folder_select)

# Notebook (tabbed interface)
notebook = ttk.Notebook(right_frame)
notebook.pack(fill=tk.BOTH, expand=True)

# File preview tab
preview_frame = ttk.Frame(notebook)
notebook.add(preview_frame, text='Preview')
file_preview = Text(preview_frame)
file_preview.pack(fill=tk.BOTH, expand=True)


# Initialize the style object
style = ttk.Style(app)

# Define a new style that inherits from the default Treeview style
# and changes the background color for both selected and normal items.
style.map('Custom.Treeview', background=[('selected', 'blue'), ('!selected', 'green')])


# Statistics tab
stats_frame = ttk.Frame(notebook)

# Create a Treeview widget within the stats_frame for the table
character_table = ttk.Treeview(stats_frame, columns=('Order', 'Character', 'Lines','Character Count','Word Count','Blocks','Scenes'), show='headings')
# Define the column headings
character_table.heading('Order', text='Order')
character_table.heading('Character', text='Character')
character_table.heading('Lines', text='Lines')
character_table.heading('Character Count', text='Character Count')
character_table.heading('Word Count', text='Word Count')
character_table.heading('Blocks', text='Blocks')
character_table.heading('Scenes', text='Scenes')

# Define the column width and alignment
character_table.column('Order', width=25, anchor='center')
character_table.column('Character', width=200, anchor='w')
character_table.column('Lines', width=50, anchor='w')
character_table.column('Character Count', width=50, anchor='w')
character_table.column('Word Count', width=50, anchor='w')
character_table.column('Blocks', width=50, anchor='w')
character_table.column('Scenes', width=50, anchor='w')

# Pack the Treeview widget with enough space
character_table.pack(fill='both', expand=True)
notebook.add(stats_frame, text='Characters')




stats_frame4 = ttk.Frame(notebook)
# Create a Treeview widget within the stats_frame for the table
breakdown_table = ttk.Treeview(stats_frame4, columns=('Line', 'Type', 'Character','Text'), show='headings', style='Custom.Treeview')
# Define the column headings
breakdown_table.heading('Line', text='Line')
breakdown_table.heading('Type', text='Type')
breakdown_table.heading('Character', text='Character')
breakdown_table.heading('Text', text='Text')

# Define the column width and alignment
breakdown_table.column('Line', width=25, anchor='center')
breakdown_table.column('Type', width=50, anchor='center')
breakdown_table.column('Character', width=100, anchor='w')
breakdown_table.column('Text', width=200, anchor='w')
# Pack the Treeview widget with enough space
breakdown_table.pack(fill='both', expand=True)
notebook.add(stats_frame4, text='Découpage')
        
        





stats_frame2 = ttk.Frame(notebook)
notebook.add(stats_frame2, text='Scenes')
stats_frame3 = ttk.Frame(notebook)
notebook.add(stats_frame3, text='Statistics')
stats_text = Text(stats_frame, height=4, state='disabled')
stats_text.pack(fill=tk.BOTH, expand=True)

# Statistics label
stats_label = ttk.Label(right_frame, text="Words: 0 Characters: 0", font=('Arial', 12))
stats_label.pack(side=tk.BOTTOM, fill=tk.X)

# Load folder button
#load_button = ttk.Button(left_frame, text="Open Folder", command=open_folder)
#load_button.pack(side=tk.TOP, fill=tk.X)

load_tree("",os.getcwd())

center_window()  # Center the window

app.mainloop()