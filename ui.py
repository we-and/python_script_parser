import os
import tkinter as tk
from tkinter import ttk, filedialog, Text,Menu

def load_tree(root_path):
    # Clear the tree view
    folders.delete(*folders.get_children())

    # Populate the tree with files and directories
    for root, dirs, files in os.walk(root_path):
        root_id = folders.insert('', 'end', text=os.path.basename(root), open=True, values=[root])
        for d in dirs:
            folders.insert(root_id, 'end', text=d, open=False, values=[os.path.join(root, d)])
        for f in files:
            file_base, file_ext = os.path.splitext(f)
            file_ext = file_ext.lstrip('.')  # Remove the dot from the extension
            if file_ext=="txt":
                folders.insert(root_id, 'end', text=f, values=[os.path.join(root, f)])

def on_folder_select(event):
    selected_item = folders.selection()[0]
    file_path = folders.item(selected_item, 'values')[0]
    
    # Check if the selected item is a file and display its content
    if os.path.isfile(file_path):
        try:
            encoding="ISO-8859-1"

            with open(file_path, 'r', encoding=encoding) as file:
                content = file.read()
            file_preview.delete(1.0, tk.END)
            file_preview.insert(tk.END, content)
            update_statistics(content)
        except Exception as e:
            file_preview.delete(1.0, tk.END)
            file_preview.insert(tk.END, f"Error opening file: {e}")

def update_statistics(content):
    words = len(content.split())
    chars = len(content)
    stats_label.config(text=f"Words: {words} Characters: {chars}")
    stats_text.config(text=f"Words: {words} Characters: {chars}")

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

# Statistics tab
stats_frame = ttk.Frame(notebook)
notebook.add(stats_frame, text='Characters')
stats_text = Text(stats_frame, height=4, state='disabled')
stats_text.pack(fill=tk.BOTH, expand=True)

# Statistics label
stats_label = ttk.Label(right_frame, text="Words: 0 Characters: 0", font=('Arial', 12))
stats_label.pack(side=tk.BOTTOM, fill=tk.X)

# Load folder button
#load_button = ttk.Button(left_frame, text="Open Folder", command=open_folder)
#load_button.pack(side=tk.TOP, fill=tk.X)

load_tree(os.getcwd())

center_window()  # Center the window

app.mainloop()