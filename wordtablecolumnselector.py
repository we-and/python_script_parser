class WordTableColumnSelector(tk.Toplevel):
    def __init__(self, parent, file_path):
        myprint7("TableColumnSelector init")
        self.parent = parent
        self.table_list = []
        self.doc = None
        self.file_path=file_path
        self.check_vars = []
        self.create_widgets()
        self.doc = Document(file_path)
        self.table_list = [table for table in self.doc.tables]
        self.update_table_listbox()
        myprint7("TableColumnSelector tablecount = "+str(len(self.table_list)))
        
    def reset(self,file_path):
        myprint7("TableColumnSelector RESET")  
        self.table_list = []
        self.doc = None
        self.check_vars = []
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        self.check_vars = []

    def create_widgets(self):
        print("TableColumnSelector create_widgets")
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Frame for table preview
        self.table_frame = tk.Frame(self.parent, borderwidth=0)
        self.table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame for table list
        self.menu_frame = tk.Frame(self.parent, width=200)
        self.menu_frame.pack(side=tk.RIGHT, fill='y')
        self.menu_frame.pack_propagate(False) 
        self.left_canvas = tk.Canvas(self.menu_frame, borderwidth=0)
        self.left_canvas.pack(side=tk.TOP, fill='both', expand=True)

        self.list_frame = tk.Frame(self.left_canvas)
        self.left_canvas.create_window((0, 0), window=self.list_frame, anchor='nw')

        self.buttonframe2 = tk.Frame(self.menu_frame)
        self.buttonframe2.pack(side=tk.TOP, fill='x')

        # Create buttons to navigate pages
        self.open_button = tk.Button(self.buttonframe2, text="Lancer le traitement", command=self.run, height=10)
        self.open_button.pack(side=tk.TOP, fill='x', expand=True, padx=10, pady=10)

        # Add a canvas to allow scrolling
        self.canvas = tk.Canvas(self.table_frame)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame within the canvas to hold the table
        self.table_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.table_frame, anchor="nw")


    def run(self):
        print("WordTableColumnSelector run")
        dialog=self.get_dialog_col()
        character=self.get_character_col()
        if character>-1 and dialog>-1:
            print(f"ch={character} di={dialog}")
            params={
                'param_type':'WORD'
                ,'character':character,
                'dialog':dialog
            }
            threading.Thread(target=runJob,args=(self.file_path,countingMethod,params)).start()

    def destroy(self):
        print("WordTableColumnSelector destroy")
        self.menu_frame.destroy()
        self.table_frame.destroy()
        self.left_canvas.destroy()
        self.canvas.destroy()
        self.table_frame.destroy()

    def process():
        myprint7("process")

    def update_table_listbox(self):
        print("WordTableColumnSelector update listbox")
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        self.check_vars = []

        for i, _ in enumerate(self.table_list):
            var = tk.BooleanVar()
            if len(self.table_list) == 1:  # Check the checkbox by default if there's only one table
                var.set(True)
            chk = tk.Checkbutton(self.list_frame, variable=var)
            lbl = tk.Label(self.list_frame, text=f"Table {i+1}")
            chk.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            lbl.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            lbl.bind("<Button-1>", lambda e, idx=i: self.on_table_select(idx))
            self.check_vars.append(var)

        self.list_frame.update_idletasks()
        self.left_canvas.config(scrollregion=self.left_canvas.bbox("all"))

        if len(self.table_list)==1 :
            self.on_table_select(0)

    def get_dialog_col(self):
        for idx,k in (self.comboboxes.items()):
            if k.get() == "DIALOGUE" or k.get()=="LES DEUX":
                return idx
        return -1
    def get_character_col(self):
        for idx,k in (self.comboboxes.items()):
            if k.get() == "PERSONNAGE" or k.get()=="LES DEUX":
                return idx
        return -1
    def on_table_select(self, index):
        print("WordTableColumnSelector on_table_select")
        table=self.table_list[index]
        print("on_table_select idx="+str(index)+" table="+str(table))
        if len(self.detect_map)==0:
            success, mode, character,dialog,map_=detect_word_table(table,"",{})        
            self.detect_map=map_
        self.show_table_preview(table,self.detect_map)
    comboboxes={}
    detect_map={}
    column_labels = []
    def show_table_preview(self,  table,map_):
        print("WordTableColumnSelector show_table_preview")
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        num_cols = len(table.rows[0].cells)
        self.column_labels = [[] for _ in range(num_cols)]
        options = ["-", "PERSONNAGE", "DIALOGUE", "LES DEUX"]
        col_widths = [0] * num_cols
          # Calculate the max width for each column based on the content
        for row in table.rows[:3]:
            sumcolwidth=0
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text
                cell_text="\n".join(cell_text.split(" ")) 
                cell_width = tkFont.Font().measure(cell_text)
                if cell_width > col_widths[col_idx]:
                    col_widths[col_idx] = cell_width
                sumcolwidth=sumcolwidth+col_widths[col_idx]
            for col_idx, cell in enumerate(row.cells):
                if sumcolwidth>600:        
                  col_widths[col_idx] = int(col_widths[col_idx]*0.7)
                else:
                  col_widths[col_idx] = int(col_widths[col_idx])
            print("sumcolwidth"+str(sumcolwidth))
        print("colwidth"+str(col_widths))

        print("set col val")
        for col_idx in range(num_cols):
            print("combobox gen"+str(col_idx))
            combobox = ttk.Combobox(self.table_frame, values=options,width=col_widths[col_idx] // 8)
            mapval=map_[col_idx]
            print("set col val"+str(mapval))
            mapvaltype=mapval['type']
            print("set col val"+str(mapvaltype))

            if mapvaltype=='CHARACTER':
                combobox.current(1)         # Set default value to "-"
            elif mapvaltype=='DIALOG':
                        combobox.current(2) # Set default value to "-"
            elif mapvaltype=='BOTH':
                combobox.current(3) # Set default value to "-"
            else:            
                combobox.current(0)         # Set default value to "-"
            combobox.grid(row=0, column=col_idx, sticky='nsew')
            def create_on_combobox_change(col):
                def on_combobox_change(event):
                    print("change col_idx" + str(col))
                    labels = self.column_labels[col]
                    bg="white"
                    fore="black"
                    headerbg="black"
                    headerfore="white"
                    val=self.comboboxes[col].get()
                    print(val)
                    if val!='DIALOGUE' and val!='PERSONNAGE' and val!='LES DEUX':
                        fore="grey"
                        bg="#ddd"
                        headerbg="#555"
                        headerfore="#cccccc"

                    rowidx=0
                    for k in labels:
                        if rowidx==0:
                            k.config(background=headerbg, foreground=headerfore)
                        else:
                            k.config(background=bg, foreground=fore)
                        rowidx=rowidx+1
                return on_combobox_change

            combobox.bind("<<ComboboxSelected>>", create_on_combobox_change(col_idx))
            self.comboboxes[col_idx]=combobox

        myprint7("colwidth"+str(col_widths))
        for row_idx, row in enumerate(table.rows[:50]):
            for col_idx, cell in enumerate(row.cells):
                mapval=map_[col_idx]
                mapvaltype=mapval['type']
                bg="white"
                fore="black"
                headerbg="black"
                headerfore="white"
                val=self.comboboxes[col_idx].get()
                if val!='DIALOGUE' and val!='PERSONNAGE' and val!='LES DEUX':
                    fore="grey"
                    bg="#ddd"
                    headerbg="#555"
                    headerfore="#cccccc"

                cell_text = cell.text
                if row_idx==0:
                    cell_text="\n".join(cell_text.split(" ")) 
                if row_idx==0:
                    header_label = tk.Label(self.table_frame, text=cell_text, borderwidth=1, relief="solid", width=col_widths[col_idx] // 8, bg=headerbg, fg=headerfore)                
                else:
                    header_label = tk.Label(self.table_frame, text=cell_text, borderwidth=1, relief="solid", width=col_widths[col_idx] // 8, bg=bg,fg=fore)
                header_label.grid(row=row_idx + 2, column=col_idx, sticky='nsew')
                self.column_labels[col_idx].append(header_label)
        self.table_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def set_char_column(self, col):
        messagebox.showinfo("Character Column", f"Column {col+1} set as Character Column")

    def set_dialog_column(self, col):
        messagebox.showinfo("Dialog Column", f"Column {col+1} set as Dialog Column")
