from pdfplumber.pdf import PDF
import pdfplumber          
import pandas as pd
from docx import Document
import logging                                                   

def myprint7(s):
    s="ui: "+str(s)
    logging.debug(s)
    #print(s.encode('utf-8'))

def convert_pdf_to_universaltables(file_path):
        res=[]
        with pdfplumber.open(file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                myprint7("readpdf page"+str(page_number))
                tables = page.extract_tables()
                
                for table_number, table in enumerate(tables, start=1):
                    
                    row_count = len(table)
                    column_count = len(table[0]) if row_count > 0 else 0
                    

                    myprint7(f"Page {page_number} Table {table_number}:")
                    myprint7(f"Row count: {row_count}")
                    myprint7(f"Column count: {column_count}")
                    
                    rescells=[]
                    for row_index, row in enumerate(table):
                        #myprint7(f"Add row : {row_index} {row}")
                        resrow=[]
                        for col_index, cell in enumerate(row):
                            #myprint7(f"myCell({row_index}, {col_index}): {cell}")
                            resrow.append(cell)
                                
                        if False:
                            for col_index in range(1,column_count):
                                myprint7(f"Cell idx={col_index}")
                                cell =row[col_index]
                                myprint7(f"Cell {col_index} cell={cell}")
                                row.append(cell)
                                myprint7(f"Cell({row_index}, {col_index}): {cell}")
                        rescells.append(resrow)
                    myprint7("\n")

                    table={
                        "row_count":row_count,
                        "col_count":column_count,
                        "cells":rescells
                    }
                    res.append(table)
        return res





def convert_docx_to_universaltables(file_path):
        myprint7("convert_docx_to_universaltables")
        res=[]
        table_dimensions=[]
        doc = Document(file_path)
        for table in doc.tables:
            num_rows = len(table.rows)
            num_columns = len(table.columns)
            table_dimensions.append((num_rows, num_columns))


        for k,table in enumerate(doc.tables):
                rescells=[]
                row_count = len(table.rows)
                column_count = len(table.columns)
                # Get the first table   
                for row_index,row in enumerate(table.rows):
                    resrow=[]
                    for col_index,cell in enumerate(row.cells):
                        t=cell.text.strip()
                        resrow.append(t)                                
                        #myprint7(f"myCell({row_index}, {col_index}): {t}")
                    rescells.append(resrow)
                table={
                    "row_count":row_count,
                    "col_count":column_count,
                    "cells":rescells
                }
                res.append(table)
        return res

def access_cell(df,i, j):
    try:
        cell_value = df.iloc[i, j]
        myprint7(f"Read {i} {j} {cell_value} ")
        
        return cell_value
    except IndexError:
        myprint7(f"ERR Index out of range {i} {j}")
        return "Index out of range"

def convert_xlsx_to_universaltables(file_path):
        res=[]
        table_dimensions=[]
        df = pd.read_excel(file_path,header=None)

        num_rows = len(table.rows)
        num_columns = len(table.columns)
        table_dimensions.append((num_rows, num_columns))


        rescells=[]
        row_count = df.shape[0]
        column_count = df.shape[1]
        # Get the first table   
        for row_index in range(0,row_count):
            row_data = df.iloc[row_index, :]
            resrow=[]
            # Loop over the cells of the row
            for col_index, cell in enumerate(row_data):   
                t=access_cell(df,row_index, col_index)

                resrow.append(t.strip())                                
                #myprint7(f"myCell({row_index}, {col_index}): {cell}")

        table={
            "row_count":row_count,
            "col_count":column_count,
            "cells":rescells
        }
        res.append(table)
        return res