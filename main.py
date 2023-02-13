import pandas as pd
from reportlab.lib import pagesizes
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import tabula
import csv
import tkinter as tk
from tkinter import filedialog, ttk
import os
                                                                                          
#  __/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|___/|_   
# |    /    /    /    /    /    /    /    /    /    /    /    /    /    /    /    /    /   
#/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __/_ __|    
# |/ __|/___|/   |/   |/   |/   |/   |/   |/   |/   |/   |/   |/   |/   |/   |/ _ |/       
#   /  |/  /___ _____/ /__     / /_  __  __   _______  ______ _____  ___       (_)___  ___ 
#  / /|_/ / __ `/ __  / _ \   / __ \/ / / /  / ___/ / / / __ `/ __ \/ _ \     / / __ \/ _ \
# / /  / / /_/ / /_/ /  __/  / /_/ / /_/ /  / /  / /_/ / /_/ / / / /  __/    / / /_/ /  __/
#/_/  /_/\__,_/\__,_/\___/  /_.___/\__, /  /_/   \__, /\__,_/_/ /_/\___/  __/ /\____/\___/ 
#    __      __      __      __   /____/_      _/____/ __      __      __/___/   __        
#   / / __  / / __  / / __  / / __  / / / __  / / __  / / __  / / __  / / / __  / / __     
#  / /_/ /_/ /_/ /_/ /_/ /_/ /_/ /_/ / /_/ /_/ /_/ /_/ /_/ /_/ /_/ /_/ / /_/ /_/ /_/ /_    
# /_/_  __/_/_  __/_/_  __/_/_  __/_/_/_  __/_/_  __/_/_  __/_/_  __/_/_/_  __/_/_  __/    
#(_) /_/ (_) /_/ (_) /_/ (_) /_/ (_|_) /_/ (_) /_/ (_) /_/ (_) /_/ (_|_) /_/ (_) /_/       
                                                                                          
def csv_to_pdf(x, y):
    def wrap_text(data, max_width):
        wrapped = []
        for row in data:
            wrapped_row = []
            for item in row:
                words = str(item).split(" ")
                line = ""
                for word in words:
                    if len(line + word) < max_width:
                        line += word + " "
                    else:
                        wrapped_row.append(line)
                        line = word + " "
                wrapped_row.append(line)
            wrapped.append(wrapped_row)
        return wrapped
    df = pd.read_csv(x, encoding='ISO-8859-1')
    df.dropna(how='all', inplace=True)

    pdf_file = SimpleDocTemplate(y, pagesize=pagesizes.A4)

    max_width = 2000000
    wrapped_data = [df.columns.tolist()] + wrap_text(df.values.tolist(), max_width)

    rows_per_table = 200000
    col_width = 1.5 * inch
    for i in range(0, len(wrapped_data), rows_per_table):
        data = wrapped_data[i:i + rows_per_table]
        table = Table(data, colWidths=[col_width] * df.shape[1], repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        pdf_file.build([table])

def pdf_to_csv(pdf_file):
    df = tabula.read_pdf(pdf_file, pages='all')[0]
    df.to_csv('output', index=False)

def choose_file():
    file_path = filedialog.askopenfilename()
    return file_path

def csv_to_pdf_command():
    file_path = choose_file()
    csv_to_pdf(file_path, file_path.replace(".csv", ".pdf"))
    status_label.config(text="CSV to PDF conversion complete!", foreground='green')

def pdf_to_excel_command():
    file_path = choose_file()
    pdf_to_csv(file_path)
    status_label.config(text="PDF to Excel conversion complete!", foreground='green')

root = tk.Tk()
root.geometry("600x400")
root.title("File Converter")

frame = ttk.Frame(root, padding=20)
frame.pack()

title_label = ttk.Label(frame, text="PDF TO CSV FOR COSMETIC CLINICAL TRIAL DATA", font=("TkDefaultFont", 15))
title_label.pack()

csv_to_pdf_button = ttk.Button(frame, text="Convert CSV to PDF", command=csv_to_pdf_command)
csv_to_pdf_button.pack()

pdf_to_excel_button = ttk.Button(frame, text="Convert PDF to Excel", command=pdf_to_excel_command)
pdf_to_excel_button.pack()
made_by_label = ttk.Label(frame, text="Made By Ryane Joe", font=("TkDefaultFont", 10))
made_by_label.pack(pady=10)

status_label = ttk.Label(frame, text="", font=("TkDefaultFont", 10))
status_label.pack(pady=10)

root.mainloop()
# __ __  __  __  ___   ____   __  _____   __ __  __  _ ___   __   __  ___  
#|  V  |/  \| _\| __| |  \ `v' / | _ \ `v' //  \|  \| | __| |_ \ /__\| __| 
#| \_/ | /\ | v | _|  | -<`. .'  | v /`. .'| /\ | | ' | _|   _\ | \/ | _|  
#|_| |_|_||_|__/|___| |__/ !_!   |_|_\ !_! |_||_|_|\__|___| /___|\__/|___| 

