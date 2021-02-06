import pandas as pd
import portable_spreadsheet as ps
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import numpy as np
import os
from pathlib import Path
import itertools
from pandasql import *
pysqldf = lambda q:sqldf(q,globals())
import tkinter as tk
from tkinter import filedialog


# Function to create a gloabal variables spec_path (Path for Specification Template File) and spec_xlsx (Specification Template File).
def path_file():
    global spec_xlsx
    global spec_path
    global spec_file
    spec_xlsx = filedialog.askopenfilename(title = 'Select Specification Template File')
    #spec_path = filedialog.askdirectory()
    spec_file = Path(spec_xlsx).name
    spec_path = os.path.dirname(spec_xlsx)
    final_str = "Pass"
    label['text'] = final_str

# Function to take user input for the Endpoint to be created
def domain_list(spec_gen):
    spec_gen = re.sub(r"\s+", "", input('Input specification to be created: ').strip().upper(), flags=re.UNICODE)
    try:
        #if spec_file not in globals():
        if len(fval_list) == 0:
            final_str = "Please Provide Domains to be Created!!!"
        elif len(spec_file) == 0:
            final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specification template file."
        else:
            final_str = f"The Domain(s) {', '.join(map(str, fval_list))} will be created in \n{spec_file} \nat \n{spec_path}. \nClick on 'Execute' to generate the specification"
        #final_str = 'City: %s \nConditions: %s \nTemperature (°F): %s' % (name, desc, temp)
        #final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specificatin template file."
    except NameError:
           final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specification template file."
    label['text'] = final_str


# Function to create a list of domains input by User.
def domain_list(fval):
    global fval_list
    fval_list = np.unique([f.strip().upper() for f in fval.split(',') if fval]).tolist()
    fval_list = list(filter(bool, fval_list))
    try:
        #if spec_file not in globals():
        if len(fval_list) == 0:
            final_str = "Please Provide Domains to be Created!!!"
        elif len(spec_file) == 0:
            final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specification template file."
        else:
            final_str = f"The Domain(s) {', '.join(map(str, fval_list))} will be created in \n{spec_file} \nat \n{spec_path}. \nClick on 'Execute' to generate the specification"
        #final_str = 'City: %s \nConditions: %s \nTemperature (°F): %s' % (name, desc, temp)
        #final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specificatin template file."
    except NameError:
           final_str = f"The Domain(s) entered to be created are: {fval_list}. \nPlease navigate to specification template file."
    label['text'] = final_str

# Function to Generate the SDTM Specification

#Creating gui
HEIGHT = 500
WIDTH = 600

root = tk.Tk()
root.title("SRDM to SDTM Mapping Specification",)

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH, bg='#AFEEEE')
canvas.pack()

#background_image = tk.PhotoImage(file='landscape.png')
#background_label = tk.Label(root, image=background_image)
#background_label.place(relwidth=1, relheight=1)

#File Browser frame
frame_browse = tk.LabelFrame(root, bg='#0086ad', bd=5, text="Navigate to Specification Template File", font=("Perpetua",12))
frame_browse.place(relx=0.5, rely=0.1, relwidth=0.45, relheight=0.1, anchor='n')

button_browse = tk.Button(frame_browse, text="Browse", font=("Gabriola",15), command=lambda: path_file())
button_browse.place(relheight=1, relwidth=1)

#Domain List by User Input Frame
frame_domain = tk.Frame(root, bg='#0086ad', bd=5)
frame_domain.place(relx=0.5, rely=0.25, relwidth=0.75, relheight=0.1, anchor='n')

entry = tk.Entry(frame_domain, font=40)
entry.place(relwidth=0.65, relheight=1)

button_domain = tk.Button(frame_domain, text="Enter Domains", font=("Gabriola",15), command=lambda: domain_list(entry.get()))
button_domain.place(relx=0.7, relheight=1, relwidth=0.3)
#root.bind('<Return>', domain_list)

lower_frame = tk.Frame(root, bg='#0086ad', bd=5)
lower_frame.place(relx=0.5, rely=0.4, relwidth=0.8, relheight=0.3, anchor='n')

label = tk.Label(lower_frame)
label.place(relwidth=1, relheight=1)

#Execution Frame
frame_execute = tk.Frame(root, bg='#6B8E23', bd=3)
frame_execute.place(relx=0.5, rely=0.8, relwidth=0.25, relheight=0.08, anchor='n')

button_execute = tk.Button(frame_execute, text="Execute", font=("Gabriola",20), command=lambda: path_file())
button_execute.place(relheight=1, relwidth=1)

root.mainloop()
