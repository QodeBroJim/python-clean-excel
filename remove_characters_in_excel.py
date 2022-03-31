import pandas as pd # pandas is a data science library that makes working with Excel easy
import numpy as np # numpy is a library used for numerical calculations and manipulating numbers
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
'''

Remove any character that is not alpha-numeric (a-z or 0-9)

Leaves # for suite numbers in address fields, but can be modified to replace # with ste or suite

User needs pandas & openpyxl, so need to figure out how to package everything without 
the end user needing to install dependencies - was thinking virtual env or wrapping inside docker container

'''
def import_excel_data():
    global v
    excel_file_path = askopenfilename()
    print(excel_file_path)
    v.set(excel_file_path)
    global data_frame 
    data_frame = pd.read_excel(excel_file_path)
    data_frame = data_frame.replace(np.nan, '', regex = True)
    return data_frame

def clean_excel():
    for column in data_frame.columns: # loop over all columns inside the excel file
        data_frame[column] = data_frame[column].astype(str).str.replace(r'[^A-Za-z-0-9-#]', ' ') # use regular expressions to remove all characters that are not alpha-numeric
    return data_frame.to_excel(asksaveasfilename())

root = tk.Tk()
tk.Label(root, text='File Path').grid(row = 0, column = 0)
v = tk.StringVar()
entry = tk.Entry(root, textvariable=v).grid(row=0, column=1)
tk.Button(root, text = 'Choose File', command = import_excel_data).grid(row = 1, column = 0)
tk.Button(root, text = 'Clean Excel File', command = clean_excel).grid(row = 2, column = 0)
tk.Button(root, text = 'Close', command = root.destroy).grid(row = 1, column = 1)
root.mainloop()