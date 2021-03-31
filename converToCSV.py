## pip install pandas
## pip install xlrd (only xls support)
## pip install openpyxl

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

## Variables
# input_filename = ''
input_sheetname = 'test'
#output_filename = ''
#output_sheetname = ''
csv_filetype_extension= '.csv'
excel_filetype_extension = '.xlsx'

## Constant
title = "Excel 2 CSV Converter"

## Inizialize File Dialog
window= tk.Tk()
window.title(title)
canvas1 = tk.Canvas(window, width = 300, height = 300, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

def chooseExcel ():
    global pd
    global input_filename
    
    input_filename = filedialog.askopenfilename(defaultextension=excel_filetype_extension)	
    canvas1.create_window(150, 150, window=sheetField)
    

## Export to CSV
def exportCSV ():
    readExcel()
    export_file_path = filedialog.asksaveasfilename(defaultextension=csv_filetype_extension)
    interested_sheets.to_csv(export_file_path, index = False, header=True, sep = '|')
    print('File creato con successo')

## Export to Excel
# def exportExcel ():
# writer = pd.ExcelWriter(output_filename+excel_filetype_extension, engine='xlsxwriter')
# interested_sheets.to_excel(writer, index=False, sheet_name=output_sheetname)
# workbook = writer.bookworksheet = writer.sheets[output_sheetname]
# header_fmt = workbook.add_format({'bold': True})
# worksheet.set_row(0, None, header_fmt)
# writer.save()

## Read Excel to import
def readExcel ():
    global interested_sheets

    input_sheetname = input_sheetname_obj.get()
#   print('SHEET: '+input_sheetname)
    xlsx = pd.ExcelFile(input_filename)
    xlsx_sheets = []
    for sheet in xlsx.sheet_names:
	    if sheet == input_sheetname :
		    xlsx_sheets.append(xlsx.parse(sheet)) #parse sheet into dataframe
    interested_sheets = pd.concat(xlsx_sheets) #concat multi data frame
#   print(interested_sheets)
    print('File letto con successo')

## Create buttons
readAsButton_XLS = tk.Button(text='Choose Excel File', command=chooseExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
saveAsButton_CSV = tk.Button(text='Export CSV', command=exportCSV, bg='red', fg='white', font=('helvetica', 12, 'bold'))
input_sheetname_obj = tk.StringVar()
sheetField = ttk.Entry(window, width = 15, textvariable = input_sheetname_obj)
canvas1.create_window(150, 100, window=readAsButton_XLS)
canvas1.create_window(150, 200, window=saveAsButton_CSV)

window.mainloop()


## It shows the first 5 rows
# DataFrame.head()