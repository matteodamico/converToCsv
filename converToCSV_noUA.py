## pip install pandas
## pip install xlrd (only xls support)
## pip install openpyxl

import pandas as pd
import sys, getopt

## Variables
# input_filename = ''
input_sheetname = 'test'
#output_filename = ''
#output_sheetname = ''
csv_filetype_extension= '.csv'
excel_filetype_extension = '.xlsx'

## Constant
title = "Excel 2 CSV Converter"


## Export to CSV
def exportCSV (input_file_path,input_sheetname,export_file_path):
    readExcel(input_file_path,input_sheetname)
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
def readExcel (input_file_path,input_sheetname):
    global interested_sheets

	
#   print('SHEET: '+input_sheetname)
    xlsx = pd.ExcelFile(input_file_path)
    xlsx_sheets = []
    for sheet in xlsx.sheet_names:
	    if sheet == input_sheetname :
		    xlsx_sheets.append(xlsx.parse(sheet)) #parse sheet into dataframe
    interested_sheets = pd.concat(xlsx_sheets) #concat multi data frame
#   print(interested_sheets)
    print('File letto con successo')

def main(argv):

    input_file_path = ''
    export_file_path = ''
    input_sheetname = ''

    try:
        opts, args = getopt.getopt(argv,"hi:o:s:",["ifile=","ofile=","sheetname="])
    except getopt.GetoptError:
        print ('USAGE: converToCSV_nonInt.py -i <inputfile> -o <outputfile> -sname <sheetname>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('USAGE: converToCSV_nonInt.py -i <inputfile> -o <outputfile> -sname <sheetname>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input_file_path = arg
        elif opt in ("-o", "--ofile"):
            export_file_path = arg
        elif opt in ("-s", "--sheetname"):
            input_sheetname = arg
    print('INPUT: '+input_file_path)
    print('OUTPUT: '+export_file_path)
    print('SHEETNAME: '+input_sheetname)
    exportCSV(input_file_path,input_sheetname,export_file_path)		 
		 
if __name__ == "__main__":
   main(sys.argv[1:])		 