import win32com.client
import os

# Set the path of the excel file and the desired output directory
input_dir = os.getcwd() + "\\02 Export Excel to PDF\\excel\\"
output_dir =  os.getcwd() + "\\02 Export Excel to PDF\\pdf\\"

# Setup the Connection
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(input_dir + "excel.xlsx")


# Iterating over every Worksheet to export it to desired path
for ws in wb.Worksheets:

    # deleting old PDFs if exists
    output_file = ws.Name + ".pdf"

    if output_file in os.listdir(output_dir):
        os.remove(output_dir + output_file)

    # saving new version
    ws.ExportAsFixedFormat(0, output_dir + ws.Name)

# Closing Connection
wb.Close(True)