from numpy import intp
import win32com.client
import os
import glob
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
input_path = "INPUT PATH HERE"
output_path = "OUTPUT PATH HERE"
files = glob.glob(input_path + "/*.xls")
for filename in files:
    file = os.path.basename(filename)
    output = output_path + '/' + file.replace('.xls','.xlsx')
    wb = o.Workbooks.Open(filename)
    wb.ActiveSheet.SaveAs(output,51)
    wb.Close(True)
