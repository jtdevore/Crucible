#! python3
#This is a program to update customerCell and locationCell and nameCell, signatureCell on the excel file templates.

import win32com.client as win32
import os
from tkinter import simpledialog
from tkinter import * 

#GUI's for user input
root = Tk()
customerCell = simpledialog.askstring("Enter Customer Name", "Enter Customer Name")
locationCell = simpledialog.askstring("Enter Location", "Enter Location")
nameCell = simpledialog.askstring("Enter Technician Name", "Enter Technician Name")
root.destroy()
root.mainloop()


# finds absolute file path for excel templates
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))

# finds the program to open, ie excel
excel = win32.gencache.EnsureDispatch('Excel.Application')

#Updates cells in excel workbook
def updateWb(file):
    wb = excel.Workbooks.Open(os.path.join(__location__, file))
    excel.Visible = False
    ws = wb.Worksheets('Copy1')
    ws.Cells(6, 3).Value = customerCell
    ws.Cells(7, 3).Value = locationCell
    ws.Cells(6, 12).Value = nameCell
    ws.Cells(34, 2).Font.Italic = True
    ws.Cells(34, 2).Value = nameCell
    wb.Save()
    excel.Application.Quit()

fileNames = ['oof', 'misuse', 'estop','frozen','damage','ebrake','ui']

for file in fileNames:
    updateWb(file)

    sys.exit()