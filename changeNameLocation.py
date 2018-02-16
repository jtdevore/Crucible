#! python3
#This is a program to update customerCell and locationCell and nameCell, signatureCell on the excel file templates in Crucible.

import win32com.client as win32
import os
import tkinter as tk
import sys

#--------FUNTIONS---------
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



def finished():
    global customerCell
    customerCell = customerInput.get()
    global locationCell 
    locationCell = locationInput.get()
    global nameCell
    nameCell = nameInput.get()

    window.destroy()    

#--------GUI--------
window = tk.Tk()
window.geometry('200x400')
window.title('Cell Editor')

#---INPUT BOX CUSTOMER---
customer = tk.Label(text= 'Enter Customer Name')
customer.grid(column=0, row=0)
customerInput = tk.Entry()
customerInput.grid(column=0, row=1, pady=35)

#---INPUT BOX LOCATION---
location = tk.Label(text= 'Enter Location')
location.grid(column=0, row=2)
locationInput = tk.Entry()
locationInput.grid(column=0, row=3, pady=35)

#---INPUT BOX NAME---
name = tk.Label(text= 'Enter Service Technicians Name')
name.grid(column=0, row=4)
nameInput = tk.Entry()
nameInput.grid(column=0, row=5, pady=35)

#---BUTTON---
button_ok = tk.Button(text='OK', command=finished)
button_ok.grid(column=0, row=6, ipadx=15, ipady=10)

window.mainloop()

__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))

excel = win32.gencache.EnsureDispatch('Excel.Application')

fileNames = ['oof', 'misuse', 'estop','frozen','damage','ebrake','ui']
for file in fileNames:
    updateWb(file)

sys.exit()