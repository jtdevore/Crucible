#! python3
#This program is first attempt at python project
#This program is meant to automate the billing process which currently involves hand writing/typing all data points
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import win32com.client as win32
import time
import os
import sys

#--------TKINTER GUI--------
window = tk.Tk()
window.geometry('800x600')
window.title('Crucible')

#---FUNTIONS---
def dropButton():
    for entry in entry_list:
        if entry.get():           
            if variable.get() != 'Select the main fault for billable event':
                global woCell
                woCell = woGUI_Input.get()
                global truckCell
                truckCell = truckGUI_Input.get()
                global laborCell
                laborCell = laborGUI_Input.get()
                global repairSdate
                repairSdate = str(repair_startDateGUI_Input.get()) + str(repair_startTimeGUI_Input.get())
                global repairCdate
                repairCdate = str(repair_compDateGUI_Input.get()) + (repair_compTimeGUI_Input.get())
                print (variable.get())
                window.destroy() 

#---WO EVENT SELECT---
window.filename =  filedialog.askopenfilename(initialdir = "Path Where the dialog should open first",title = "Select Event File",filetypes = (("xml files","*.xml"),("all files","*.*")))
window.focus_set()
window.focus_force()

#---INPUT BOX WO---
woGUI = tk.Label(text= '-Enter Work order Number')
woGUI.grid(column=0, row=0, sticky='W')
woGUI_Input = tk.Entry()
woGUI_Input.grid(column=1, row=0, padx=35, pady=35)

#---INPUT BOX TRUCK---
truckGUI = tk.Label(text= '-Enter Truck Number')
truckGUI.grid(column=0, row=1, sticky='W')
truckGUI_Input = tk.Entry()
truckGUI_Input.grid(column=1, row=1, padx=35, pady=35)

#---INPUT BOX REPAIR START---
repair_startDateGUI = tk.Label(text= '-Enter Repair Start Date (Y-M-D)')
repair_startDateGUI.grid(column=0, row=2, sticky='W')
repair_startDateGUI_Input = tk.Entry()
repair_startDateGUI_Input.grid(column=1, row=2, padx=35, pady=35)
repair_startTimeGUI = tk.Label(text= '-Enter Repair Start Time')
repair_startTimeGUI.grid(column=2, row=2, sticky='W')
repair_startTimeGUI_Input = tk.Entry()
repair_startTimeGUI_Input.grid(column=3, row=2, padx=35, pady=35)

#---INPUT BOX REPAIR COMPLETE---
repair_compDateGUI = tk.Label(text= '-Enter Repair Completed Date (Y-M-D)')
repair_compDateGUI.grid(column=0, row=3, sticky='W')
repair_compDateGUI_Input = tk.Entry()
repair_compDateGUI_Input.grid(column=1, row=3, padx=35, pady=35)
repair_compTimeGUI = tk.Label(text= '-Enter Repair Start Time')
repair_compTimeGUI.grid(column=2, row=3, sticky='W')
repair_compTimeGUI_Input = tk.Entry()
repair_compTimeGUI_Input.grid(column=3, row=3, padx=35, pady=35)

#---INPUT BOX LABOR---
laborGUI = tk.Label(text= '-Enter Labor Hours')
laborGUI.grid(column=0, row=4, sticky='W')
laborGUI_Input = tk.Entry()
laborGUI_Input.grid(column=1, row=4, padx=35, pady=35)

#---DROP MENU---
variable = tk.StringVar(window)
variable.set('Select the main fault for billable event') # default value
w = tk.OptionMenu(window, variable, 'Out of fuel', 'Dead battery from misuse', 'E-stop', 'Frozen', 'Damage', 'E-brake unplugged', 'OI unplugged', 'Other')
w.grid( column=0, row=5, padx=35, pady=35)
w.config(bg = 'YELLOW')

#---LIST---
entry_list = [
    woGUI_Input, 
    truckGUI_Input, 
    repair_startDateGUI_Input, 
    repair_startTimeGUI_Input, 
    repair_compDateGUI_Input, 
    repair_compTimeGUI_Input, 
    repair_compDateGUI_Input]

#---BUTTONS---
button = tk.Button(window, text="OK", command=dropButton)
button.grid(column=2, row=5, ipadx=15, ipady=10)
window.wait_window(button)

window.mainloop()

dateCell = time.strftime("%Y/%m/%d")

alarmOptions = ['Unit out of fuel','EWN FC Lockout','Estop','FC stack frozen','UI can-bus failure']

tree = ET.parse(window.filename)#this is where the file location to be parsed goes
root = tree.getroot()

for data in root:
    dataDict =root.attrib
    fchoursCell = (dataDict['FuelCellHours'])
    systemCell = (dataDict['SerialNumber'])
    break

for child in root:
    alarmDict = child.attrib      #This lists the alarm codes as dicts
    for alarm in alarmOptions:
        if alarm in alarmDict.values():
            print ('contains this alarm',alarm)
            alarmTime = (alarmDict['CreatedDate'])  #This pulls timestamp of alarm from xml file
            print (alarmTime)            

try:
    alarmTime      
except NameError:
    print ('No operator error alarm code found')
    window = tk.Tk()
    window.update()
    alarmTime = tk.simpledialog.askstring("Enter Event Occured  Time", "Enter Occured Time\nFormat yyyy-mm-dd hh:mm")
    print (alarmTime)
    window.destroy()
    

occuredCell = alarmTime #TODO make alarmTime = alarm from optionmenu selection the occured time
partsCostCell = None #TODO make this total all parts cost
fillableform_for_parts = None #TODO figure out how to make a fillable form for all parts entered

__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))

# this script updates excel cells.
excel = win32.gencache.EnsureDispatch('Excel.Application')
#Change this location to match the file template name as long as same folder. IE. E-stop.xlxs
selection = variable.get()
if selection == 'Out of fuel':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'oof'))
if selection == 'Dead battery from misuse':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'misuse'))
if selection == 'E-stop':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'estop'))
if selection == 'Frozen':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'frozen'))
if selection == 'Damage':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'damage'))
if selection == 'E-brake unplugged':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'ebrake'))
if selection == 'OI unplugged':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'ui'))
if selection == 'Other':
    wb = excel.Workbooks.Open(os.path.join(__location__, 'damage'))  

ws = wb.Worksheets('Copy1')

excel.Visible = False

#This is the cell formating  ws.Cells(row , column)
ws.Cells(6, 9).Value = woCell  
ws.Cells(9, 3).Value = systemCell 
ws.Cells(10, 3).Value = truckCell 
ws.Cells(11, 3).Value = fchoursCell 
ws.Cells(12, 3).Value = dateCell 
ws.Cells(8, 7).Value = occuredCell
ws.Cells(8, 10).Value = repairSdate 
ws.Cells(8, 13).Value = repairCdate 
ws.Cells(31, 15).Value = laborCell 
ws.Cells(31, 17).Value = partsCostCell  
ws.Cells(32, 15).Value = None # laborCell + partsCostCell
#todo add all the cells for the parts consumed form

#Savefile GUI and saving excel file
window = tk.Tk()
window.update()
saveFile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
saveFileNoSlash = os.path.normpath(saveFile)
wb.SaveAs(saveFileNoSlash)

excel.Application.Quit()

window.destroy()

sys.exit()


