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

#----Init----
partsCost_entry1 = None
partsCost_entry2 = None
partsCost_entry3 = None
partsCost_entry4 = None
partsCost_entry5 = None
partReplaced_entry1 = None
partReplaced_entry2 = None
partReplaced_entry3 = None
partReplaced_entry4 = None
partReplaced_entry5 = None
partSNo_entry1 = None
partSNo_entry2 = None
partSNo_entry3 = None
partSNo_entry4 = None
partSNo_entry5 = None
partsSNi_entry1 = None
partsSNi_entry2 = None
partsSNi_entry3 = None
partsSNi_entry4 = None
partsSNi_entry5 = None
partsDesk_entry1 = None
partsDesk_entry2 = None
partsDesk_entry3 = None
partsDesk_entry4 = None
partsDesk_entry5 = None
partsHours_entry1 = None
partsHours_entry2 = None
partsHours_entry3 = None
partsHours_entry4 = None
partsHours_entry5 = None

#--------TKINTER GUI--------
window = tk.Tk()
window.geometry('850x600')
window.resizable(width=False, height=False)
window.title('Crucible')
warning = tk.Label(text='ALL FIELDS MUST ENTERED \nINCLUDING DROPDOWN MENU')
warning.grid (column=2, row=0)
warning.config(bg='YELLOW')

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
                repairSdate = str(repair_startDateGUI_Input.get()) + ' ' + str(repair_startTimeGUI_Input.get())
                global repairCdate
                repairCdate = str(repair_compDateGUI_Input.get()) + ' ' + str(repair_compTimeGUI_Input.get())
                print (variable.get())
                window.destroy()
                
def callback(selectedOE):
    if selectedOE != 'Select the main fault for billable event':
        w.config(bg = 'GREEN')

def cancel():
    window.destroy()
    sys.exit()

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
repair_compTimeGUI = tk.Label(text= '-Enter Repair Completed Time')
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
w = tk.OptionMenu(window, variable, 'Out of fuel', 'Dead battery from misuse', 'E-stop', 'Frozen', 'Damage', 'E-brake unplugged', 'OI unplugged', 'Other', command = callback)
w.grid( column=0, row=5, padx=35, pady=35)
w.config(bg = 'YELLOW')

#---PARTS WINDOW---
def partsWindow():
    
    topwindow = tk.Toplevel(master=window)
    topwindow.geometry('850x400')
    topwindow.resizable(width=False, height=False)
    topwindow.title('Parts')

    partsReplaced = tk.Label(topwindow, text= 'Part Number')
    partsReplaced.grid(column=0, row=0)
    partsReplaced_Input1 = tk.Entry(topwindow, width=8)
    partsReplaced_Input1.grid(column=0, row=1, padx=5, pady=15)
    partsReplaced_Input2 = tk.Entry(topwindow, width=8)
    partsReplaced_Input2.grid(column=0, row=2, padx=5, pady=15)
    partsReplaced_Input3 = tk.Entry(topwindow, width=8)
    partsReplaced_Input3.grid(column=0, row=3, padx=5, pady=15)
    partsReplaced_Input4 = tk.Entry(topwindow, width=8)
    partsReplaced_Input4.grid(column=0, row=4, padx=5, pady=15)    
    partsReplaced_Input5 = tk.Entry(topwindow, width=8)
    partsReplaced_Input5.grid(column=0, row=5, padx=5, pady=15)
    
    partsSNo = tk.Label(topwindow, text='Serial # Out')
    partsSNo.grid(column=1, row=0)
    partsSNo_Input1 = tk.Entry(topwindow, width=10)
    partsSNo_Input1.grid(column=1, row=1, padx=5, pady=15)
    partsSNo_Input2 = tk.Entry(topwindow, width=10)
    partsSNo_Input2.grid(column=1, row=2, padx=5, pady=15)
    partsSNo_Input3 = tk.Entry(topwindow, width=10)
    partsSNo_Input3.grid(column=1, row=3, padx=5, pady=15)
    partsSNo_Input4 = tk.Entry(topwindow, width=10)
    partsSNo_Input4.grid(column=1, row=4, padx=5, pady=15)
    partsSNo_Input5 = tk.Entry(topwindow, width=10)
    partsSNo_Input5.grid(column=1, row=5, padx=5, pady=15)
    
    partsSNi = tk.Label(topwindow, text='Serial # In')
    partsSNi.grid(column=2, row=0)
    partsSNi_Input1 = tk.Entry(topwindow, width=10)
    partsSNi_Input1.grid(column=2, row=1, padx=5, pady=15)
    partsSNi_Input2 = tk.Entry(topwindow, width=10)
    partsSNi_Input2.grid(column=2, row=2, padx=5, pady=15)
    partsSNi_Input3 = tk.Entry(topwindow, width=10)
    partsSNi_Input3.grid(column=2, row=3, padx=5, pady=15)
    partsSNi_Input4 = tk.Entry(topwindow, width=10)
    partsSNi_Input4.grid(column=2, row=4, padx=5, pady=15)
    partsSNi_Input5 = tk.Entry(topwindow, width=10)
    partsSNi_Input5.grid(column=2, row=5, padx=5, pady=15)
    
    partsDesc = tk.Label(topwindow, text='Part Description')
    partsDesc.grid(column=3, row=0)
    partsDesk_Input1 = tk.Entry(topwindow, width=70)
    partsDesk_Input1.grid(column=3, row=1, padx=5, pady=15)
    partsDesk_Input2 = tk.Entry(topwindow, width=70)
    partsDesk_Input2.grid(column=3, row=2, padx=5, pady=15)
    partsDesk_Input3 = tk.Entry(topwindow, width=70)
    partsDesk_Input3.grid(column=3, row=3, padx=5, pady=15)
    partsDesk_Input4 = tk.Entry(topwindow, width=70)
    partsDesk_Input4.grid(column=3, row=4, padx=5, pady=15)
    partsDesk_Input5 = tk.Entry(topwindow, width=70)
    partsDesk_Input5.grid(column=3, row=5, padx=5, pady=15)
    
    partsHours = tk.Label(topwindow, text='Hours')
    partsHours.grid(column=4, row=0)
    partsHours_Input1 = tk.Entry(topwindow, width=3)
    partsHours_Input1.grid(column=4, row=1, padx=5, pady=15)
    partsHours_Input2 = tk.Entry(topwindow, width=3)
    partsHours_Input2.grid(column=4, row=2, padx=5, pady=15)
    partsHours_Input3 = tk.Entry(topwindow, width=3)
    partsHours_Input3.grid(column=4, row=3, padx=5, pady=15)
    partsHours_Input4 = tk.Entry(topwindow, width=3)
    partsHours_Input4.grid(column=4, row=4, padx=5, pady=15)
    partsHours_Input5 = tk.Entry(topwindow, width=3)
    partsHours_Input5.grid(column=4, row=5, padx=5, pady=15)
        
    partsCost = tk.Label(topwindow, text='Part Cost')
    partsCost.grid(column=5, row=0)
    partsCost_Input1 = tk.Entry(topwindow, width =8)
    partsCost_Input1.grid(column=5, row=1, padx=5, pady=15)
    partsCost_Input2 = tk.Entry(topwindow, width =8)
    partsCost_Input2.grid(column=5, row=2, padx=5, pady=15)
    partsCost_Input3 = tk.Entry(topwindow, width =8)
    partsCost_Input3.grid(column=5, row=3, padx=5, pady=15)
    partsCost_Input4 = tk.Entry(topwindow, width =8)
    partsCost_Input4.grid(column=5, row=4, padx=5, pady=15)
    partsCost_Input5 = tk.Entry(topwindow, width =8)
    partsCost_Input5.grid(column=5, row=5, padx=5, pady=15)    

    partsTotal = tk.Label(topwindow, text='Total Part Cost')
    partsTotal.grid(column=4, row=6)
    partsTotal_Input = tk.Entry(topwindow, width=8)
    partsTotal_Input.grid(column=5, row=6, padx=5, pady=15)
    
    def saveParts():

        global partReplaced_entry1
        partReplaced_entry1 = partsReplaced_Input1.get()
        global partReplaced_entry2
        partReplaced_entry2 = partsReplaced_Input2.get()
        global partReplaced_entry3
        partReplaced_entry3 = partsReplaced_Input3.get()
        global partReplaced_entry4
        partReplaced_entry4 = partsReplaced_Input4.get()
        global partReplaced_entry5
        partReplaced_entry5 = partsReplaced_Input5.get()

        global partSNo_entry1
        partSNo_entry1 = partsSNo_Input1.get()
        global partSNo_entry2
        partSNo_entry2 = partsSNo_Input2.get()
        global partSNo_entry3
        partSNo_entry3 = partsSNo_Input3.get()
        global partSNo_entry4
        partSNo_entry4 = partsSNo_Input4.get()
        global partSNo_entry5
        partSNo_entry5 = partsSNo_Input5.get()

        global partsSNi_entry1
        partsSNi_entry1 = partsSNi_Input1.get()
        global partsSNi_entry2
        partsSNi_entry2 = partsSNi_Input2.get()
        global partsSNi_entry3
        partsSNi_entry3 = partsSNi_Input3.get()
        global partsSNi_entry4
        partsSNi_entry4 = partsSNi_Input4.get()
        global partsSNi_entry5
        partsSNi_entry5 = partsSNi_Input5.get()

        global partsDesk_entry1
        partsDesk_entry1 = partsDesk_Input1.get()
        global partsDesk_entry2
        partsDesk_entry2 = partsDesk_Input2.get()
        global partsDesk_entry3
        partsDesk_entry3 = partsDesk_Input3.get()
        global partsDesk_entry4
        partsDesk_entry4 = partsDesk_Input4.get()
        global partsDesk_entry5
        partsDesk_entry5 = partsDesk_Input5.get()

        global partsHours_entry1
        partsHours_entry1 = partsHours_Input1.get()
        global partsHours_entry2
        partsHours_entry2 = partsHours_Input2.get()
        global partsHours_entry3
        partsHours_entry3 = partsHours_Input3.get()
        global partsHours_entry4
        partsHours_entry4 = partsHours_Input4.get()
        global partsHours_entry5
        partsHours_entry5 = partsHours_Input5.get()
        
        global partsCost_entry1
        if partsCost_Input1.get() is None:
            partsCost_entry1 = float(0)
        else:
            partsCost_entry1 = partsCost_Input1.get()
        global partsCost_entry2
        if partsCost_Input2.get() is None:
            partsCost_entry2  = float(0)
        else:
            partsCost_entry2 = partsCost_Input2.get()
        global partsCost_entry3
        if partsCost_Input3.get() is None:
            partsCost_entry3  = float(0)
        else:
            partsCost_entry3 = partsCost_Input3.get()
        global partsCost_entry4
        if partsCost_Input4.get() is None:
            partsCost_entry4  = float(0)
        else:
            partsCost_entry4 = partsCost_Input4.get()
        global partsCost_entry5        
        if partsCost_Input5.get() is None:
            partsCost_entry5  = float(0)
        else:
            partsCost_entry5 = partsCost_Input5.get()

        topwindow.destroy()
    
    partsAcceptButton = tk.Button(topwindow, text='Add Parts', command=saveParts)
    partsAcceptButton.grid(column=3, row=7, ipadx=15, ipady=10)    
    partsCloseButton = tk.Button(topwindow, text='Cancel', command=topwindow.destroy)
    partsCloseButton.grid(column=5, row=7, ipadx=15, ipady=10)

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
button = tk.Button(window, text='OK', command=dropButton)
button.grid(column=2, row=6, ipadx=15, ipady=10)
partsButton = tk.Button(window, text ='Parts Replaced', command = partsWindow)
partsButton.grid(column=2, row=5, ipadx=10, ipady=10)
cancelButton = tk.Button(window, text='Cancel', command=cancel)
cancelButton.grid(column=3,row=6, ipadx=15, ipady=10)
window.wait_window(button)

window.mainloop()

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

    
#--------EXCEL FORM--------
occuredCell = alarmTime #TODO make alarmTime = alarm from optionmenu selection the occured time

dateCell = time.strftime("%Y/%m/%d")

__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))

excel = win32.gencache.EnsureDispatch('Excel.Application')
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
ws.Cells(23, 1).Value = partReplaced_entry1 
ws.Cells(24, 1).Value = partReplaced_entry2
ws.Cells(25, 1).Value = partReplaced_entry3
ws.Cells(26, 1).Value = partReplaced_entry4
ws.Cells(27, 1).Value = partReplaced_entry5
ws.Cells(23, 3).Value = partSNo_entry1
ws.Cells(24, 3).Value = partSNo_entry2
ws.Cells(25, 3).Value = partSNo_entry3
ws.Cells(26, 3).Value = partSNo_entry4
ws.Cells(27, 3).Value = partSNo_entry5
ws.Cells(23, 5).Value = partsSNi_entry1 
ws.Cells(24, 5).Value = partsSNi_entry2
ws.Cells(25, 5).Value = partsSNi_entry3
ws.Cells(26, 5).Value = partsSNi_entry4
ws.Cells(27, 5).Value = partsSNi_entry5
ws.Cells(23, 8).Value = partsDesk_entry1
ws.Cells(24, 8).Value = partsDesk_entry2
ws.Cells(25, 8).Value = partsDesk_entry3
ws.Cells(26, 8).Value = partsDesk_entry4
ws.Cells(27, 8).Value = partsDesk_entry5
ws.Cells(23, 15).Value = partsHours_entry1
ws.Cells(24, 15).Value = partsHours_entry2
ws.Cells(25, 15).Value = partsHours_entry3
ws.Cells(26, 15).Value = partsHours_entry4
ws.Cells(27, 15).Value = partsHours_entry5
ws.Cells(23, 17).Value = partsCost_entry1
ws.Cells(24, 17).Value = partsCost_entry2
ws.Cells(25, 17).Value = partsCost_entry3
ws.Cells(26, 17).Value = partsCost_entry4
ws.Cells(27, 17).Value = partsCost_entry5

#Savefile GUI and saving excel file
window = tk.Tk()
window.update()
saveFile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
saveFileNoSlash = os.path.normpath(saveFile)
wb.SaveAs(saveFileNoSlash)

excel.Application.Quit()

window.destroy()

sys.exit()