#! python3
#This program is first attempt at python project
#This program is meant to automate the billing process which currently involves hand writing/typing all data points
import xml.etree.ElementTree as ET
import tkinter
from tkinter import *
from  tkinter import filedialog
from tkinter import simpledialog
import win32com.client as win32
import time
import os

#TODO make customerCell and locationCell and nameCell, signatureCell not have to be entered every time.  Some way for it to hold the data unless changed by user.

#User selects file to parse
master = Tk()
master.filename =  filedialog.askopenfilename(initialdir = "Path Where the dialog should open first",title = "Select Event File",filetypes = (("xml files","*.xml"),("all files","*.*")))
print (master.filename)

tree = ET.parse(master.filename)#this is where the file location to be parsed goes
root = tree.getroot()
         
woCell = tkinter.simpledialog.askstring("Enter WO #", "Enter WO Number") #Stores work order number
print ('Work Order #: ' + woCell)
               
#Stores user input variable for later use
truckCell = tkinter.simpledialog.askstring("Enter Truck  Number", "Enter Truck number")
print ('Truck #: ' + truckCell)
     
repairSdate = tkinter.simpledialog.askstring("Enter Repair Start Time", "Enter Repair Start Time\nFormat yyyy-mm-dd hh:mm")
print ('Start Time: ' + repairSdate)

repairCdate = tkinter.simpledialog.askstring("Enter Repair completed Time", "Enter Repair completed Time\nFormat yyyy-mm-dd hh:mm")
print ('Completed Time: ' + repairCdate)

laborCell = tkinter.simpledialog.askstring("Enter labor hours", "Enter labor hours")
print ('Labor Hours: ' +laborCell)

#this is a gui for user to select which template to use and variable selected will be used to open an existing excel template that will be included in program folder.

variable = StringVar(master)
variable.set('Select the main fault for billable event') # default value
w = OptionMenu(master, variable, 'Out of fuel', 'Dead battery from misuse', 'E-stop', 'Frozen', 'Damage', 'E-brake unplugged', 'OI unplugged', 'Other')
w.pack()

def ok():
    if variable.get() != 'Select the main fault for billable event':
        print ('Billable event due to ' +  variable.get()) #gui button.
        master.destroy()

button = Button(master, text="OK", command=ok)

button.pack()
master.wait_window(button)

master.mainloop()

dateCell = time.strftime("%Y/%m/%d")

alarmOptions = ['Unit out of fuel','EWN FC Lockout','Estop','FC stack frozen','UI can-bus failure']

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
    alarmTime = tkinter.simpledialog.askstring("Enter Event Occured  Time", "Enter Occured Time\nFormat yyyy-mm-dd hh:mm")
    print (alarmTime)
    

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
saveFile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))
saveFileNoSlash = os.path.normpath(saveFile)
wb.SaveAs(saveFileNoSlash)

excel.Application.Quit()

sys.exit()

