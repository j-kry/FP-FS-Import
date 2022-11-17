import tkinter as tk
from tkinter.filedialog import askopenfilename
import openpyxl
import os

from openpyxl import Workbook
from openpyxl import load_workbook

import requests

#Will store data from csv in a Ticket Object
class Ticket:

    def __init__(self, subject, description):
        self.subject = subject
        self.description = description
    
    def __repr__(self):
        return self.subject + " " + self.description

    def getSubject(self):
        return self.subject
    def getDescription(self):
        return self.description

    def setSubject(self, subject):
        self.subject = subject
    def setDescription(self, description):
        self.description = description

def FileOpen():

    #Path = the one from the window
    filePath = askopenfilename(filetypes=[("Microsoft Excel Worksheet", ".xlsx"), ("All Files", "*.*")])

    if not filePath:
        print("No file present or none selected!")
        return

    return filePath 

##################
#TICKET CONSTANTS#
##################

#Phil ID
REQUESTER_ID = 20001321742
#Justin-Test ID
JUSTIN_REQUESTER_ID = 20000651361
#Phil Group EMR Billing
GROUP_ID = 20000342334
#Justin-Test Group ID
JUSTIN_GROUP_ID = 20000342352

#open
STATUS = 2
#low
PRIORITY = 1
#slack, not to muddy reports
SOURCE = 10
#hidden category
CATEGORY = "Imported Billing Ticket"
#tags need to be an array
TAGS = ["Imported Billing Ticket"]


####################
#####Begin MAIN#####
####################

# Load workbook from file prompt
wb = load_workbook(filename=FileOpen())

# Load worksheet
sheet = wb.active

# Get number of tickets in spreadsheet
numTickets = sheet.max_row

print("READ " + str(numTickets-1) + " TICKETS FROM EXCEL")

#tickets array will hold all Ticket Objects
tickets = []

#Iterate through the sheet starting at the second row
#
for row in range(numTickets-1):
    tickets.append(Ticket(sheet.cell(row=row+2, column=2).value, sheet.cell(row=row+2, column=9).value))
    print("ROW " + str(row))

for i in tickets:
    print(i)