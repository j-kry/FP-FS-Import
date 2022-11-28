import tkinter as tk
from tkinter.filedialog import askopenfilename
import openpyxl

from openpyxl import Workbook
from openpyxl import load_workbook

import requests

import time

#Will store data from csv in a Ticket Object
class Ticket:

    def __init__(self, subject, description):
        self.subject = subject
        self.description = description
    
    def __repr__(self):
        return self.subject + "\n" + self.description

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
#####CONSTANTS####
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

JSONHEADER = {'Content-Type' : 'application/json'}

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

##Prep for category by column
##if(sheet.cell == "Demographic Change")
##category = "Demographic Change"
##switch(sheet.cell)
##case "Demographic Change": category = "Demographic Change" break

#Iterate through the sheet starting at the second row
#numTickets-1 because we are adding 2 to the row. Loop range is not inclusive.
for row in range(numTickets-1):
    tickets.append(
        Ticket(sheet.cell(row=row+2, column=9).value, 
        "Original ticket number: " + sheet.cell(row=row+2, column=2).value + "<br>" +
        "Created on: " + str(sheet.cell(row=row+2, column=5).value) + "<br>" +
        "Reason for request: " + sheet.cell(row=row+2, column=10).value + "<br>" +
        "Request area: " + sheet.cell(row=row+2, column=11).value + "<br>" +
        "Request sub-type: " + sheet.cell(row=row+2, column=12).value + "<br>" +
        "Originally submitted by: " + sheet.cell(row=row+2, column=15).value + " " + sheet.cell(row=row+2, column=16).value + "<br>" +
        "Job Title: " + sheet.cell(row=row+2, column=19).value + "<br>" +
        "Team: " + sheet.cell(row=row+2, column=20).value + "<br>" +
        "CC: " + sheet.cell(row=row+2, column=27).value + "<br>" +
        "Attachment names: " + sheet.cell(row=row+2, column=30).value + "<br>" +
        "Original Description: " + sheet.cell(row=row+2, column=13).value + "<br>")
        )

# for i in tickets:
#     print(i)

counter = 1

for j in tickets:

    #Source is Slack to differentiate from 'real' tickets in reports
    payload = {'requester_id': JUSTIN_REQUESTER_ID, 
    'group_id' : JUSTIN_GROUP_ID, 
    'subject' : j.getSubject(),
    'status' : STATUS,
    'priority' : PRIORITY,
    'description' : j.getDescription(),
    'source' : SOURCE,
    'category' : CATEGORY
    }

    r = requests.post(
        'https://thresholds.freshservice.com/api/v2/tickets',
        json=payload, 
        headers=JSONHEADER,
        auth=("API KEY GOES HERE", "X")
        )

    print("SENDING TICKET " + str(counter) + " of " + str(numTickets-1))
    counter+=1

    #Wait a second every other ticket to avoid API rate limit
    if(counter%2 == 0):
        time.sleep(1)