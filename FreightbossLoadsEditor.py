import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *

# This is the connection with creds.json that allows the app to connect to the sheet
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

# This is where the GUI code is
window = Tk()
window.geometry('450x600')
window.title("Freightboss Loads Sheet Editor")
month = StringVar(window)
month.set("DIC 19 - JAN 20")
dispatch = StringVar(window)
dispatch.set("JB")

sheetsDecisionLabel = Label(window, text="Which Sheet?")
sheetsDecisionLabel.grid(column=0, row=0)
sheetsDecision = OptionMenu(window, month,
                            "DIC 19 - JAN 20", "FEB 20", "MARCH 20", "APRIL 20", "MAY 20", "JUNE 20", "JULY 20",
                            "AUG 20", "SEP 20", "OCT 20", "NOV 20", "DEC 20")
sheetsDecision.grid(column=1, row=0)

dispatchDecisionLabel = Label(window, text="Which dispatch?")
dispatchDecisionLabel.grid(column=0, row=1)
dispatchDecision = OptionMenu(window, dispatch,
                              "JB", "MJ")
dispatchDecision.grid(column=1, row=1)

dateDecisionLabel = Label(window, text="What's the date?")
dateDecisionLabel.grid(column=0, row=2)
dateDecision = Entry(window, width=15)
dateDecision.grid(column=1, row=2)

driverDecisionLabel = Label(window, text="Which driver?")
driverDecisionLabel.grid(column=0, row=3)
driverDecision = Entry(window, width=15)
driverDecision.grid(column=1, row=3)

tkDecisionLabel = Label(window, text="What's the truck number?")
tkDecisionLabel.grid(column=0, row=4)
tkDecision = Entry(window, width=15)
tkDecision.grid(column=1, row=4)

fcDecisionLabel = Label(window, text="What's the FC#?")
fcDecisionLabel.grid(column=0, row=5)
fcDecision = Entry(window, width=15)
fcDecision.grid(column=1, row=5)

pdDecisionLabel = Label(window, text="What's the P-D?")
pdDecisionLabel.grid(column=0, row=6)
pdDecision = Entry(window, width=15)
pdDecision.grid(column=1, row=6)

lnDecisionLabel = Label(window, text="What's the Load Number?")
lnDecisionLabel.grid(column=0, row=7)
lnDecision = Entry(window, width=15)
lnDecision.grid(column=1, row=7)

brokerNameDecisionLabel = Label(window, text="What's the Broker Name?")
brokerNameDecisionLabel.grid(column=0, row=8)
brokerNameDecision = Entry(window, width=15)
brokerNameDecision.grid(column=1, row=8)

brokerPhoneDecisionLabel = Label(window, text="What's the Broker Phone Number?")
brokerPhoneDecisionLabel.grid(column=0, row=9)
brokerPhoneDecision = Entry(window, width=15)
brokerPhoneDecision.grid(column=1, row=9)

fromLabel = Label(window, text="Where is the load pick up?")
fromLabel.grid(column=0, row=10)
fromCol = Entry(window, width=15)
fromCol.grid(column=1, row=10)

toLabel = Label(window, text="Where is the load drop off?")
toLabel.grid(column=0, row=11)
toCol = Entry(window, width=15)
toCol.grid(column=1, row=11)

payLabel = Label(window, text="What is the gross pay?")
payLabel.grid(column=0, row=12)
pay = Entry(window, width=15)
pay.grid(column=1, row=12)

deliveryLabel = Label(window, text="When is the delivery and time?")
deliveryLabel.grid(column=0, row=13)
delivery = Entry(window, width=15)
delivery.grid(column=1, row=13)


# This is mostly hardcoded for the specific form my mom uses

def clicked():
    decisionSheet = month.get()
    if decisionSheet == "DIC 19 - JAN 20":
        sheetNumber = 0
    elif decisionSheet == "FEB 20":
        sheetNumber = 1
    elif decisionSheet == "MARCH 20":
        sheetNumber = 2
    elif decisionSheet == "APRIL 20":
        sheetNumber = 3
    elif decisionSheet == "MAY 20":
        sheetNumber = 4
    elif decisionSheet == "JUNE 20":
        sheetNumber = 5
    elif decisionSheet == "JULY 20":
        sheetNumber = 6
    elif decisionSheet == "AUG 20":
        sheetNumber = 7
    elif decisionSheet == "SEP 20":
        sheetNumber = 8
    elif decisionSheet == "OCT 20":
        sheetNumber = 9
    elif decisionSheet == "NOV 20":
        sheetNumber = 10
    else:
        sheetNumber = 11

    sheet = client.open("Freightboss - Loads").get_worksheet(sheetNumber)
    rowStart = len(sheet.col_values(1)) + 1

    sheet.update_cell(rowStart, 1, dispatch.get())
    sheet.update_cell(rowStart, 2, dateDecision.get())
    sheet.update_cell(rowStart, 6, driverDecision.get())
    sheet.update_cell(rowStart, 7, tkDecision.get())
    sheet.update_cell(rowStart, 8, fcDecision.get())
    sheet.update_cell(rowStart, 11, pdDecision.get())
    sheet.update_cell(rowStart, 12, lnDecision.get())
    sheet.update_cell(rowStart, 13, brokerNameDecision.get())
    sheet.update_cell(rowStart, 14, brokerPhoneDecision.get())
    sheet.update_cell(rowStart, 15, fromCol.get())
    sheet.update_cell(rowStart, 17, toCol.get())
    sheet.update_cell(rowStart, 18, pay.get())
    sheet.update_cell(rowStart, 19, delivery.get())


quit = Button(window, text='Quit', command=window.quit)
quit.grid(column=1, row=15, pady=50)
insertToSheet = Button(window, text='Click When Done', command=clicked)
insertToSheet.grid(column=1, row=14)

window.mainloop()
