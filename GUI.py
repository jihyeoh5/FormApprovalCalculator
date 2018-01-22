from tkinter import *
import os
import openpyxl
import sys
import re
from openpyxl.utils import column_index_from_string
import tkinter.messagebox
from monthlyFinanceBoardUpdated import MFBscript
import csv

root = Tk()
root.title("PowerBI Stat Sheet Update")
root.iconbitmap(r"C:\Users\joh\MS Society of Canada\InfoPath - Documents\Code\Monthly Finance Board\MontlyFinanceBoardUpdated\Broken_MS.ico")
Label(root, text = "Input directory with finance forms and press Enter. Ensure that files are titled csp/ex/vp/pr (lowercase). There must be no other files in the directory.", wraplength = 300, padx = 40, pady = 10).grid()
userEntry = Entry(root)
userEntry.grid(row=1, column = 0,pady=10,ipadx=20)
def startScript():
    global userEntry
    location = userEntry.get()
    if os.path.isdir(location)==False:
        tkinter.messagebox.showinfo("Checking user input...", "This is an invalid directory.")
    else:
        output = MFBscript(location)
        table = Listbox(root, width=45, height = 5)
        table.grid(columnspan = 4, rowspan = 5)
        line1 = ["Form","# Forms","Manager Avg","Finance Avg"]
        lines = [["EX",output["Expense"]["numForms"],output["Expense"]["manager"],output["Expense"]["finance"]],
        ["PR",output["Payment Requisition"]["numForms"],output["Payment Requisition"]["manager"],output["Payment Requisition"]["finance"]],
        ["VP",output["Vendor Payment"]["numForms"],output["Vendor Payment"]["manager"],output["Vendor Payment"]["finance"]],
        ["CSP",output["CS Payment"]["numForms"],output["CS Payment"]["manager"],output["CS Payment"]["finance"]]]
        row_format = "{:<15}  {:<15}  {:<25} {:1}"
        row1_format = "{:<8} {:<12} {:<15} {:1}"
        table.insert(END,row1_format.format(*line1))
        for items in lines:
            table.insert(END, row_format.format(*items))
submitButton = Button(root, text = "Submit", command = startScript).grid(pady=5)
root.mainloop()
