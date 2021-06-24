#Daily Timesheet
import tkinter as tk
from openpyxl import Workbook
import datetime as dt
# Options menu
from tkinter import OptionMenu
from tkinter import StringVar
# Save dialog
from tkinter.filedialog import asksaveasfilename
# Imports necessary style classes for openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors

# Variables for openpyxl
today = dt.date.today()
workbook = Workbook()
sheet = workbook.active

# Set column width of Excel doc
sheet.column_dimensions['A'].width = 16
sheet.column_dimensions['B'].width = 12
sheet.column_dimensions['C'].width = 12
sheet.column_dimensions['D'].width = 12
sheet.column_dimensions['E'].width = 12
sheet.column_dimensions['F'].width = 12
sheet.column_dimensions['G'].width = 12

# Styles for Excel sheet
bold_font = Font(bold=True)

# Defines the save_xlsx command below
def save_xlsx():
    # Takes entries then inputs to spreadsheet
    sheet['A1'].alignment = Alignment(horizontal='center')
    sheet["A1"].font = bold_font
    sheet["A1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet["A1"] = name.get()
    sheet["A2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['A2'].alignment = Alignment(horizontal='center')
    sheet["A2"] = "Location:"
    sheet["A3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['A3'].alignment = Alignment(horizontal='center')
    sheet["A3"].font = bold_font
    sheet["A3"] = loc.get()
    sheet['A4'].alignment = Alignment(horizontal='center')
    sheet["A4"].font = bold_font
    sheet["A4"] = loc_02.get()
    sheet["A4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["A5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet['A5'].alignment = Alignment(horizontal='center')
    sheet["A5"].font = bold_font
    sheet["A5"] = loc_03.get()
    sheet["A6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["B1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet["B2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['B2'].alignment = Alignment(horizontal='center')
    sheet["B2"] = "Travel Time:"
    sheet["B3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['B3'].alignment = Alignment(horizontal='center')
    sheet["B3"].font = bold_font
    sheet["B4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["B4"] = travel.get()
    sheet['B4'].alignment = Alignment(horizontal='center')
    sheet["B4"].font = bold_font
    sheet["B5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["B5"] = travel_02.get()
    sheet['B5'].alignment = Alignment(horizontal='center')
    sheet["B5"].font = bold_font
    sheet["B6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["C1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['C1'].alignment = Alignment(horizontal='center')
    sheet["C1"] = "Date:"
    sheet["C2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['C2'].alignment = Alignment(horizontal='center')
    sheet["C2"] = "Arrival Time:"
    sheet["C3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['C3'].alignment = Alignment(horizontal='center')
    sheet["C3"].font = bold_font
    sheet["C3"] = start_time.get()
    sheet["C4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["C4"] = arrival_02.get()
    sheet['C4'].alignment = Alignment(horizontal='center')
    sheet["C4"].font = bold_font
    sheet["C5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["C5"] = arrival_02.get()
    sheet['C5'].alignment = Alignment(horizontal='center')
    sheet["C5"].font = bold_font
    sheet["C6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["D1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['D1'].alignment = Alignment(horizontal='center')
    sheet["D1"].font = bold_font
    sheet["D1"] = date_.get()
    sheet["D2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['D2'].alignment = Alignment(horizontal='center')
    sheet["D2"] = "End Time:"
    sheet["D3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['D3'].alignment = Alignment(horizontal='center')
    sheet["D3"].font = bold_font
    sheet["D3"] = end_time.get()
    sheet["D4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["D4"] = end_01.get()
    sheet['D4'].alignment = Alignment(horizontal='center')
    sheet["D4"].font = bold_font
    sheet["D5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["D5"] = end_02.get()
    sheet['D5'].alignment = Alignment(horizontal='center')
    sheet["D5"].font = bold_font
    sheet["D6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["E1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['E1'].alignment = Alignment(horizontal='center')
    sheet["E1"] = "Lunch Time:"
    sheet["E2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet["E3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet["E4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["E5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["E6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["F1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['F1'].alignment = Alignment(horizontal='center')
    sheet["F1"].font = bold_font
    sheet["F1"] = lunch_out.get()
    sheet["F2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['F2'].alignment = Alignment(horizontal='center')
    sheet["F2"] = "Miles:"
    sheet["F3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['F3'].alignment = Alignment(horizontal='center')
    sheet["F4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["F4"] = miles.get()
    sheet['F4'].alignment = Alignment(horizontal='center')
    sheet["F4"].font = bold_font
    sheet["F5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["F5"] = miles_02.get()
    sheet['F5'].alignment = Alignment(horizontal='center')
    sheet["F5"].font = bold_font
    sheet["F6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["G1"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['G1'].alignment = Alignment(horizontal='center')
    sheet["G1"].font = bold_font
    sheet["G1"] = lunch_in.get()
    sheet["G2"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['G2'].alignment = Alignment(horizontal='center')
    sheet["G2"] = "Expenses:"
    sheet["G3"].border = Border(top = Side(border_style='thin', color='000000'),
                                right = Side(border_style='thin', color='000000'),
                                bottom = Side(border_style='thin', color='000000'),
                                left = Side(border_style='thin', color='000000'))
    sheet['G3'].alignment = Alignment(horizontal='center')
    sheet['G4'].alignment = Alignment(horizontal='center')
    sheet["G4"].font = bold_font
    sheet["G4"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["G4"] = expenses.get()
    sheet['G5'].alignment = Alignment(horizontal='center')
    sheet["G5"].font = bold_font
    sheet["G5"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))
    sheet["G5"] = expenses_02.get()
    sheet["G6"].border = Border(top=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'))

    files = [('Excel Workbook', '*.xlsx')]
    workbook.save(asksaveasfilename(filetypes=files, defaultextension=files, initialfile=f"{today}-{name.get()}.xlsx"))

# Title
master = tk.Tk()
master.title("Daily Time (Version 1.1)")

# Dropdown options
OPTIONS = [
"----------",
"Joe Alvarado",
"Jane Doe",
"John Smith",
]

name = StringVar(master)
name.set(OPTIONS[0]) # default value

name_select = OptionMenu(master, name, *OPTIONS)
name_select.grid(column=1, sticky="w")

# Labels for entry fields
tk.Label(master,
         text="Name").grid(row=0)
tk.Label(master,
         text="Date").grid(row=0, column=2)
tk.Label(master,
         text="Location 1").grid(row=2)
tk.Label(master,
         text="Location 2").grid(row=3)
tk.Label(master,
         text="Location 3").grid(row=4)
tk.Label(master,
         text="Arrival Time").grid(row=2, column=2)
tk.Label(master,
         text="Lunch Out").grid(row=0, column=4)
tk.Label(master,
         text="Lunch In").grid(row=0, column=6)
tk.Label(master,
         text="End Time").grid(row=2, column=4)
tk.Label(master,
         text="Miles").grid(row=3, column=9)
tk.Label(master,
         text="Expenses").grid(row=3, column=11)
tk.Label(master,
         text="Travel Time").grid(row=3, column=2)
tk.Label(master,
         text="Arrival Time").grid(row=3, column=4)
tk.Label(master,
         text="End Time").grid(row=3, column=6)
tk.Label(master,
         text="Travel Time").grid(row=4, column=2)
tk.Label(master,
         text="Arrival Time").grid(row=4, column=4)
tk.Label(master,
         text="End Time").grid(row=4, column=6)
tk.Label(master,
         text="Miles").grid(row=4, column=9)
tk.Label(master,
         text="Expenses").grid(row=4, column=11)

# Entry fields
date_ = tk.Entry(master, width=9)
date_.insert(10, today)
loc = tk.Entry(master, width=9)
loc.insert(10, "Remote")
loc_02 = tk.Entry(master, width=9)
loc_03 = tk.Entry(master, width=9)
start_time = tk.Entry(master, width=9)
start_time.insert(10, "08:00 AM")
lunch_out = tk.Entry(master, width=9)
lunch_out.insert(10, "12:00 PM")
lunch_in = tk.Entry(master, width=9)
lunch_in.insert(10, "01:00 PM")
end_time = tk.Entry(master, width=9)
end_time.insert(10, "05:00 PM")
expenses = tk.Entry(master, width=9)
miles = tk.Entry(master, width=9)
travel = tk.Entry(master, width=9)
travel_02 = tk.Entry(master, width=9)
arrival_01 = tk.Entry(master, width=9)
end_01 = tk.Entry(master, width=9)
expenses_02 = tk.Entry(master, width=9)
miles_02 = tk.Entry(master, width=9)
travel_02 = tk.Entry(master, width=9)
arrival_02 = tk.Entry(master, width=9)
end_02 = tk.Entry(master, width=9)

date_.grid(row=0, column=3)
loc.grid(row=2, column=1)
loc_02.grid(row=3, column=1)
loc_03.grid(row=4, column=1)
start_time.grid(row=2, column=3)
lunch_out.grid(row=0, column=5)
lunch_in.grid(row=0, column=7)
end_time.grid(row=2, column=5)
expenses.grid(row=3, column=10)
miles.grid(row=3, column=12)
travel.grid(row=3, column=3)
travel_02.grid(row=4, column=3)
arrival_01.grid(row=3, column=5)
end_01.grid(row=3, column=7)
expenses_02.grid(row=4, column=10)
miles_02.grid(row=4, column=12)
arrival_02.grid(row=4, column=5)
end_02.grid(row=4, column=7)


# Buttons
tk.Button(master,
          text='Quit',
          command=master.quit).grid(row=23,
                                    column=0,
                                    sticky=tk.W,
                                    pady=1)

tk.Button(master,
          text='Save', command=save_xlsx).grid(row=23,
                                                       column=1,
                                                       sticky=tk.W,
                                                       pady=1)

# Footer
tk.Label(master, text="Joe Alvarado", font=("Courier", 9)).grid(row=26, column=0, sticky="w")
tk.Label(master, text="(v. 1.1)", font=("Courier", 9)).grid(row=26, column=1, sticky="w")

tk.mainloop()
