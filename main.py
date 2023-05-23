import os
import customtkinter as ctk
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Functions
def validate_numeric_input(new_value):
    if new_value == "":
        return True

    try:
        float(new_value)
        return True
    except ValueError:
        return False
    
def on_button_click():
    item_name = item_name_entry_box.get()
    item_price = item_price_entry_box.get()
    selected_date = date_combobox.get()

    if item_name and item_price and selected_date != "Select Date":
        workbook = load_workbook(file_path) if os.path.exists(file_path) else Workbook()
        sheet = workbook.active


        last_row = sheet.max_row

        if last_row >= 3:
            sheet.delete_rows(last_row)

        total_items = last_row - 2

        total_price = sum(
            float(sheet.cell(row=row_num, column=2).value)
            for row_num in range(2, last_row)
            if sheet.cell(row=row_num, column=2).value is not None
        )

        sheet.insert_rows(2)

        sheet.cell(row=2, column=1).value = item_name
        sheet.cell(row=2, column=2).value = item_price
        sheet.cell(row=2, column=3).value = selected_date

        # Create a new footer with the updated values
        sheet.cell(row=last_row + 1, column=1).value = f"Total items: {total_items}"
        sheet.cell(row=last_row + 1, column=2).value = f"Total price: {total_price}"

        workbook.save(file_path)

        message_label.configure(text="Information noted!")
        main_window.after(2000, lambda: message_label.configure(text=""))
    else:
        message_label.configure(text="Please fill all fields!")
        main_window.after(2000, lambda: message_label.configure(text=""))


# Database stuff
folder_path = "C:/Expense-Tracker"
file_path = f"{folder_path}/Expenses.xlsx"

# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

if os.path.exists(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active
    headers = ["Item", "Price in ₱", "Date of Purchase"]
    existing_headers = [sheet.cell(row=1, column=col_num).value for col_num in range(1, sheet.max_column + 1)]

    if existing_headers != headers:
        sheet.delete_rows(1)
        for col_num, header in enumerate(headers, start=1):
            sheet.cell(row=1, column=col_num).value = header

        footer_labels = ["Total items: ", "Total price: "]
        for col_num, label in enumerate(footer_labels, start=1):
            sheet.cell(row=3, column=col_num).value = label

        sheet.column_dimensions["A"].width = 30
        sheet.column_dimensions["B"].width = 30
        sheet.column_dimensions["C"].width = 30

        workbook.save(file_path)
else:
    workbook = Workbook()
    sheet = workbook.active

    headers = ["Item", "Price in ₱", "Date of Purchase"]
    for col_num, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_num).value = header

    footer_labels = ["Total items: ", "Total price: "]
    for col_num, label in enumerate(footer_labels, start=1):
        sheet.cell(row=3, column=col_num).value = label

    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["B"].width = 30
    sheet.column_dimensions["C"].width = 30

    workbook.save(file_path)
    
# Main window stuff 
main_window = ctk.CTk()
ctk.set_appearance_mode("dark")
main_window.title("Tkinter GUI")
screen_width = main_window.winfo_screenwidth()
screen_height = main_window.winfo_screenheight()
window_width = 500
window_height = 330
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
main_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
main_window.resizable(False, False)

# Widget stuff
main_frame = ctk.CTkFrame(master=main_window)
main_frame.pack(pady=15, padx=20, fill="both", expand=True)
frame_label = ctk.CTkLabel(master=main_frame, text="Expense Tracker", font=("Roboto", 15))
frame_label.pack(pady=10, padx=13)

# Product info stuff
item_frame = ctk.CTkFrame(master=main_frame)
item_frame.pack(pady=10, padx=13)
item_frame.grid_columnconfigure(0, weight=1)
item_frame.grid_columnconfigure(1, weight=1)
item_frame.grid_columnconfigure(2, weight=1)

item_name_entry_box = ctk.CTkEntry(master=item_frame, placeholder_text="Item", placeholder_text_color="white")
item_name_entry_box.grid(row=0, column=0, padx=13, pady=10, sticky="nsew")

item_price_entry_box = ctk.CTkEntry(master=item_frame, placeholder_text="Price", placeholder_text_color="white")
item_price_entry_box.grid(row=0, column=1, padx=13, pady=10, sticky="nsew")
validate_cmd = (main_window.register(validate_numeric_input), '%P')
item_price_entry_box.configure(validate='key', validatecommand=validate_cmd)


start_date = datetime.now().date()
end_date = start_date.replace(month=12, day=31)  
delta = timedelta(days=1) 
date_list = [] 
current_date = start_date
while current_date <= end_date:
    date_list.append(current_date.strftime("%B %d"))
    current_date += delta

date_combobox = ctk.CTkComboBox(master=item_frame, values=date_list, state="readonly")
date_combobox.grid(row=0, column=2, padx=13, pady=10, sticky="nsew")
date_combobox.set("Select Date")

# Enter information 
enter_button = ctk.CTkButton(master=main_frame, text="Enter", command=on_button_click)
enter_button.pack(pady=5, padx=13)
message_label = ctk.CTkLabel(master=main_frame, text="")
message_label.pack(pady=3, padx=4)

# Monthly report 
report_label = ctk.CTkLabel(master=main_frame, text="Monthly report:")
report_label.pack(pady=5, padx=13, anchor="w")

report_frame = ctk.CTkScrollableFrame(master=main_frame)
report_frame.pack(pady=8, padx=13, fill="both", expand=True)
last_day_of_month = (start_date.replace(day=1, month=start_date.month + 1) - timedelta(days=1)).day
text_to_display = ""

new_workbook = load_workbook(file_path)
new_worksheet = new_workbook["Sheet"]
item_column = new_worksheet["A"]
number_of_items = 0
price_column = new_worksheet["B"]
total_price = 0
column_letter = 'A'


if start_date.day == last_day_of_month:
    number_of_items = (sum(1 for cell in item_column if cell.value is not None)) - 2
    column_letter = 'A'
    column = new_worksheet[column_letter]

    column_sum = 0
    for cell in column:
        if isinstance(cell.value, (int, float)):
            column_sum += cell.value

    text_to_display = f"By the end of the month, {start_date}, you bought {number_of_items} items and spent ₱{column_sum} within this month."
else:
    text_to_display = "It's not end of the month yet."

text_label = ctk.CTkLabel(master=report_frame, text=text_to_display, wraplength=340, justify="center")
text_label.pack(pady=10, padx=13)

main_window.mainloop()