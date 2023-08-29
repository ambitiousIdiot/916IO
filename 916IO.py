import tkinter
from tkinter import filedialog
import customtkinter
import pandas as pd

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("600x450")  # Adjusted for better visual layout

def process_excel_file(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Identify rows that match the given conditions using the provided column names
    rows_to_delete = df[(df['Component Name'].str.contains("M916_IO", na=False)) & 
                        (df['Description'].str.contains("ControlListSelector", na=False))]

    # Drop these rows
    df = df.drop(rows_to_delete.index)

    # Save the modified dataframe back to the same Excel file
    df.to_excel(file_path, index=False)
    print(f"Processed file: {file_path}")

def select_excel_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xls;*.xlsx")])
    if file_path:
        process_excel_file(file_path)

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

# Placing the main label at the top using grid
label = customtkinter.CTkLabel(master=frame, text="Please select card types for each slot:")
label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

# Option values for the dropdown menus
options = ["No Card", "16 Digital Input", "16 Digital Output", "32 Digital Input", "32 Digital Output", "16 Analog Input", "16 Analog Output", "32 Analog Input", "32 Analog Output"]

# Creating labels and dropdown menus for each card
menus = []
for i in range(4):
    card_label = customtkinter.CTkLabel(master=frame, text=f"Card {i+1}")
    card_label.grid(row=i+1, column=0, padx=10, pady=10, sticky="e")
    
    menu = customtkinter.CTkOptionMenu(master=frame, values=options)
    menu.grid(row=i+1, column=1, padx=10, pady=10, sticky="w")
    menus.append(menu)

# Creating the button at the bottom
button = customtkinter.CTkButton(master=frame, text="Select Excel File", command=select_excel_file)
button.grid(row=5, column=0, columnspan=2, pady=20, padx=10)  # Place the button at the bottom

root.mainloop()