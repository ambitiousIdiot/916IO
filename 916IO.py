import tkinter
from tkinter import filedialog
import customtkinter
import pandas as pd

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("600x450")  # Adjusted for better visual layout

def generate_rows_based_on_cards_updated(cards):
    copypasta_df = pd.read_csv("916DigitalCopypasta.csv")

    rows = []

    for card_num, card_type in enumerate(cards, start=1):
        if card_type.startswith("32 Digital"):
            iterations = 32
        elif card_type.startswith("16 Digital"):
            iterations = 16
        else:
            iterations = 0  # For "No Card" option or other card types

        col_name = f"{card_type} (Slot {card_num})"

        for i in range(iterations):
            if col_name in copypasta_df.columns:
                en_us = copypasta_df[col_name].iloc[i].replace("{::[P01]local:1", f"{{::[P01]local:{card_num}")
            else:
                en_us = ""
            desc = f"ControlListSelector{card_num}.{i}.St_Caption"
            rows.append({"Server": server_name, "Component Type": "Graphic Display", "Component Name": "M916_IO", "Description": desc, "en-US": en_us})

    return rows

def process_excel_file(file_path):
    global server_name

    df = pd.read_excel(file_path)

    # Storing server name from the original data for further use
    server_name = df["Server"].iloc[0]

    # Identify rows that match the given conditions using the provided column names
    rows_to_delete = df[(df['Component Name'].str.contains("M916_IO", na=False)) & 
                        (df['Description'].str.contains("ControlListSelector", na=False))]

    # Drop these rows
    df = df.drop(rows_to_delete.index)

    # Generating rows based on card selections
    cards = [menu.get() for menu in menus]
    new_rows = generate_rows_based_on_cards_updated(cards)

    # Appending new rows to the dataframe
    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # Save the modified dataframe back to the same Excel file
    df.to_csv(file_path, index=False)
    print(f"Processed file: {file_path}")

def select_excel_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xls;*.xlsx")])
    if file_path:
        process_excel_file(file_path)

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Please select card types for each slot:")
label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

options = ["No Card", "16 Digital Input", "16 Digital Output", "32 Digital Input", "32 Digital Output"]

menus = []
for i in range(4):
    card_label = customtkinter.CTkLabel(master=frame, text=f"Card {i+1}")
    card_label.grid(row=i+1, column=0, padx=10, pady=10, sticky="e")

    menu = customtkinter.CTkOptionMenu(master=frame, values=options)
    menu.grid(row=i+1, column=1, padx=10, pady=10, sticky="w")
    menus.append(menu)

button = customtkinter.CTkButton(master=frame, text="Select Excel File", command=select_excel_file)
button.grid(row=5, column=0, columnspan=2, pady=20, padx=10)

root.mainloop()