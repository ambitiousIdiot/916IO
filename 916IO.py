import tkinter
from tkinter import messagebox
from tkinter import filedialog
import customtkinter
import pandas as pd

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("375x350")  # Adjusted for better visual layout

# Dynamically generate the data_arrays structure
print("Generating data arrays...")
num_slots = 4  # As per your current design

card_types = {
    "16 Digital Input": ("I", 16, 1),
    "16 Digital Output": ("O", 16, 2),
    "32 Digital Input": ("I", 32, 1),
    "32 Digital Output": ("O", 32, 2)
}

data_arrays = {}
selector_counters = {1: 0, 2: 0}

for slot in range(1, num_slots + 1):
    for card_name, (prefix, count, selector_num) in card_types.items():
        key = f"{card_name} (Slot {slot})"
        data_arrays[key] = [(f"{{::[P01]local:{slot}:{prefix}.Data.{i}}}", selector_num) for i in range(count)]

def generate_rows_based_on_cards(cards):
    rows = []

    # Reset the selector_counts every time this function is called
    selector_counts = {
        "I": 0,  # For Digital Inputs
        "O": 0,  # For Digital Outputs
    }

    for card_num, card_type in enumerate(cards, start=1):
        if card_type == "No Card":
            continue  # Skip this iteration if "No Card" is selected

        col_name = f"{card_type} (Slot {card_num})"
        values_array = data_arrays.get(col_name, [])
        prefix = card_types[card_type][0]  # Correctly retrieve the prefix

        for i, (en_us_entry, selector_num) in enumerate(values_array):
            en_us = en_us_entry.replace("{::[P01]local:1", f"{{::[P01]local:{card_num}")
            desc = f"ControlListSelector{selector_num}.{selector_counts[prefix]}.St_Caption"
            
            # Increment the selector count for the specific type (I or O)
            selector_counts[prefix] += 1

            # Update the en-US format with card number
            en_us_formatted = f"/*N:1 {en_us} NOFILL DP:0*/   local:{card_num}:{en_us.split(':')[-1]}  {prefix}{card_num}/{i}"
            
            rows.append({"Server": server_name, "Component Type": "Graphic Display", "Component Name": "M916_IO", "Description": desc, "en-US": en_us_formatted})

    print(f"Generated {len(rows)} rows.")
    return rows

def process_excel_file(file_path):
    global server_name

    try:
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
        new_rows = generate_rows_based_on_cards(cards)

        # Appending new rows to the dataframe
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

        # Save the modified dataframe back to the same Excel file
        df.to_excel(file_path, index=False)

        print(f"Processed file: {file_path}")
        
        # Success message
        messagebox.showinfo("Success", "File processed successfully!")
        
    except Exception as e:
        # Error message
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

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