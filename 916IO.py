import tkinter
from tkinter import messagebox
from tkinter import filedialog
import customtkinter
import pandas as pd

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("375x350")

card_types = {
    "16 Digital Input": ("I", 16),
    "16 Digital Output": ("O", 16),
    "32 Digital Input": ("I", 32),
    "32 Digital Output": ("O", 32)
}

def generate_all_selectors(server_name):
    rows = []
    for selector_num in range(1, 5):
        for i in range(256):
            desc = f"ControlListSelector{selector_num}.{i}.St_Caption"
            rows.append({
                "Server": server_name,
                "Component Type": "Graphic Display",
                "Component Name": "M916_IO",
                "Description": desc,
                "en-US": ""
            })
    return rows

def generate_rows(cards, server_name, all_rows):
    selector_counts = [0, 0, 0, 0]

    for slot_num, card_type in enumerate(cards, start=1):
        if card_type == "No Card":
            continue

        prefix, count = card_types[card_type]
        selector_num = 1 if "Input" in card_type else 2

        for i in range(count):
            en_us = f"{{::[P01]Local:{slot_num}:{prefix}.Data.{i}}}"
            idx = (selector_num - 1) * 256 + selector_counts[selector_num-1]
            all_rows[idx]["en-US"] = f"/*N:1 {en_us} NOFILL DP:0*/   local:{slot_num}:{en_us.split(':')[-1].split('}')[0]}  {prefix}{slot_num}/{i}"
            
            selector_counts[selector_num-1] += 1

def process_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        server_name = df["Server"].iloc[0]
        df = df[~df['Description'].str.startswith("ControlListSelector")]

        all_rows = generate_all_selectors(server_name)
        cards = [menu.get() for menu in menus]
        generate_rows(cards, server_name, all_rows)
        
        df = pd.concat([df, pd.DataFrame(all_rows)], ignore_index=True)
        df.to_excel(file_path, index=False)
        
        messagebox.showinfo("Success", "File processed successfully!")
    except Exception as e:
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