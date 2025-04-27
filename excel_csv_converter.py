import tkinter as tk
from tkinter import filedialog, messagebox, ttk, BooleanVar
import pandas as pd
import os
from datetime import datetime


# Logging Function
def log_conversion(log_text):
    with open("conversion_log.txt", "a") as log_file:
        log_file.write(f"{datetime.now()} - {log_text}\n")


# Excel to CSV Conversion
def excel_to_csv():
    try:
        excel_files = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not excel_files:
            return
        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            return

        combined_data = pd.DataFrame()

        for idx, excel_file in enumerate(excel_files):
            excel_data = pd.ExcelFile(excel_file)
            filename = os.path.basename(excel_file).split('.')[0]

            for sheet_name in excel_data.sheet_names:
                sheet_data = excel_data.parse(sheet_name)
                output_file = f"{output_folder}/{filename}_{sheet_name}.csv"
                sheet_data.to_csv(output_file, index=False)
                log_conversion(f"Excel to CSV: {output_file}")

                if merge_var.get():
                    combined_data = pd.concat([combined_data, sheet_data], ignore_index=True)

            progress['value'] = ((idx + 1) / len(excel_files)) * 100
            app.update_idletasks()

        if merge_var.get():
            combined_file = f"{output_folder}/combined_output.csv"
            combined_data.to_csv(combined_file, index=False)
            log_conversion(f"Combined CSV created: {combined_file}")

        messagebox.showinfo("Success", "Excel files converted successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# CSV to Excel Conversion
def csv_to_excel():
    try:
        csv_files = filedialog.askopenfilenames(title="Select CSV Files", filetypes=[("CSV Files", "*.csv")])
        if not csv_files:
            return
        output_file = filedialog.asksaveasfilename(title="Save Excel File", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not output_file:
            return

        with pd.ExcelWriter(output_file) as writer:
            for idx, csv_file in enumerate(csv_files):
                sheet_name = os.path.basename(csv_file).replace(".csv", "")[:31]
                csv_data = pd.read_csv(csv_file)
                csv_data.to_excel(writer, index=False, sheet_name=sheet_name)
                log_conversion(f"CSV to Excel Sheet: {sheet_name}")

                progress['value'] = ((idx + 1) / len(csv_files)) * 100
                app.update_idletasks()

        messagebox.showinfo("Success", f"CSV files converted to Excel: {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Toggle Dark Mode
def toggle_dark_mode():
    if dark_mode.get():
        app.configure(bg="#121212")
        frame.configure(style="Dark.TFrame")
        style.configure("TButton", foreground="white", background="#1E1E1E")
        style.configure("TCheckbutton", foreground="white", background="#121212")
        style.configure("TLabel", foreground="white", background="#121212")
    else:
        app.configure(bg="SystemButtonFace")
        frame.configure(style="Light.TFrame")
        style.configure("TButton", foreground="black", background="SystemButtonFace")
        style.configure("TCheckbutton", foreground="black", background="SystemButtonFace")
        style.configure("TLabel", foreground="black", background="SystemButtonFace")


# GUI Setup
app = tk.Tk()
app.title("Excel & CSV Converter")
app.geometry("600x400")
app.resizable(False, False)

style = ttk.Style(app)
style.theme_use("clam")

style.configure("Dark.TFrame", background="#121212")
style.configure("Light.TFrame", background="SystemButtonFace")

frame = ttk.Frame(app, padding=20, style="Light.TFrame")
frame.pack(fill="both", expand=True)

title_label = ttk.Label(frame, text="Excel & CSV Converter", font=("Helvetica", 18, "bold"))
title_label.pack(pady=10)

btn_excel_to_csv = ttk.Button(frame, text="Convert Excel to CSV", command=excel_to_csv)
btn_excel_to_csv.pack(pady=10, fill='x')

btn_csv_to_excel = ttk.Button(frame, text="Convert CSV to Excel", command=csv_to_excel)
btn_csv_to_excel.pack(pady=10, fill='x')

merge_var = BooleanVar()
merge_check = ttk.Checkbutton(frame, text="Merge all Excel Sheets to a single CSV", variable=merge_var)
merge_check.pack(pady=5)

progress = ttk.Progressbar(frame, orient="horizontal", length=500, mode="determinate")
progress.pack(pady=20)

dark_mode = BooleanVar()
dark_check = ttk.Checkbutton(frame, text="Enable Dark Mode", variable=dark_mode, command=toggle_dark_mode)
dark_check.pack(pady=5)

app.mainloop()
