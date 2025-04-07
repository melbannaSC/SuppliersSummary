import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_supplier_and_total(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df = df.dropna(how='all').dropna(axis=1, how='all')

    supplier_row = df[df.astype(str).apply(lambda x: x.str.contains("Supplier Code", case=False, na=False)).any(axis=1)]
    supplier_name = supplier_row.iloc[0, 0].split(":")[-1].strip() if not supplier_row.empty else sheet_name

    if sheet_name == xls.sheet_names[-1]:
        df = df[:-2]
    last_row = df.iloc[-1, :].dropna()
    total_value = last_row.values[-1] if not last_row.empty else None
    total_value = round(float(total_value), 2) if total_value is not None else None

    return {"Supplier Name": supplier_name, "Total": total_value}

def process_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        return

    xls = pd.ExcelFile(file_path)
    summary_data = [extract_supplier_and_total(xls, sheet) for sheet in xls.sheet_names]
    summary_df = pd.DataFrame(summary_data)

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        summary_df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"Summary saved at:\n{save_path}")

# GUI
root = tk.Tk()
root.withdraw()  # Hide the root window
process_file()