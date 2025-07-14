import pandas as pd
import numpy as np
import os

import tkinter as tk
from tkinter import filedialog

def load_sheets_with_dynamic_header(file_path):
# This function loads all sheets from an Excel file, checking the first two rows for the header row.
# It skips sheets named 'summary', 'changelog', 'warnings', 'errors', and 'vcsin xref'.
# If a sheet cannot be read, it is skipped with a warning printed to the console.

    result = {}
    excel_file = pd.ExcelFile(file_path)

    for sheet_name in excel_file.sheet_names:
        try:
            
            if sheet_name.strip().lower() not in {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref'}:
                preview = excel_file.parse(sheet_name=sheet_name, header=None)

                # Check first two rows for "Repyear"
                if "Repyear" in preview.iloc[1].values:
                    header_row = 1
                else:
                    header_row = 0

                preview.columns = preview.iloc[header_row]
                df = preview.iloc[header_row + 1:].reset_index(drop=True)

                # Drop 'Notes' column if present
                df = df.loc[:, df.columns.str.lower() != "notes"]
                result[sheet_name] = df

        except Exception as e:
            print(f"Skipped sheet '{sheet_name}' due to error: {e}")
            continue

    return result

def select_file(title):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title=f"Select {title} Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    return file_path

def change_log(file1_path, file2_path, result_path):
    # SKIP_SHEETS = {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref'}

    # Load all sheets, using header on row 2 (index 1)
    xl1 = load_sheets_with_dynamic_header(file1_path)
    xl2 = load_sheets_with_dynamic_header(file2_path)

    writer = pd.ExcelWriter(result_path, engine='openpyxl')
    summary = ['test']
    pd.DataFrame({"Summary": summary}).to_excel(writer, sheet_name="Comparison Summary", index=False) # Write an empty summary sheet first to have this sheet be first in the file

    # Get union of all sheet names
    all_sheet_names = set(xl1.keys()).union(set(xl2.keys()))

    for sheet in all_sheet_names:

        df1 = xl1.get(sheet)
        df2 = xl2.get(sheet)

        if df1 is not None and df2 is not None:
            # Standardize column order
            df1_sorted = df1[sorted(df1.columns)]
            df2_sorted = df2[sorted(df2.columns)]

            if df1_sorted.equals(df2_sorted):
                pd.DataFrame([["No differences found."]]).to_excel(writer, sheet_name=sheet, index=False, header=False)
                summary.append(f"Sheet '{sheet}': Exists in both files, no differences.")
            else:
                df1_tagged = df1.copy()
                df1_tagged["__source__"] = "File 1"

                df2_tagged = df2.copy()
                df2_tagged["__source__"] = "File 2"

                combined = pd.concat([df1_tagged, df2_tagged], ignore_index=True)
                combined["__duplicated__"] = combined.duplicated(subset=combined.columns.difference(["__source__"]), keep=False)

                def classify(row):
                    if row["__duplicated__"]:
                        return "Both"
                    return row["__source__"]

                combined["Comparison Result"] = combined.apply(classify, axis=1)
                output_df = combined.drop(columns=["__source__", "__duplicated__"])
                output_df.to_excel(writer, sheet_name=sheet, index=False)

                only_in_file1 = (combined["Comparison Result"] == "File 1").sum()
                only_in_file2 = (combined["Comparison Result"] == "File 2").sum()
                in_both_files = (combined["Comparison Result"] == "Both").sum()
                summary.append(f"Errors {sheet} exists in both files. "
                               f"{only_in_file1} rows only in File 1, {only_in_file2} rows only in File 2, and {in_both_files} rows in both files.")
        elif df1 is not None:
            # pd.DataFrame([["This sheet only exists in File 1."]]).to_excel(writer, sheet_name=sheet, index=False, header=False) # Uncomment if you want an entire sheet that says this
            summary.append(f"Error {sheet} no longer present in latest file.")
        elif df2 is not None:
            # pd.DataFrame([["This sheet only exists in File 2."]]).to_excel(writer, sheet_name=sheet, index=False, header=False) # Uncomment if you want an entire sheet that says this
            summary.append(f"Error {sheet} added to latest file.")

    # Write summary
    pd.DataFrame({"Summary": summary}).to_excel(writer, sheet_name="Comparison Summary", index=False)
    writer.close()
    print(f"Comparison complete. Output saved to: {result_path}")

if __name__ == "__main__":
    
    # dir = r"C:\Users\ywe4kw\OneDrive - University of Virginia\Documents\3Projects\SCHEV FA change log"
    # os.chdir(dir)
    # file1 = r"v4 (2024-10-02) (dvn3m).xlsx"
    # file2 = r"v9 (2024-10-18) (dvn3m).xlsx"
    # save_name = f'{file1[:-5]}_{file2[:-5]}.xlsx'
    # if not os.path.exists(dir):
    #     print(f"Folder path does not exist: {dir}")
    # change_log(dir + '\\' + file1, dir + '\\' + file2,save_name)

    file_path1 = select_file('first')
    file_path2 = select_file('second')
    file_name = file_path1.split 
    print(f"Selected file: {file_path1}")
