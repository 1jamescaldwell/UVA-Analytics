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
    meta_data = {}
    excel_file = pd.ExcelFile(file_path)
    for sheet_name in excel_file.sheet_names:
        try:
            
            if sheet_name.strip().lower() not in {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref'}:

                # result[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, header=2) # Default to header on row 3 (index 2)
                preview = excel_file.parse(sheet_name=sheet_name, header=2)
                result[sheet_name] = preview #.iloc[header_row + 1:].reset_index(drop=True)
                meta_data[sheet_name] = excel_file.parse(sheet_name=sheet_name, nrows=2)

                # if file_path == "X:/SCHEV/2425 (2024-2025)/Error & Warning Reports/Download and Iteration Check scripts/V2.xlsx":
                #     print(f'Preview of {sheet_name}:')
                #     print(preview.head())
                # preview = excel_file.parse(sheet_name=sheet_name, header=None)

                # # Check first two rows for "Repyear"
                # if "Repyear" in preview.iloc[1].values:
                #     header_row = 1
                # else:
                #     header_row = 0
                
                # header_row = 0

                # preview.columns = preview.iloc[header_row]
                # df = preview.iloc[header_row + 1:].reset_index(drop=True)

                # # Drop 'Notes' column if present
                # # df = df.loc[:, df.columns.str.lower() != "notes"] # Decided to keep Notes column and deal with differences during comparisons
                # result[sheet_name] = df

        except Exception as e:
            print(f"Skipped sheet '{sheet_name}' due to error: {e}")
            continue

    return result, meta_data

def select_file(title):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title=f"Select {title} Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    return file_path

def collect_notes_columns(df):
    """
    This function checks if any column in the DataFrame contains 'notes' (case insensitive).
    If found, it returns a DataFrame with those columns.
    """
    mask = df.columns.str.lower().str.contains('notes')
    matching_positions = mask.nonzero()[0].tolist()  # Get positions of matching columns
    notes_df = pd.DataFrame()
    i = 1  # Start numbering notes columns from 1
    for pos in matching_positions:
        notes_df[f'Notes_{i}'] = df.iloc[:, pos]
        notes_df['SOCSEC1'] = df['SOCSEC1']
        i += 1
    
    return notes_df if not notes_df.empty else None

def change_log(file1_path, file2_path, result_path):
    # SKIP_SHEETS = {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref'}

    # Load all sheets, using header on row 2 (index 1)  
    xl1,xl1_meta = load_sheets_with_dynamic_header(file1_path)
    xl2,xl2_meta = load_sheets_with_dynamic_header(file2_path)

    file_name1 = file1_path.split('/')[-1][:-5] # Remove .xlsx and filepath
    file_name2 = file2_path.split('/')[-1][:-5]

    writer = pd.ExcelWriter(result_path, engine='openpyxl')
    summary = [f'V1 file: {file_name1}', f'V2 file: {file_name2}']
    pd.DataFrame({"Summary": summary}).to_excel(writer, sheet_name="Comparison Summary", index=False)
    # Add SCHEV summary sheet from the second file (latest file)
    SCHEV_summary_from_website = pd.read_excel(file2_path, sheet_name='Summary', header=None)
    SCHEV_summary_from_website.to_excel(writer, sheet_name="SCHEV Error Summary", index=False)
    
    # Get union of all sheet names
    all_sheet_names = set(xl1.keys()).union(set(xl2.keys()))
    all_sheet_names = sorted(all_sheet_names)  # Sort sheet names for consistent order

    print("xl1.keys():", xl1.keys())
    print("xl2.keys():", xl2.keys())


    for sheet in all_sheet_names:

        # If either sheet is missing from V1 or V2, df1 or df2 will be None
        df1 = xl1.get(sheet)
        df2 = xl2.get(sheet)

        if df1 is not None and df2 is not None:

            # Rename NaN column names to "missing"
            df1.columns = ["missing col name" if pd.isna(col) else col for col in df1.columns]
            df2.columns = ["missing col name" if pd.isna(col) else col for col in df2.columns]
            
            # Drop FAKeyint column
            col_to_ignore = 'FAKeyint'
            if col_to_ignore in df1.columns:
                df1 = df1.drop(columns=[col_to_ignore])
            if col_to_ignore in df2.columns:
                df2 = df2.drop(columns=[col_to_ignore])
            # Standardize column order
            df1_sorted = df1[sorted(df1.columns)]
            df2_sorted = df2[sorted(df2.columns)]
            # Sort by 2nd column, which is usually SOCSEC1 
            df1_sorted = df1_sorted.sort_values(by=df1_sorted.columns[1]) 
            df2_sorted = df2_sorted.sort_values(by=df2_sorted.columns[1])

            # Remove 'Comments' columns if they exist, and collect them separately to be added back later after comparison
            if df1_sorted.columns.str.lower().str.contains('comments').any() or df2_sorted.columns.str.lower().str.contains('comments').any(): # Check if any column name contains 'notes' (Notes, notes, notes:, etc)
                notes1 = collect_notes_columns(df1_sorted)
                notes2 = collect_notes_columns(df2_sorted)
                df1 = df1_sorted.drop(columns=[col for col in df1_sorted.columns if 'comments' in col.lower()])
                df2 = df2_sorted.drop(columns=[col for col in df2_sorted.columns if 'comments' in col.lower()])
                # print(f'Notes columns found in {sheet}.')
            if df1_sorted.equals(df2_sorted): # If files are identical
                # pd.DataFrame([["No differences found."]]).to_excel(writer, sheet_name=sheet, index=False, header=False)
                df1.to_excel(writer, sheet_name=sheet, index=False) # Write the original DataFrame to the output file, since it still has the Notes columns
                summary.append(f"Sheet '{sheet}': Exists in both files, no differences.")
            else:
                df1_tagged = df1_sorted.copy()
                df1_tagged["__source__"] = "File 1"

                df2_tagged = df2_sorted.copy()
                df2_tagged["__source__"] = "File 2"

                combined = pd.concat([df1_tagged, df2_tagged], ignore_index=True)
                combined["__duplicated__"] = combined.duplicated(subset=combined.columns.difference(["__source__"]), keep=False)

                def classify(row):
                    if row["__duplicated__"]:
                        return "Both"
                    return row["__source__"]

                combined["New Error?"] = combined.apply(classify, axis=1)
                output_df = combined.drop(columns=["__source__", "__duplicated__"])

                # Drop duplicate rows, which occurs if there are identical rows in both files. Only keep one instance.
                output_df = output_df.drop_duplicates().reset_index(drop=True)

                # add back the two rows of metadata (Error description and link) from the first file if they exist
                # output_df = pd.concat([xl2_meta[sheet], output_df], ignore_index=True) 
                # Add notes columns back in if they exist
                try:
                    if notes1 is not None:
                        output_df = output_df.merge(notes1, how='left', on='SOCSEC1')
                    if notes2 is not None:
                        output_df = output_df.merge(notes2, how='left', on='SOCSEC1')
                except:
                    output_df['Notes'] = 'Error merging notes columns from python script. Check manually.'
                    print(f'Error merging notes columns in {sheet}.')
                output_df.to_excel(writer, sheet_name=sheet, index=False,startrow=2) # Leave two rows at top for metadata
                
                # Add back the two rows of metadata (Error description and link) from the first file
                ws = writer.sheets[sheet]
                ws.cell(row=1, column=1, value=xl2_meta[sheet].columns[0]) # Error description
                ws.cell(row=2, column=2, value=xl2_meta[sheet].iloc[0, 1]) # Link

                only_in_file1 = (combined["New Error?"] == "File 1").sum()
                only_in_file2 = (combined["New Error?"] == "File 2").sum()
                in_both_files = (combined["New Error?"] == "Both").sum()
                if only_in_file1 == 0 and only_in_file2 == 0:
                    # pd.DataFrame([["No differences found."]]).to_excel(writer, sheet_name=sheet, index=False, header=False)
                    df1.to_excel(writer, sheet_name=sheet, index=False,startrow=2)
                    summary.append(f"Sheet '{sheet}': Exists in both files, no differences.")
                else:
                    summary.append(f"Errors {sheet} exists in both files. "
                                f"{only_in_file1} rows only in '{file_name1}', {only_in_file2} rows only in '{file_name2}', and {in_both_files} rows in both files.")
        elif df1 is not None:
            # pd.DataFrame([["This sheet only exists in File 1."]]).to_excel(writer, sheet_name=sheet, index=False, header=False) # Uncomment if you want an entire sheet that says this
            summary.append(f"Error {sheet} no longer present in latest file.")
        elif df2 is not None:
            df2.to_excel(writer, sheet_name=sheet, index=False) # Write the original DataFrame to the output file, since it still has the Notes columns
            summary.append(f"Error {sheet} added to latest file.")

    # Write summary
    pd.DataFrame({"Summary": summary}).to_excel(writer, sheet_name="Comparison Summary", index=False)
    writer.close()
    print(f"Comparison complete. Output saved to: {result_path}")

if __name__ == "__main__":

    file_path1 = select_file('first')
    file_path2 = select_file('second')
    dir = os.path.dirname(file_path1)
    os.chdir(dir)
    file_name1 = file_path1.split('/')[-1]
    file_name2 = file_path2.split('/')[-1]
    save_name = f'{file_name1[:-5]}_{file_name2[:-5]}.xlsx'

    change_log(file_path1, file_path2 ,save_name)


