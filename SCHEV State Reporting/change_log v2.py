# James Caldwell, Fall 2025

import pandas as pd
import numpy as np  
import os
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import Font
from dotenv import load_dotenv


def load_sheets_with_dynamic_header(file_path):
# This function loads all sheets from an Excel file, checking the first two rows for the header row.
# It skips sheets named 'summary', 'changelog', 'warnings', 'errors', and 'vcsin xref'.
# If a sheet cannot be read, it is skipped with a warning printed to the console.

    result = {}
    meta_data = {}
    excel_file = pd.ExcelFile(file_path)
    for sheet_name in excel_file.sheet_names:
        try:
            
            if sheet_name.strip().lower() not in {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref','comparison summary','Missing VCSIN'}:

                # result[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, header=2) # Default to header on row 3 (index 2)
                preview = excel_file.parse(sheet_name=sheet_name, header=3)
                result[sheet_name] = preview #.iloc[header_row + 1:].reset_index(drop=True)
                meta_data[sheet_name] = excel_file.parse(sheet_name=sheet_name, nrows=3)

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

def collect_notes_columns(df,sheet_name):
    """
    This function checks if any column in the DataFrame contains 'comments' (case insensitive).
    If found, it returns a DataFrame with those columns.
    """

    comment_cols = df[[col for col in df.columns if 'comments' in col.lower()]]

    # first try to add SOCSEC1 and Rowid if they exist. Error count will be 2 if both are missing
    try:
        notes_df = comment_cols.copy()
        notes_df['SOCSEC1'] = df['SOCSEC1'].copy()
    except:
        print('Error finding comments for sheet: ' + sheet_name + 'col names: ' + str(df.columns))
    return notes_df if not notes_df.empty else None

def change_log(file1_path, file2_path, result_path):

    # Load all sheets, using header on row 2 (index 1)  
    xl1,xl1_meta = load_sheets_with_dynamic_header(file1_path)
    xl2,xl2_meta = load_sheets_with_dynamic_header(file2_path)

    # print(xl2_meta)

    file_name1 = file1_path.split('/')[-1][:-5] # Remove .xlsx and filepath
    file_name2 = file2_path.split('/')[-1][:-5]

    writer = pd.ExcelWriter(result_path, engine='openpyxl')
    summary = []#[f'V1 file: {file_name1}', f'V2 file: {file_name2}']

    pd.DataFrame({"Summary": summary}).to_excel(writer, sheet_name="Comparison Summary", index=False)

    # Load program plan mappings, to be merged with each error/warning that has SSID as a column
            # This Adds the latest student program and plan to sheets that have SSIDs
    print('Loading plan mappings...')
    load_dotenv()
    program_plan_mapping= os.getenv('program_plan_mapping')
    program_mapping_df = pd.read_excel(program_plan_mapping)
        # Remove rows that have '-' for both columns of program/plan
    cols = ['FA.Primary Academic Program', 'FA.Academic Plan']
    program_mapping_df = program_mapping_df[~(program_mapping_df[cols] == '-').all(axis=1)]
    program_mapping_df.sort_values(by=['Student System ID','Term'],ascending=[True,False], inplace=True)
    print('mapping sorting is:')
    program_mapping_df = program_mapping_df.drop_duplicates(subset=['Student System ID'],keep='first')
    print('loaded')

    # Get union of all sheet names
    all_sheet_names = set(xl1.keys()).union(set(xl2.keys()))
    all_sheet_names = sorted(all_sheet_names)  # Sort sheet names for consistent order
    if 'Missing VCSIN' in all_sheet_names:
        all_sheet_names.remove('Missing VCSIN')

    meta_df_created = []
    missing_VCSIN = []
    warning_list = []
    warning_status_list = []

    for sheet in all_sheet_names:
        # If either sheet is missing from V1 or V2, df1 or df2 will be None
        df1 = xl1.get(sheet)
        df2 = xl2.get(sheet)

        if sheet == "BFA010W10":  #fix SCHEV table error 
            df1 = df1.rename(columns={'Note2': 'SOCSEC1'})
            df2 = df2.rename(columns={'Note2': 'SOCSEC1'})

        # Collect students who have a missing VCSIN mapping
        if 'Student System ID' in df2.columns:
            missing_VCSIN.extend(
                df2.loc[df2['Student System ID'].isna(), 'SOCSEC1'].tolist()
            )
    
        # Drop FAKeyint column. This SCHEV internal code changes between submissions
        # col_to_ignore = ['FAKeyint','Missing VCSIN']
        col_to_ignore = ['FAKeyint','Rowid','DateStamp','Errdate','New Error?']
        if df1 is not None:
            df1 = df1.drop(columns=[c for c in col_to_ignore if c in df1.columns])
        if df2 is not None:
            df2 = df2.drop(columns=[c for c in col_to_ignore if c in df2.columns])

        # This can be deleted after testing/building done. was only a v1 thing i think.
        if 'SSID' in df1.columns:
            df1 = df1.rename(columns={'SSID': 'Student System ID'})
        if 'SSID' in df2.columns:
            df1 = df2.rename(columns={'SSID': 'Student System ID'})

        df1["Error/Warning Status"] = 'Resolved'
        df2["Error/Warning Status"] = 'Active'
        # I want to add a 3rd one that says "Active again after being resolved previously. But not for now.."

        # Find all columns containing "comments"
        comment_cols = [c for c in df1.columns.tolist() + df2.columns.tolist()
                        if ('comments' in c.lower() or 'error/warning status' in c.lower() )]

        # Stack them
        df = pd.concat([df1, df2], ignore_index=True, sort=False)

        # Sort by Active then Resolved. If an row is in the newer df2, then it will be active and the resolved will be dropped in the groupby statement below
        df["Error/Warning Status"] = pd.Categorical(
        df["Error/Warning Status"],
        categories=['Active', 'Resolved'],
        ordered=True
        )
        df = df.sort_values("Error/Warning Status")

        # Everything except comments is used to identify a unique row
        match_cols = [c for c in df.columns if c not in comment_cols]

        # Combine duplicate rows, keeping comments from each source
            # This drops duplicate rows
        df = df.groupby(match_cols, dropna=False, as_index=False).first()

        # Put Status in the first column
        col = df.pop("Error/Warning Status")
        df.insert(0, "Error/Warning Status", col)

        
        # # need to move this outside the loop
        # load_dotenv()
        # program_plan_mapping= os.getenv('program_plan_mapping')
        # program_mapping_df = pd.read_excel(program_plan_mapping)
        # # print(program_mapping_df.head())
        # # print(df.head(10))

        # # Add the latest student program and plan to sheets that have SSIDs
        # program_mapping_df.sort_values(by=['Student System ID','Term'],ascending=[True,False], inplace=True)
        # print('mapping sorting is:')
        # print(program_mapping_df.head(15))
        # program_mapping_df.drop_duplicates(subset=['Student System ID'],keep='first')

        # Lookup and join Academic plan and program on SSID
        if 'Student System ID' in df.columns.to_list():
            print('yay!') 
            df = pd.merge(
                        df,
                        program_mapping_df[['Student System ID','FA.Primary Academic Program','FA.Academic Plan']],
                        on='Student System ID',
                        how='left',
                        suffixes=('', '_assigned')
                    )
            # df = df.drop(columns='Student System ID_assigned')
            print(sheet)
            print('added mapping:')
            print(df.head(10))

        df.to_excel(writer, sheet_name=sheet, index=False,startrow=3) 
        # summary.append(f"Error '{sheet}' present in latest file. No SSN column to compare or duplicate SSN rows, so no summary/comparison stats.")
        # Add back the two rows of metadata (Error description and link) from the first file
        ws = writer.sheets[sheet]
        ws.cell(row=1, column=1, value=xl2_meta[sheet].columns[0]) # Error description
        ws.cell(row=2, column=2, value=xl2_meta[sheet].iloc[0, 1]) # Definition
        ws.cell(row=3, column=5, value=xl2_meta[sheet].iloc[1, 2]) # Link
        # warning_list.append(xl2_meta[1])
        warning_list.append(xl2_meta[sheet].columns[0])

        resolved_counts = df['Error/Warning Status'].value_counts()
        # warning_status_list.append(str(resolved_counts))
        warning_status_list.append(f"Active: {resolved_counts.get('Active', 0)}, Resolved: {resolved_counts.get('Resolved', 0)}")

        # except:
            # print(str(sheet) + ' had errors')

    warning_df = pd.DataFrame(warning_list, columns=['Warning'])
    # warning_df['Error Code'] = warning_df['Warning'].str.split(':').str[0] # Add column with just the error code from the first :'
    warning_df['Error Link'] = warning_df['Warning'].str.split(':').str[0].apply(lambda x: f'=HYPERLINK("#{x}!A1", "{x}")')

    # Collect assigned to column
    summary_df_old =pd.read_excel(file1_path, sheet_name='Comparison Summary')
    assigned_to_df =  summary_df_old[['Assigned To:','Error Link']]

    # warning_status_df = pd.DataFrame(warning_status_list, columns=['Status'])

    comparison_summary = pd.DataFrame({
        "Error Link": warning_df['Error Link'],
        "Warning": warning_list,
        "Status": warning_status_list
        # 'Assigned To:': warning_df['Assigned To:']
    })

    # Add the assigned to column from the previous iteration
    comparison_summary = pd.merge(
        comparison_summary,
        assigned_to_df,
        left_on=comparison_summary['Warning'].str.split(':').str[0],
        right_on='Error Link',
        how='left',
        suffixes=('', '_assigned')
    )
    comparison_summary = comparison_summary.drop(columns='Error Link_assigned')

    # Put Assigned To in the first column
    col = comparison_summary.pop("Assigned To:")
    comparison_summary.insert(0, "Assigned To:", col)

    comparison_summary.to_excel(
        writer,
        sheet_name="Comparison Summary",
        index=False
    )

    # Write missing VCSIN
    missing_VCSIN = list(dict.fromkeys(missing_VCSIN)) # unduplicated list of missing VICSINs
    if missing_VCSIN is not None:
        pd.DataFrame({"Missing VCSIN": missing_VCSIN}).to_excel(writer, sheet_name="Missing VCSIN", startrow=1, index=False)

    writer.close()

    return warning_df
    
# def classify_status(text):
#     if 'longer' in str(text).lower():
#         return 'Complete'
#     elif 'added' in str(text).lower():
#         return 'New Error'
#     else:
#         return 'Recurring error'

def summary_page(v1_file_path, result_path, warning_df):
    # summary_df_new =pd.read_excel(result_path, sheet_name='Comparison Summary')
    summary_df_old =pd.read_excel(v1_file_path, sheet_name='Comparison Summary')
    assigned_to =  summary_df_old[['Assigned To:','Error']]

    summary_df_new['Error'] = summary_df_new.iloc[:, 0].str.extract(r"'([^']*)'")
    summary_df_old['Error'] = summary_df_new.iloc[:, 0].str.extract(r"'([^']*)'")
    summary_df_new = pd.merge(summary_df_new, summary_df_old[['Assigned To:','Error']], how='left',on='Error')
    # summary_df_new = pd.merge(summary_df_new, summary_df_old['Error','SCHEV Error Summary Explanations (add notes here if there is an explanation for all records in that error)'], how='left',on='Error')

    load_dotenv()
    meta_data_path= os.getenv('meta_data_path')
    meta_data_df = pd.read_excel(meta_data_path)
    # summary_df_new = pd.merge(summary_df_new, meta_data_df[['Description','ErrCode']], how='left',right_on='ErrCode', left_on='Error').drop(columns=['ErrCode'])
    summary_df_new = pd.merge(summary_df_new, warning_df[['Error Code','Warning']], how='left',right_on='Error Code', left_on='Error').drop(columns=['Error Code'])
    #here

    # warning_df
    # summary_df_new['Status'] = summary_df_new['Summary'].apply(classify_status)

    summary_df_new['Error_Link'] = summary_df_new['Error'].apply(
    lambda x: f'=HYPERLINK("#{x}!A1", "{x}")'
    )

    # print(summary_df_new)
    summary_df_new = summary_df_new[['Assigned To:', 'Error_Link', 'Warning_x']] # dropped summary and status

    summary_df_new.rename(columns={"Warning_x": "Warning"})

    with pd.ExcelWriter(result_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        summary_df_new.to_excel(writer, sheet_name="Comparison Summary", index=False)

    # Load workbook
    wb = load_workbook(result_path)
    ws = wb['Comparison Summary']
    # Apply hyperlink style to the entire "Error_Link" column (assuming column C, adjust as needed)
    for cell in ws['C'][1:]:  # skip header
        cell.font = Font(color="0000FF", underline="single")

    # Save workbook
    wb.save(result_path)

def add_summary_sheet_hyperlinks(result_path):
   # Add hyperlinks to each page that links to the summary sheet
    wb = load_workbook(result_path)

    # Sheet to link to
    sheet_name = "Comparison Summary"

    # Loop through all sheets except the one we are linking to
    for ws in wb.worksheets:
        if ws.title != sheet_name:
            # Add hyperlink in cell B3
            ws.cell(row=3, column=2).value = f'=HYPERLINK("#\'{sheet_name}\'!A1", "{sheet_name}")'
            #style as hyperlink (blue & underlined)
            ws.cell(row=3, column=2).font = Font(color="0000FF", underline="single")
        if ws.title == sheet_name: # Highlight summary sheet column of links
            for cell in ws["B"][1:]:
                cell.font = Font(color="0000FF", underline="single")

    # Save workbook
    wb.save(result_path)

if __name__ == "__main__":

    file_path1 = select_file('first')
    file_path2 = select_file('second')
    dir = os.path.dirname(file_path1)
    os.chdir(dir)
    file_name1 = file_path1.split('/')[-1]
    file_name2 = file_path2.split('/')[-1]
    save_name = f'{file_name1[:-5]}_{file_name2[:-5]}.xlsx'
    if os.path.exists(save_name):
        os.remove(save_name) # Remove existing file if it exists to avoid overwrite issues

    warning_df = change_log(file_path1, file_path2 ,save_name)
    # summary_page(file_path1,save_name, warning_df)
    add_summary_sheet_hyperlinks(save_name)
    print(f"Comparison complete.")

