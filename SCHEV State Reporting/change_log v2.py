# James Caldwell, Fall 2026

import pandas as pd
import numpy as np  
import os
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import Font
from dotenv import load_dotenv

from docx import Document


def load_sheets_with_dynamic_header(file_path):
# This function loads all sheets from an Excel file, checking the first two rows for the header row.
# It skips sheets named 'summary', 'changelog', 'warnings', 'errors', and 'vcsin xref'.
# If a sheet cannot be read, it is skipped with a warning printed to the console.

    result = {}
    meta_data = {}
    excel_file = pd.ExcelFile(file_path)
    for sheet_name in excel_file.sheet_names:
        try:
            
            if sheet_name.strip().lower() not in {'summary', 'changelog', 'warnings', 'errors', 'vcsin xref','comparison summary','missing vcsin','all students'}:
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

def create_all_student_errors_sheet(all_students_df,writer):
    # Create hyperlink for error
    all_students_df['Error Link'] = '=HYPERLINK("#' + all_students_df['Error'] + '!A1", "' + all_students_df['Error'] + '")'

    # Pivot so one row per student, and the list of errors next to each
    student_errors = (
    all_students_df.groupby('Student System ID')['Error Link']
        .agg(list)
        .to_dict()
    )

    # Sort students by number of errors, most first
    student_errors = dict(
        sorted(student_errors.items(), key=lambda x: len(x[1]), reverse=True)
    )

    # Create a dataframe with 1 column for each error, one row per student
    student_errors_df = pd.DataFrame([
        {
            'Student System ID': student,
            # 'Errors': ', '.join(errors)
            **{f'Error {i+1}': error for i, error in enumerate(errors)}
        }
        for student, errors in student_errors.items()
    ])
    
    student_errors_df.to_excel(
        writer,
        sheet_name="All Students",
        index=False
    )

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
    missing_academic_plan_mapping = []
    warning_list = []
    warning_status_list = []
    all_students_df = pd.DataFrame()

    for sheet in all_sheet_names:
        # If either sheet is missing from V1 or V2, df1 or df2 will be None
        df1 = xl1.get(sheet)
        df2 = xl2.get(sheet)
    
        # Drop FAKeyint column. This SCHEV internal code changes between submissions
        col_to_ignore = ['FAKeyint','Rowid','rowid','DateStamp','Errdate','New Error?','FA.Primary Academic Program','FA.Academic Plan'] 
        # this could probably be a function if I want to clean this up in the future
        if df1 is not None:
            df1 = df1.drop(columns=[c for c in col_to_ignore if c in df1.columns])
            df1["Error/Warning Status"] = 'Resolved'
            # Find all columns containing "comments"
            try:
                comment_cols_1 = [c for c in df1.columns.tolist() if ('comments' in c.lower() or 'error/warning status' in c.lower() )]
            except:
                print('error sheet: ' + sheet)
        if df2 is not None:
            df2 = df2.drop(columns=[c for c in col_to_ignore if c in df2.columns])
            df2["Error/Warning Status"] = 'Active'
            # Find all columns containing "comments"
            if sheet == 'BFA003E03':
                print(str(sheet) + 'here')
                print(df2.head(10))
                comment_cols_testing = [c for c in df2.columns.tolist() if ('comments' in c.lower() or 'error/warning status' in c.lower())]

                print(comment_cols_testing)
                print('made past here')
            
            
            comment_cols_2 = [c for c in df2.columns.tolist() if ('comments' in c.lower() or 'error/warning status' in c.lower() )]

            # Collect students who have a missing VCSIN mapping
            if 'Student System ID' in df2.columns:
                missing_VCSIN.extend(
                    df2.loc[df2['Student System ID'].isna(), 'SOCSEC1'].tolist()
                )
                # Also collect all SSIDs to log for a summary sheet
                temp_df = pd.DataFrame({'Student System ID': df2['Student System ID'],'Error': sheet})
                all_students_df = pd.concat([all_students_df, temp_df], ignore_index=True)

        else: # Error is resolved, just pull in previous data for history
            print('Missing df2 is sheet ' + str(sheet))
            df2 = df1
            xl2_meta[sheet]=xl1_meta[sheet]

        comment_cols = comment_cols_1 + comment_cols_2

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

        # Lookup and join Academic plan and program on SSID
        if 'Student System ID' in df.columns.to_list():
            df = pd.merge(
                        df,
                        program_mapping_df[['Student System ID','FA.Primary Academic Program','FA.Academic Plan']],
                        on='Student System ID',
                        how='left',
                        suffixes=('', '_assigned'),
                        indicator=True
                    )

            # For next iteration, save this info into a new sheet.
            unmatched = df[df['_merge'] == 'left_only']
            unmatched = unmatched[unmatched['Student System ID'].notna()]
            if not unmatched.empty:
                print(f'Unmatched Student System IDs in {sheet}:')
                print(unmatched['Student System ID'].to_list())
                missing_academic_plan_mapping.extend(unmatched['Student System ID'].to_list())
            df = df.drop(columns='_merge')

        df.to_excel(writer, sheet_name=sheet, index=False,startrow=3) 
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

    # Add a note if explanation is approved
    try:
        load_dotenv()
        filepath= os.getenv('summary_word_doc')
        doc = Document(filepath)
        matching_errors = []
        for table in doc.tables:
            for row in table.rows:
                if len(row.cells) >= 2:
                    error = row.cells[0].text.strip()
                    explanation = row.cells[1].text.strip()
                    if "ok to pass" in explanation.lower():
                        matching_errors.append(error)
    except:
        print('Summary Word Doc errors')

    comparison_summary.loc[
        comparison_summary['Error Link'].str.extract(r'HYPERLINK\("#([^!]+)!')[0].isin(matching_errors),
        'Explanation'
    ] = 'Explanation Approved'

    # If Active = 0, mark as resolved
    comparison_summary.loc[
        comparison_summary['Status'].str.contains('Active: 0', na=False),
        'Explanation'
    ] = 'Resolved'

    comparison_summary.to_excel(
        writer,
        sheet_name="Comparison Summary",
        index=False
    )

    create_all_student_errors_sheet(all_students_df, writer) 

    # Write missing VCSIN
    missing_VCSIN = list(dict.fromkeys(missing_VCSIN)) # unduplicated list of missing VICSINs
    if missing_VCSIN is not None:
        pd.DataFrame({"Missing VCSIN": missing_VCSIN}).to_excel(writer, sheet_name="Missing VCSIN", startrow=1, index=False)

    writer.close()

    return warning_df, all_students_df

def summary_page(v1_file_path, result_path, warning_df):
    summary_df_old =pd.read_excel(v1_file_path, sheet_name='Comparison Summary')

    summary_df_new['Error'] = summary_df_new.iloc[:, 0].str.extract(r"'([^']*)'")
    summary_df_old['Error'] = summary_df_new.iloc[:, 0].str.extract(r"'([^']*)'")
    summary_df_new = pd.merge(summary_df_new, summary_df_old[['Assigned To:','Error']], how='left',on='Error')

    summary_df_new = pd.merge(summary_df_new, warning_df[['Error Code','Warning']], how='left',right_on='Error Code', left_on='Error').drop(columns=['Error Code'])

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
        if ws.title == sheet_name: # Highlight summary sheet column of links
            for cell in ws["B"][1:]:
                cell.font = Font(color="0000FF", underline="single")
        elif ws.title == 'All Students': # Highligh all hyperlinks on All Students
            for col in ws.iter_cols(min_col=2, max_col=26, min_row=2):
                for cell in col:
                    cell.font = Font(color="0000FF", underline="single")
        else:
        # if ws.title != sheet_name:
            # Add hyperlink in cell B3 for error sheets
            ws.cell(row=3, column=2).value = f'=HYPERLINK("#\'{sheet_name}\'!A1", "{sheet_name}")'
            #style as hyperlink (blue & underlined)
            ws.cell(row=3, column=2).font = Font(color="0000FF", underline="single")

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

    warning_df, all_students_df = change_log(file_path1, file_path2 ,save_name)

    # summary_page(file_path1,save_name, warning_df)
    add_summary_sheet_hyperlinks(save_name)

    print(f"Comparison complete.")

