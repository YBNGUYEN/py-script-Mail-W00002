# py-script-Mail-W00002
Application for Mail Automated
import os, psutil, subprocess, openpyxl, yaml, tkinter as tk, pandas as pd
from datetime import datetime
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font, PatternFill, NamedStyle
import win32com.client as win32
from win32com.client import Dispatch
import win32gui
import win32con
import time
#-----------------Step1: Get the basic specific, variable and rules for running file from yaml and data from excel
def select_yaml_file():
    root = tk.Tk()
    root.withdraw() 
    file_path = filedialog.askopenfilename(title="Chọn tệp YAML", filetypes=[("YAML files", "*.yml;*.yaml"), ("All files", "*.*")])
    return file_path
yaml_path = select_yaml_file()  #Address of yaml file path
if os.path.exists(yaml_path):
    with open(yaml_path, 'r', encoding='utf-8') as f:
        try:
            path = yaml.safe_load(f)
            if not isinstance(path, dict): raise ValueError("YAML File is not correct with the basic structure! Please check and correct it!.")
        except yaml.YAMLError as e: raise ValueError(f"YAML is not right: {e}")
        
        if path['day'] == 'today': # process the timeline and time definition in this file
            today_Ymd = datetime.today().strftime('%Y%m%d')
            today_Y = datetime.today().strftime('%Y')
            today_Ym = datetime.today().strftime('%Y%m')
            today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
            ver_date = f"{datetime.now().year}-{datetime.now().strftime('%b')}-{datetime.now().day:02d}"  # Create the version
        else:
            try:
                today = datetime.strptime(path['day'], '%Y%m%d')
                today_Ymd = today.strftime('%Y%m%d')
                today_Ym = today.strftime('%Y%m')
                today_Y = today.strftime('%Y')
                ver_date= today_Ymd
            except ValueError: raise ValueError(f"Invalid date format in YAML: {path['day']}")
        #Some parameter need to fixed to ymal file and #--------parameter and variable
        Root_FolderName = path['Recording_FolderName']# folder name for recording historical data
        ListMail_FromSheetExcel = path['Sheet_MailList']# Sheet for reading data list of mailling
        DataBase_FromSheetExcel = path['Sheet_ReportData']# Sheet for reading data for sending
        SourceReport_FromSheetExcel = path['Sheet_SourceReport']# sheet source report
        BodyMail_FromSheetExcel = path['Sheet_Content']# Sheet for reading content of body mail of mailling
        Cons_VendorCode_FromExcelCol = path['Condition_Tracking']
        Cons_ToMail_FromExcelCol = path['Condition_To']
        Cons_CCMail_FromExcelCol = path['Condition_Mail']
        Cons_Subject_FromExcelCol = path['Condition_Subject']
        Cons_Dear_FromExcelCol = path['Condition_Dear']
        Cons_VendorCode_FilterCheck = path['Filter_Contraints']
        ColFormat_Nums1 = path['Format_Cons_Nums_1']
        ColFormat_date1 = path['Format_Cons_Date_ETD1']
        ColFormat_date2 = path['Format_Cons_Date_ETA1']    
        ColFormat_date3 = path['Format_Cons_Date_Shortage1']
        Others_Mail_Sender= path['Outlook_Sender']
        Path_For_ProcessingFile= path['Fixed_Processing_File_Path']
else: raise FileNotFoundError("Can not file the YAML file.")

if path.get('file_tracking') is None or not path['file_tracking'].strip(): # check condition1 if file tracking is none import from dynamic link else get address from yaml file
    def get_excel_files(): 
        root = tk.Tk()
        root.withdraw()
        file_paths = filedialog.askopenfilenames(title="Select files to import!", filetypes=[("Excel files", "*.xlsx;*.xls")])
        return file_paths
    selected_files = get_excel_files()
    if not selected_files: raise ValueError("No files selected. Please select a valid Excel file.")
    selected_file = selected_files[0]
else: selected_file = path['file_tracking']# -> Import file to extract and tracking

if path.get('Recording_Folder_Add') is None or not path['Recording_Folder_Add'].strip():
    Root_FolderAddress = os.path.join(os.path.join(os.path.expanduser("~"), "Desktop"), Root_FolderName)# link of http to Desktop and folder name
else: Root_FolderAddress = os.path.join(path['Recording_Folder_Add'],Root_FolderName) # folder address for accessing 
#--------------------------------Step 1 ís Done for import database and specification'
df = pd.read_excel(selected_file, sheet_name= SourceReport_FromSheetExcel, header=None, engine='openpyxl') # Load the source sheet into a DataFrame
expected_columns = ["Order No.", "Material", "Material Desc.", "Supplier", "Supplier Desc", "ETD", "ETA", "Backlog status 3", "Delay in ETD (today - ETD)", "Qty", "Running GIT Total", "Outstanding Qty", "Note"]# Define the expected columns (strip spaces for robustness)
header_row_idx = df[df.apply(lambda row: row.astype(str).str.contains(expected_columns[0], na=False).any(), axis=1)].index[0]# Find the header row index dynamically
df = pd.read_excel(selected_file, sheet_name= SourceReport_FromSheetExcel, header= header_row_idx)# Read the table starting from the identified header row
matched_columns = [col for col in expected_columns if col in df.columns]# Keep only the required columns (match dynamically)
if not matched_columns: raise KeyError("None of the expected columns were found in the detected table. Check column names.")# Validate if required columns exist
df_selected = df[matched_columns]# Select only matched columns

rename_map = {'Order No.': 'MMSA','Material': 'PN','Material Desc.': 'PN_Name','Supplier': 'Vendor','Supplier Desc': 'Vendor_Name','Backlog status 3': 'Backlog_STT',
    'Delay in ETD (today - ETD)': 'Delay_in_ETD','Qty': 'Order_Qty','Running GIT Total': 'GIT'}
df_selected.rename(columns=rename_map, inplace=True)

if 'ETD' in df_selected.columns:
    df_selected['ETD'] = pd.to_datetime(df_selected['ETD'], errors='coerce')
    df_selected['ETD'] = df_selected['ETD'].dt.date 

if 'ETA' in df_selected.columns:
    df_selected['ETA'] = pd.to_datetime(df_selected['ETA'], errors='coerce')
    df_selected['ETA'] = df_selected['ETA'].dt.date
        
df_selected.sort_values(by=['MMSA', 'PN', 'ETD'], ascending=[True, True, True], inplace=True)# Sort the DataFrame by PN, Supplier (MMSA assumed as Supplier), and ETD in ascending order (FIFO)
if 'ETD' in df_selected.columns: df_selected['ETD'] = pd.to_datetime(df_selected['ETD'], errors='coerce').dt.strftime('%m/%d/%Y')
if 'ETA' in df_selected.columns: df_selected['ETA'] = pd.to_datetime(df_selected['ETA'], errors='coerce').dt.strftime('%m/%d/%Y')

if 'Order_Qty' in df_selected.columns: # Assumption by FIFO [MMSA-PN] -ETD Follow to Qty- Good in Transit
    df_selected['Outstanding Qty'] = None
    grouped = df_selected.groupby(['MMSA', 'PN'])
    for (MMSA, PN), group in grouped:
        cumulative_qty = 0
        for index, row in group.iterrows():
            open_qty = row['Order_Qty'] if pd.notnull(row['Order_Qty']) else 0
            git = row['GIT'] if pd.notnull(row['GIT']) else 0
            if cumulative_qty == 0: cumulative_qty = open_qty - git
            else: cumulative_qty += open_qty
            df_selected.at[index, 'Outstanding Qty'] = cumulative_qty

df_selected = df_selected[df_selected['Outstanding Qty'] > 0]  # Remove rows where Outstanding Qty <= 0
df_selected = df_selected[df_selected['Note'] != 'No']# Delete rows where Note is 'No'
df_selected.drop(columns=['Note'], inplace=True)# Drop the 'Note' column

vendor_mail_df = pd.read_excel(selected_file, sheet_name= ListMail_FromSheetExcel, engine='openpyxl')# read sheet Vendor-mail in file excel from path selected file
with pd.ExcelWriter(Path_For_ProcessingFile, engine='openpyxl') as writer: # Create new excel file for processing the python using path fixed from yaml file as specificspecific
    df_selected.to_excel(writer, sheet_name= DataBase_FromSheetExcel, index=False)# push data from dataframe (df_selected) to new excel filefile in sheet Report
    vendor_mail_df.to_excel(writer, sheet_name= ListMail_FromSheetExcel, index=False)# push data from dataframe (df_selected) to new excel filefile in sheet Vendor-Mail

    workbook = writer.book
    for sheet_name in ['Report', 'Vendor-Mail']:
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:  # Necessary to avoid errors on header cells
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Adjust column width proportionally
            worksheet.column_dimensions[col_letter].width = adjusted_width
        worksheet.auto_filter.ref = worksheet.dimensions # Apply autofilter to the worksheet
# ----------------------------- Create the folder tree for recording history data
selected_file=Path_For_ProcessingFile
version = 0
year_folder_path = os.path.join(Root_FolderAddress, str(datetime.now().year)) # Create subdirectories for year and month
month_folder_path = os.path.join(year_folder_path, datetime.now().strftime("%b"))
date_folder_fullpath= os.path.join(month_folder_path, ver_date)
while os.path.exists(date_folder_fullpath):
    version += 1
    date_folder_fullpath = os.path.join(month_folder_path, f"{ver_date}_v{version}")
os.makedirs(date_folder_fullpath, exist_ok=True)

# -------------------------------------- Process specific sheets from imported file
try:
    DataBaseWH_Des = pd.read_excel(selected_file, sheet_name = DataBase_FromSheetExcel) # read sheet report-database
except ValueError as e: raise ValueError(f"Error reading sheets from the file: {e}")
Constraints_Array = DataBaseWH_Des[Cons_VendorCode_FilterCheck].dropna().unique() # Remove duplicate list vendor for using the loop in <=> unique list of vendor

for Result_Filter in Constraints_Array: # Export vendor-specific reports with table and filter
    Results_Output = DataBaseWH_Des[DataBaseWH_Des[Cons_VendorCode_FilterCheck] == Result_Filter].copy()
    safe_vendor_name = "".join(c if c.isalnum() else "_" for c in str(Result_Filter))  # Ensure valid filename
    output_file = os.path.join(date_folder_fullpath, f"{safe_vendor_name}.xlsx")
    wb = Workbook()
    ws = wb.active # Write to Excel with table and filter
    ws.title = "Data"
    for r_idx, row in enumerate(dataframe_to_rows(Results_Output, index=False, header=True), start=1): # Use dataframe_to_rows for efficient data transfer
        for c_idx, value in enumerate(row, start=1): ws.cell(row=r_idx, column=c_idx, value=value)
    
    for col_idx in range(1, Results_Output.shape[1] + 1): # Adjust column widths
        col_letter = get_column_letter(col_idx)
        max_length = max((len(str(cell.value)) for cell in ws[col_letter] if cell.value is not None), default=0)
        ws.column_dimensions[col_letter].width = max_length + 2  # Add padding for better readability

    if ColFormat_Nums1 not in Results_Output.columns: raise ValueError(f"Column '{ColFormat_Nums1}' not found in the dataset.") # formating cell as numbering with type "000,000.00"
    filter_column_idx = None
    for idx, col_name in enumerate(Results_Output.columns, start=1):
        if col_name == ColFormat_Nums1: filter_column_idx = idx
        break
    Results_Output[ColFormat_Nums1] = Results_Output[ColFormat_Nums1].astype(str).str.strip() # Trim whitespace for the specified column
    for row in ws.iter_rows(min_col=filter_column_idx, max_col=filter_column_idx, min_row=2, max_row=Results_Output.shape[0] + 1):# Apply formatting to the identified column
        for cell in row:
            if cell.value: 
                cell.value = str(cell.value).strip()  # Remove the blank before and after value
                try: cell.value = float(cell.value.replace(",", ""))  # convert to numberic if can
                except ValueError:
                    pass  # Ignore if wrong
            cell.number_format = "#,##0.00"  # Format as number with 2 decimal places
        
    date_columns_array = [ColFormat_date1, ColFormat_date2, ColFormat_date3]  # Replace with your actual column names
    for date_col in date_columns_array:
        if date_col in Results_Output.columns:
            print(f"Formatting column: {date_col} for vendor {Result_Filter}")
            date_column_idx = list(Results_Output.columns).index(date_col) + 1
            for row in ws.iter_rows(min_col=date_column_idx, max_col=date_column_idx, min_row=2, max_row=Results_Output.shape[0] + 1):
                for cell in row: cell.number_format = "dd/mm/yy"
        else: print(f"Warning: Column '{date_col}' not found for vendor {Result_Filter}")
    
    last_col = get_column_letter(Results_Output.shape[1])
    table_range = f"A1:{last_col}{Results_Output.shape[0] + 1}"
    table = Table(displayName=f"Table_{safe_vendor_name}", ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", # Apply a style to the table
        showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)
    wb.save(output_file)# Save file
#------------------------------------Open Mail Outlook    
def is_outlook_running():
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if proc.info['name'] and "OUTLOOK.EXE" in proc.info['name']:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return False

if is_outlook_running():
    print("Outlook is already running.")
else:
    print("Outlook is closed. Opening now...")
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    if os.path.exists(outlook_path):
        subprocess.Popen(outlook_path)
        time.sleep(5)  # Chờ Outlook mở hoàn toàn
    else:
        print("Check Outlook path! Outlook was not found.")

List_Mail = pd.read_excel(selected_file, sheet_name = ListMail_FromSheetExcel)# read sheet get the vendor code for searching
List_Mail.columns = List_Mail.columns.str.strip()
vendor_mail_dict = {} # Create dictionary from Vendor-Mail for mapping
def close_sharepoint_prompt(): #Function to close the window 'save to sharepoint
    try:
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "Save to SharePoint" in window_text: 
                    print(f"Closing window: {window_text}")
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)  # sent to close
        for _ in range(5):  # Try 5 times
            win32gui.EnumWindows(callback, None)
            time.sleep(1)  # delay 1 s
    except Exception as e: print(f"Error closing SharePoint prompt: {e}")
outlook = win32.Dispatch("Outlook.Application")  # Open Apps Outlook

for index, row in List_Mail.iterrows(): # loop in each rows from sheet mail
    try:
        vendor = row[Cons_VendorCode_FromExcelCol]
        subject_content= row[Cons_Subject_FromExcelCol] # Subject of the email
        receiver_mail = row[Cons_ToMail_FromExcelCol]
        cc_mail = row[Cons_CCMail_FromExcelCol]
        dear_mail=row[Cons_Dear_FromExcelCol]
        attachment_path = os.path.join(date_folder_fullpath, f"{vendor}.xlsx")
        alternate_email = Others_Mail_Sender
        
        if not os.path.exists(attachment_path):
            print(f"Attachment not found for {vendor}. Skipping email.")
            continue
        
        def dataframe_to_html_table(df):# Function to generate HTML table from DataFrame
            base_html = df.to_html(index=False, border=1, justify="left", classes="table table-striped")
            styled_html = base_html.replace(
            '<thead>',
            '<thead style="background-color: #007BFF; color: white; font-weight: bold; text-align: left;">')
            return styled_html
        
        if os.path.exists(attachment_path): html_table = dataframe_to_html_table(pd.read_excel(attachment_path))  # Generate HTML table
        else: html_table = "<p>No data available for this vendor.</p>"
        
        Outlook_mainBody= f"""
            <p style="font-weight: bold; font-family: 'Segoe UI', sans-serif; font-size: 13px;"> Dear {dear_mail},</p>
            <p>Please refer the table as below and more details in attachment!:</p>
            {html_table}
            <p style="color: #FF0000; font-weight: bold; font-family: 'Segoe UI', sans-serif; font-size: 13px;">Best regards,</p>
            <p> Automatically generated from system</p>
        """
        def send_email_with_outlook(): #function send out email
            mail = outlook.CreateItem(0)
            mail.Subject = subject_content
            mail.To = receiver_mail
            mail.CC = cc_mail
            mail.HTMLBody = Outlook_mainBody
            #mail.Body = f"{dear_mail},\n\nPlease find the attached report.\n\nBest regards,\nAutomatic System"
            
            if os.path.exists(attachment_path): mail.Attachments.Add(attachment_path)
            else: print(f"Attachment not found for {vendor}: {attachment_path}") 

            if isinstance(path, dict): # Check if path is a dictionary
                if path.get('Outlook_Sender') is None or not path['Outlook_Sender'].strip():
                    close_sharepoint_prompt()
                    time.sleep(2)  # Đợi 1 giây để xử lý
                    print(f"Email successfully sent to {vendor}: {receiver_mail}")   
                    mail.Send()             
                else: 
                    accounts = outlook.Session.Accounts # Setting other outlook email for alternative sending (must exist in Mailbox)
                    for account in accounts:
                        if account.SmtpAddress == alternate_email:
                            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # 64209: PR_ACCOUNT, change Account
                            break
                    else:
                        raise Exception(f"Account {alternate_email} not found in configured Outlook accounts.")
                    
                    mail.Send()
                    print(f"Email sent successfully from {alternate_email} to {receiver_mail}.")
            else: print("Error: 'path' is not a dictionary.")
            
        send_email_with_outlook()
    except Exception as e:
        print(f"Error sending email: {e}")

print(List_Mail)
print(vendor_mail_dict)      
