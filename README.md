# Automatic Importing Daily Reports into your Server

## Introduction

Our client receives daily email reports and desires a comprehensive data analysis. To achieve this, I propose implementing a historical data view within a dedicated server. By using a server, data storage becomes seamless, and accessing and viewing the information becomes effortless.

To address concerns about management, I assure the client that we can streamline the process with a Python script, automating the data transfer and sending confirmation messages once the task is completed. This ensures a smooth and reliable workflow, eliminating the need for manual intervention.

The client enthusiastically embraces this solution and expresses interest in its continuation.

## Python Code - Report Import into Server

Below is the scipt that will transform, and import the data into the server.  

#### Imports Being Used

```
import win32com.client
import datetime
import pandas as pd
import io
import os
import pyodbc
```

### Setting Up the application, accessing the email, and searching for the report
```
# *Setting up application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# *Access the email
inbox = outlook.Folders("user@email.com").Folders("Inbox").Folders("SubFolder01").Folders("SubFolder02").Folders("SubFolder03")

# *Search Crieria
today = datetime.date.today() - datetime.timedelta(days=1)
target_subject = "Email Subject"
messages = inbox.items.Restrict("[ReceivedTime] >='" + today.strftime('%m/%d/%Y') + "'")
```

### Automating Process will begin
```
# *Automation will Begin
if len(messages) == 0:
    print("Nothing was found")
else:
    for message in messages:
        if target_subject in message.Subject:
            print("Found email with Subject:", message.Subject)
            attachments = message.Attachments
            report_a_df = None
            for i in range(attachments.Count):
                attachment = attachments.Item(i+1)
                if attachment.filename == "Daily_Report.xlsx":
                    print("Process 'Daily_Report.xlsx'..........")
                    temp_file_path = os.path.join(os.getcwd(), attachment.FileName)
                    attachment.SaveAsFile(temp_file_path)
                    report_a_df = pd.read_excel(temp_file_path, header=2)
                    report_a_df.insert(0,"COLUMN_HEADER01", pd.Timestamp.now().strftime("%Y-%m-%d"))
                    report_a_df['CUSTOM_HEADER01'] = report_a_df['COLUMN_HEADER02'].apply(get_CUSTOM_HEADER01)
                    report_a_df['CUSTOM_HEADER02'] = report_a_df['COLUMN_HEADER03'].apply(get_inv_type)
                    # *Converting Columns
                    report_a_df['COLUMN_HEADER01'] = pd.to_datetime(report_a_df['COLUMN_HEADER01'])
                    report_a_df['COLUMN_HEADER04'] = report_a_df['COLUMN_HEADER04'].astype('str')
                    report_a_df['COLUMN_HEADER03'] = report_a_df['COLUMN_HEADER03'].astype('str')
                    report_a_df['CUSTOM_HEADER02'] = report_a_df['CUSTOM_HEADER02'].astype('str')
                    report_a_df['CUSTOM_HEADER01'] = report_a_df['CUSTOM_HEADER01'].astype('str')
                    report_a_df['COLUMN_HEADER02'] = report_a_df['COLUMN_HEADER02'].astype('str')
                    report_a_df['COLUMN_HEADER05'] = report_a_df['COLUMN_HEADER05'].astype('str')
                    report_a_df['COLUMN_HEADER06'] = report_a_df['COLUMN_HEADER06'].astype('str')
                    report_a_df['COLUMN_HEADER07'] = pd.to_numeric(report_a_df['COLUMN_HEADER07'], errors='coerce').fillna(0).astype('int64') 
                        
                    # *This Data-Frame is met to record the current Inventory Level 
                    current_report = report_a_df[['COLUMN_HEADER01','COLUMN_HEADER04','COLUMN_HEADER03','CUSTOM_HEADER01','COLUMN_HEADER02', 'COLUMN_HEADER05', 'COLUMN_HEADER06', 'COLUMN_HEADER07']]
                        
                    # *This Data-Frame is met to record historical data
                    historical_report = report_a_df[['COLUMN_HEADER01','COLUMN_HEADER04','CUSTOM_HEADER02','CUSTOM_HEADER01', 'COLUMN_HEADER05', 'COLUMN_HEADER06', 'COLUMN_HEADER07']]
                    historical_report = historical_report.groupby(['COLUMN_HEADER01','COLUMN_HEADER04','CUSTOM_HEADER02','CUSTOM_HEADER01', 'COLUMN_HEADER05', 'COLUMN_HEADER06']).sum().reset_index()
                        
                    # *Establishing connection to SQL Sever
                    server = '<insert_server>'
                    database = '<insert_database>'
                    username = '<insert_username>'
                    password = '<insert_password>'
                    cnxn = pyodbc.connect(f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}')
                        
                    # *Creating a cursor
                    cursor = cnxn.cursor()
                        
                    # *Execute the SQL statment to delete current Data
                    cursor.execute('DELETE FROM dbo.[Current_Inventory]')
                    cursor.commit()
                        
                    # *Adding New Data
                    table1 = '[Current_Inventory]'
                    table2 = '[Historical_Inventory]'
                        
                    for index, row in current_report.iterrows():
                        cursor.execute(f"""
                                           INSERT INTO {table1} (
                                               [COLUMN_HEADER01], [COLUMN_HEADER04], [COLUMN_HEADER03], [CUSTOM_HEADER01], [COLUMN_HEADER02], [COLUMN_HEADER05], [COLUMN_HEADER06], [COLUMN_HEADER07])
                                               values (?,?,?,?,?,?,?,?)""",
                                               (row['COLUMN_HEADER01'], row['COLUMN_HEADER04'], row['COLUMN_HEADER03'], row['CUSTOM_HEADER01'], row['COLUMN_HEADER02'], row['COLUMN_HEADER05'], row['COLUMN_HEADER06'], row['COLUMN_HEADER07']))
                    cursor.commit()
                        
                    for index, row in historical_report.iterrows():
                        cursor.execute(f"""
                                           INSERT INTO {table2} (
                                               [COLUMN_HEADER01], [COLUMN_HEADER04], [CUSTOM_HEADER02], [CUSTOM_HEADER01],  [COLUMN_HEADER05], [COLUMN_HEADER06], [COLUMN_HEADER07])
                                               values (?,?,?,?,?,?,?)""",
                                               (row['COLUMN_HEADER01'], row['COLUMN_HEADER04'], row['CUSTOM_HEADER02'], row['CUSTOM_HEADER01'], row['COLUMN_HEADER05'], row['COLUMN_HEADER06'], row['COLUMN_HEADER07']))
                        
                    cursor.commit()
                        
                    # *Closing Connection
                    cnxn.close()
                    break
            if report_a_df is not None:
                print("Data successfully pulled as 'report_a_df'.")
                print("Data successfully imported into server")
            else:
                print("No report was found in this email")
```
