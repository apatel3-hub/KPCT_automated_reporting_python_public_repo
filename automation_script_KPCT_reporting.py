# =============================================================================
# Creating automated KPCT Excel reports 
# =============================================================================

import os
import pyodbc
import win32com.client
import shutil
import datetime

# Creating connection to Oracle Database (EDW)
conn_str = 'DSN=servername;UID=username;PWD={}'.format(os.environ.get('password'))
conn = pyodbc.connect(conn_str, autocommit=True)
cursor = conn.cursor()
#=========================================================
#Put the location of the script in the link variable below
#=========================================================
inputdir = r'SQL file path.sql'

f = open(inputdir)
full_sql = f.read()
sql_commands = full_sql.split(';')
for sql_command in sql_commands:
    cursor.execute(sql_command)

conn.close()


#=========================================================
#Put the location of the excel files in the path variable below
#=========================================================

# Start an instance of Excel
xlapp = win32com.client.DispatchEx("Excel.Application")

#create a loop
excel_folder_path = r'folder path f excel files'

for filename in os.listdir(excel_folder_path):
    file = (os.path.join(excel_folder_path, filename))
    wb = xlapp.Workbooks.Open(file)
    #xlapp.Visible = True   #this is just for testing
    # Refresh all data connections.
    wb.RefreshAll()
    #This will help wait till the queries run
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    wb.Close(True)
# Quit excel app
xlapp.Quit()


#================================================================
#Copy and Rename the files
#================================================================

#changes to current directory (where the new folder needs to be created )
directory = os.chdir(r'Paste the path of new directory')
#os.listdir()
#os.remove('PCP reporting 20191214')

#Adds a folder with current date to the directory defined above
os.makedirs(datetime.date.today().strftime("PCP reporting %Y%m%d")) 
folder_name =  datetime.date.today().strftime("PCP reporting %Y%m%d")
suffix_file_name = (datetime.date.today().replace(day=1) - datetime.timedelta(days=1) - datetime.timedelta(3*365/12)).strftime("%Y%m%d")

#Path to paste the refreshed files in the folder 
destinatin_path = rf'folderpath\{folder_name}'

#use shutil for copy the files
for filename in os.listdir(excel_folder_path):
    files = (os.path.join(excel_folder_path, filename))
    shutil.copy(files, destinatin_path)

#renames all files by adding date in the end of file name
for filename in os.listdir(destinatin_path):
    file_name_wo_ext  = os.path.splitext(filename)[0]
    os.rename(os.path.join(destinatin_path,filename), os.path.join(destinatin_path,f'{file_name_wo_ext}'+f'_{suffix_file_name}'+'.xlsx'))


#================================================================
#Send emails
#================================================================

current_year_month = datetime.today().strftime("%b %Y")

outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email_id'
mail.Subject = 'KPCT Process'
#mail.Body = 'Hello Automation'
mail.HTMLBody = f'''Congratulations.!! The test is successful for {current_year_month} <br>
<br>
<br>
''' #this field is optional

# To attach a file to the email (optional):
#attachment  = r'Path to the attachment'
#mail.Attachments.Add(attachment)

#sends the email 
mail.Send()

