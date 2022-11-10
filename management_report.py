import pandas as pd, cx_Oracle, os, datetime, re, sys, subprocess as sp, smtplib, os.path, xlsxwriter
from os.path import isfile, join
from datetime import date, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from openpyxl import Workbook
from openpyxl import load_workbook as lw


#sql_files = ['gcc_managed.sql','gcc_info.sql','ddi_inc.sql','ddi_rfc.sql','ecs_inc.sql','ecs_rfc.sql','lan_inc.sql','wan_inc.sql','wan_rfc.sql','uc_inc.sql','uc_rfc.sql']
sql_files = ['gcc_managed.sql','gcc_info.sql','all_inc.sql','all_rfc.sql','wan_rfc.sql']
today_start = date.today() - datetime.timedelta(days=1)
today = datetime.datetime.now()
today_end = today.replace(hour=23, minute=59, second=59)
#today = datetime.date.today()
saturday = today + datetime.timedelta( (5-today.weekday()) % 7 )
sunday = today + datetime.timedelta( (6-today.weekday()) % 7 )
saturday_start = saturday.replace(hour=00, minute=00, second=00)
sunday_end = sunday.replace(hour=23, minute=59, second=59)



def send_mail(send_from, send_to, subject, text, files=None,server="127.0.0.1"):

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    #msg['Cc'] = cc_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    if files:
      for f in files:
        with open(f, "rb") as file:
            part = MIMEApplication(
                file.read(),
                Name=os.path.basename(f)
            )
        #after the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(f)
        msg.attach(part)

    smtp = smtplib.SMTP(server)
    #smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.send_message(msg)
    smtp.close()


def run_sql(path,sheet,df_ddi_inc,df_ddi_rfc,df_ecs_inc,df_ecs_rfc,df_lan_inc,df_uc_inc,df_uc_rfc,df_wan_inc):
  con = cx_Oracle.connect('BAPITSM/PrD#biss#0710@itsmprod.pfizer.com')
  file_name = "Management_Daily_Report_"+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")+".xlsx"
  try:

    if sheet == "GCC":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dddd, mmmm d, yyyy at hh:mm AM/PM'})
      worksheet.set_column('C:C', 35, fmt)
      worksheet.set_column(1, 1, 20)
      worksheet.set_column(3, 3, 50)
      worksheet.set_column(4, 4, 20)
      worksheet.set_column(5, 5, 40)
      worksheet.set_column(6, 6, 20)


      with open(os.path.join(sql_files_path,"gcc_managed.sql"), 'r') as f:
        sql1 = f.read()
      f.close()
      with open(os.path.join(sql_files_path,"gcc_info.sql"), 'r') as f:
        sql2 = f.read()
      f.close()

      data1 = pd.read_sql(sql1, con)
      data2 = pd.read_sql(sql2, con)
      con.close()


      if not data1.empty:

        #print(data1)
        #data1['SUMMARY'] = data1['SUMMARY'].str.split('<li>').str[1].str.split('</li>').str[0]
        ##data1['SUMMARY'] = data1['SUMMARY'].apply(lambda x: x.partition("<li>")[2].partition("</li>")[0])
        list_of_chars = ['<li>', '</li>', '<ul>', '</ul>']
        #Remove multiple characters from the string
        for character in list_of_chars:
          data1['SUMMARY'] = data1['SUMMARY'].str.replace(character, '')
        #start row, start col, end row, end col
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "GCC Managed Incidents for SNS from "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #first row, first col, last row, last column, data, cell format
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

      if not data2.empty:

        start_row = data1.shape[0]+3
        worksheet.add_table(start_row+5, 1, data2.shape[0]+start_row+5, data2.shape[1],
        {'data': data2.values.tolist(),
        'columns': [{'header': c} for c in data2.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title2 = "InfoCenter Communications"
        #Merge 3 cells.
        worksheet.merge_range(start_row+4, 1, start_row+4, data2.shape[1], title2, merge_format)

    if sheet == "DDI":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dd/mm/yy hh:mm'})
      worksheet.set_column('C:C', 20, fmt)
      worksheet.set_column('E:E', 20, fmt)
      worksheet.set_column(1, 1, 20)
      worksheet.set_column(3, 3, 40)
      worksheet.set_column(5, 5, 40)
      worksheet.set_column(6, 6, 20)


      #with open(os.path.join(sql_files_path,"ddi_inc.sql"), 'r') as f:
      #  sql1 = f.read()
      #f.close()
      #with open(os.path.join(sql_files_path,"ddi_rfc.sql"), 'r') as f:
      #  sql2 = f.read()
      #f.close()
      ##with open(os.path.join(sql_files_path,"ddi_rfc_mig.sql"), 'r') as f:
      ##  sql3 = f.read()
      ##f.close()

      data1 = df_ddi_inc
      data2 = df_ddi_rfc
      #data1 = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      #data2 = pd.read_sql(sql2.format(datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      ##data3 = pd.read_sql(sql3.format(datetime.date.strftime(saturday_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(sunday_end, "%m-%d-%Y %H:%M:%S")), con)
      #con.close()


      if not data1.empty:

        data1['CURRENT_STATUS'] = data1['CURRENT_STATUS'].str.split(':',3).str[3]
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "Network Operations Daily Report on HIGH/CRITICAL Tickets "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #Merge 3 cells.
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

      if not data2.empty:

        start_row = data1.shape[0]+3
        worksheet.add_table(start_row+5, 1, data2.shape[0]+start_row+5, data2.shape[1],
        {'data': data2.values.tolist(),
        'columns': [{'header': c} for c in data2.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title2 = "Upcoming Major Events"
        #Merge 3 cells.
        worksheet.merge_range(start_row+4, 1, start_row+4, data2.shape[1], title2, merge_format)

      #if not data3.empty:

        #start_row = data2.shape[0]+start_row+5
        ##start row, start col, end row, end col
        #worksheet.add_table(start_row+5, 1, data3.shape[0]+start_row+5, data3.shape[1],
        #{'data': data3.values.tolist(),
        #'columns': [{'header': c} for c in data3.columns.tolist()],
        #'style': 'Table Style Medium 9'})

        ##Create a format to use in the merged range.
        #merge_format = workbook.add_format({
        #'align': 'center',
        #'font_color': 'white',
        #'fg_color': 'blue'})

        #title3 = "Upcoming Weekend Migration"
        ##Merge 3 cells.
        #worksheet.merge_range(start_row+4, 1, start_row+4, data3.shape[1], title3, merge_format)

    if sheet == "ECS":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dd/mm/yy hh:mm'})
      worksheet.set_column('C:C', 20, fmt)
      worksheet.set_column('E:E', 20, fmt)
      worksheet.set_column(1, 1, 20)
      worksheet.set_column(3, 3, 40)
      worksheet.set_column(5, 5, 40)
      worksheet.set_column(6, 6, 20)
      worksheet.set_column(7, 7, 20)
      worksheet.set_column(8, 8, 20)


      #with open(os.path.join(sql_files_path,"ecs_inc.sql"), 'r') as f:
      #  sql1 = f.read()
      #f.close()
      #with open(os.path.join(sql_files_path,"ecs_rfc.sql"), 'r') as f:
      #  sql2 = f.read()
      #f.close()
      ##with open(os.path.join(sql_files_path,"ecs_rfc_mig.sql"), 'r') as f:
      ##  sql3 = f.read()
      ##f.close()

      data1 = df_ecs_inc
      data2 = df_ecs_rfc
      #data1 = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      #data2 = pd.read_sql(sql2.format(datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      ##data3 = pd.read_sql(sql3.format(datetime.date.strftime(saturday_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(sunday_end, "%m-%d-%Y %H:%M:%S")), con)
      #con.close()

      if not data1.empty:

        data1['CURRENT_STATUS'] = data1['CURRENT_STATUS'].str.split(':',3).str[3]
        #start row, start col, end row, end col
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "Network Operations Daily Report on HIGH/CRITICAL Tickets "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #Merge 3 cells.
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

      if not data2.empty:

        data2["DATE_OPENED"] = data2["DATE_OPENED"].fillna(' ')
        #print(data2)
        start_row = data1.shape[0]+3
        worksheet.add_table(start_row+5, 1, data2.shape[0]+start_row+5, data2.shape[1],
        {'data': data2.values.tolist(),
        'columns': [{'header': c} for c in data2.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title2 = "Upcoming Major Events"
        #Merge 3 cells.
        worksheet.merge_range(start_row+4, 1, start_row+4, data2.shape[1], title2, merge_format)

      #if not data3.empty:

        #start_row = data2.shape[0]+start_row+5
        ##start row, start col, end row, end col
        #worksheet.add_table(start_row+5, 1, data3.shape[0]+start_row+5, data3.shape[1],
        #{'data': data3.values.tolist(),
        #'columns': [{'header': c} for c in data3.columns.tolist()],
        #'style': 'Table Style Medium 9'})

        ##Create a format to use in the merged range.
        #merge_format = workbook.add_format({
        #'align': 'center',
        #'font_color': 'white',
        #'fg_color': 'blue'})

        #title3 = "Upcoming Weekend Migration"
        ##Merge 3 cells.
        #worksheet.merge_range(start_row+4, 1, start_row+4, data3.shape[1], title3, merge_format)

    if sheet == "LAN":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dd/mm/yy hh:mm'})
      worksheet.set_column('C:C', 20, fmt)
      worksheet.set_column(1, 1, 20)
      worksheet.set_column(3, 3, 20)
      worksheet.set_column(4, 4, 40)
      worksheet.set_column(5, 5, 20)
      worksheet.set_column(6, 6, 40)
      worksheet.set_column(7, 7, 30)


      #with open(os.path.join(sql_files_path,"lan_inc.sql"), 'r') as f:
      #  sql1 = f.read()
      #f.close()

      data1 = df_lan_inc
      #data1 = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      #con.close()

      if not data1.empty:

        data1['CURRENT_STATUS'] = data1['CURRENT_STATUS'].str.split(':',3).str[3]
        #start row, start col, end row, end col
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "Network Operations Daily Report on HIGH/CRITICAL Tickets "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #Merge 3 cells.
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

    if sheet == "WAN":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dd/mm/yy hh:mm'})
      worksheet.set_column('C:C', 20, fmt)
      worksheet.set_column('D:D', 20, fmt)
      worksheet.set_column('F:F', 20, fmt)

      worksheet.set_column(1, 1, 20)
      #worksheet.set_column(3, 3, 20)
      worksheet.set_column(4, 4, 40)
      worksheet.set_column(5, 5, 20)
      worksheet.set_column(6, 6, 40)
      worksheet.set_column(7, 7, 25)


      #with open(os.path.join(sql_files_path,"wan_inc.sql"), 'r') as f:
      #  sql1 = f.read()
      #f.close()
      with open(os.path.join(sql_files_path,"wan_rfc.sql"), 'r') as f:
        sql2 = f.read()
      f.close()

      data1 = df_wan_inc
      #data1 = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      data2 = pd.read_sql(sql2.format(datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      con.close()


      if not data1.empty:

        data1['CURRENT_STATUS'] = data1['CURRENT_STATUS'].str.split(':',3).str[3]
        #start row, start col, end row, end col
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "Network Operations Daily Report on HIGH/CRITICAL Tickets "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #Merge 3 cells.
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

      if not data2.empty:

        data2["DATE_OPENED"] = data2["DATE_OPENED"].fillna(' ')
        start_row = data1.shape[0]+3
        worksheet.add_table(start_row+5, 1, data2.shape[0]+start_row+5, data2.shape[1],
        {'data': data2.values.tolist(),
        'columns': [{'header': c} for c in data2.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title2 = "WAN Maintenance at Data Center"
        #Merge 3 cells.
        worksheet.merge_range(start_row+4, 1, start_row+4, data2.shape[1], title2, merge_format)

    if sheet == "UC":
      worksheet = workbook.add_worksheet(sheet)
      worksheet.hide_gridlines(2)
      fmt = workbook.add_format({'num_format':'dd/mm/yy hh:mm'})
      worksheet.set_column('C:C', 20, fmt)
      worksheet.set_column('E:E', 20, fmt)
      worksheet.set_column(1, 1, 20)
      worksheet.set_column(3, 3, 40)
      #worksheet.set_column(4, 4, 40)
      worksheet.set_column(5, 5, 20)
      worksheet.set_column(6, 6, 40)
      worksheet.set_column(7, 7, 25)


      #with open(os.path.join(sql_files_path,"uc_inc.sql"), 'r') as f:
      #  sql1 = f.read()
      #f.close()
      #with open(os.path.join(sql_files_path,"uc_rfc.sql"), 'r') as f:
      #  sql2 = f.read()
      #f.close()

      data1 = df_uc_inc
      data2 = df_uc_rfc
      #data1 = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      #data2 = pd.read_sql(sql2.format(datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
      #con.close()

      if not data1.empty:

        data1['CURRENT_STATUS'] = data1['CURRENT_STATUS'].str.split(':',3).str[3]
        #start row, start col, end row, end col
        worksheet.add_table(2, 1, data1.shape[0]+2, data1.shape[1],
        {'data': data1.values.tolist(),
        'columns': [{'header': c} for c in data1.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title1 = "Network Operations Daily Report on HIGH/CRITICAL Tickets "+datetime.date.strftime(today_start, "%m-%d-%Y")+" "+datetime.date.strftime(today_end, "%m-%d-%Y")
        #Merge 3 cells.
        worksheet.merge_range(1, 1, 1, data1.shape[1], title1, merge_format)

      if not data2.empty:

        start_row = data1.shape[0]+3
        worksheet.add_table(start_row+5, 1, data2.shape[0]+start_row+5, data2.shape[1],
        {'data': data2.values.tolist(),
        'columns': [{'header': c} for c in data2.columns.tolist()],
        'style': 'Table Style Medium 9'})

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
        'align': 'center',
        'font_color': 'white',
        'fg_color': 'blue'})

        title2 = "Upcoming Major Events"
        #Merge 3 cells.
        worksheet.merge_range(start_row+4, 1, start_row+4, data2.shape[1], title2, merge_format)


  except cx_Oracle.DatabaseError as e:
    #print("Cannot connect to DB. Please check if DB is available.")
    print(e)



#Main
if __name__ == "__main__":

  emaillst = "guillermoantonio.cortes@pfizer.com"
  sql_files_path = "sql/"
  outputdir = "output/"
  #print(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"))

  sheets = ['GCC','DDI','ECS','LAN','WAN','UC']
  #create empty .xlsx file
  file = "Management_Daily_Report_"+datetime.date.strftime(today_start, "%m-%d-%Y")+"_"+datetime.date.strftime(today_end, "%m-%d-%Y")+".xlsx"

  #sql files validation
  validation = 0
  list_sql_files = os.listdir(sql_files_path)
  for sql_file in sql_files:
    if sql_file not in list_sql_files:
      print("File: "+sql_file+" missing. Please check SQL files directory.")
      validation = validation + 1
  if validation == 0:
    print("SQL files validated succesfully !!!")
  else:
    sys.exit(1)

  con = cx_Oracle.connect('BAPITSM/PrD#biss#0710@itsmprod.pfizer.com')
  with open(os.path.join(sql_files_path,"all_inc.sql"), 'r') as f:
    sql1 = f.read()
  f.close()
  with open(os.path.join(sql_files_path,"all_rfc.sql"), 'r') as f:
    sql2 = f.read()
  f.close()
  df_all_inc = pd.read_sql(sql1.format(datetime.date.strftime(today_start, "%m-%d-%Y %H:%M:%S"),datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
  df_all_rfc = pd.read_sql(sql2.format(datetime.date.strftime(today_end, "%m-%d-%Y %H:%M:%S")), con)
  con.close()
  ###DDI
  df_ddi_inc = df_all_inc[df_all_inc['ASSIGNMENT'].str.contains('GBL-NETWORK DDI')][["INCIDENT#","DATE_OPENED","DESCRIPTION","PRIORITY","CURRENT_STATUS","STATUS"]]
  df_ddi_rfc = df_all_rfc[(df_all_rfc['ASSIGN_DEPT'].str.contains('GBL-NETWORK DDI')) & (df_all_rfc['DESCRIPTION'].str.lower().str.contains('migration|upgrade|replacement')==False)][["RFC_REFERENCE","DATE_OPENED","DESCRIPTION","DATE_OF_EVENT"]]
  ###ECS
  df_ecs_inc = df_all_inc[df_all_inc['ASSIGNMENT'].str.contains('GBL-NETWORK ECS')][["INCIDENT#","DATE_OPENED","DESCRIPTION","PRIORITY","CURRENT_STATUS","STATUS","LOCATION"]]
  df_ecs_rfc = df_all_rfc[(df_all_rfc['ASSIGN_DEPT'].str.contains('GBL-NETWORK ECS')) & (df_all_rfc['DESCRIPTION'].str.lower().str.contains('migration|upgrade|replacement')==False)][["RFC_REFERENCE","DATE_OPENED","DESCRIPTION","DATE_OF_EVENT"]]
  ###LAN
  df_lan_inc = df_all_inc[(df_all_inc['ASSIGNMENT'].str.contains('GBL-NETWORK LAN')) & (df_all_inc['DESCRIPTION'].str.lower().str.contains('related')==False)][["INCIDENT#","DATE_OPENED","PRIORITY","DESCRIPTION","STATUS","CURRENT_STATUS","LOCATION"]]
  ###UC
  df_uc_inc = df_all_inc[(df_all_inc['ASSIGNMENT'].str.contains('GBL-NETWORK UC')) & (df_all_inc['DESCRIPTION'].str.lower().str.contains('dial-peer')==False)][["INCIDENT#","DATE_OPENED","DESCRIPTION","PRIORITY","STATUS","CURRENT_STATUS","LOCATION"]]
  df_uc_rfc = df_all_rfc[(df_all_rfc['ASSIGN_DEPT'].str.contains('GBL-NETWORK UC')) & (df_all_rfc['DESCRIPTION'].str.lower().str.contains('migration|upgrade|replacement')==False)][["RFC_REFERENCE","DATE_OPENED","DESCRIPTION","DATE_OF_EVENT"]]
  ###WAN
  df_wan_inc = df_all_inc[(df_all_inc['ASSIGNMENT'].str.contains('GBL-NETWORK WAN')) & (df_all_inc['DESCRIPTION'].str.lower().str.contains('related')==False)][["INCIDENT#","DATE_OPENED","PRIORITY","DESCRIPTION","STATUS","CURRENT_STATUS","LOCATION"]]

  workbook = xlsxwriter.Workbook(os.path.join(outputdir,file))
  for sheet in sheets:
    run_sql(outputdir,sheet,df_ddi_inc,df_ddi_rfc,df_ecs_inc,df_ecs_rfc,df_lan_inc,df_uc_inc,df_uc_rfc,df_wan_inc)
  workbook.close()
