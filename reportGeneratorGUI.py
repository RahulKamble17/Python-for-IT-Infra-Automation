import customtkinter
import CTkMessagebox
import tkinter as tk
from tkinter import filedialog
from PIL import Image
import win32com.client as win32
import csv
import pandas as pd
import datetime
import pytz
import xlwings as xw
import pyautogui
import imageExcel
import win32clipboard

import excelEdit

#GUI Window
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root=customtkinter.CTk()
root.geometry("930x550")
root.title("Digital Command Center")

def clearClipBoard():
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()

#Dialog Box for selecting input file
def openFile(b_name):
    smo_buttons=['aging','change','updated']
    file_path=filedialog.askopenfilename()
    if file_path.endswith(".csv"):
        if b_name not in smo_buttons:
            mail(file_path,b_name)
        else:
            CTkMessagebox.CTkMessagebox(title="Error", message="Please select an Excel file") 
    elif file_path.endswith(".xlsx"):
        if b_name in smo_buttons:
            smo_Reports(file_path,b_name)
        else:
            CTkMessagebox.CTkMessagebox(title="Error", message="Please select a CSV file")

#Function for SMO Reports
def smo_Reports(file_path,b_name):
    clearClipBoard()
    # *** Work In Progress ***  Updated date report: P1/P2 Status
    #if b_name=='updated':
    #    inc_file_path=filedialog.askopenfilename()

    
    # Opening excel file 
    workbook=xw.Book(file_path, read_only=True)
    worksheet=workbook.sheets.active
    app= xw.apps.active
    app.display_alerts=False

    #Reading data from old excel file to new
    data= worksheet.used_range.value
    new_wb=xw.Book()
    new_sheet=new_wb.sheets.active
    new_sheet.range('A1').value = data
    new_wb.save("C:\\Users\\2106624\\Downloads\\modified.xlsx")

    # Getting row count
    last_row=new_sheet.range('A'+ str(new_sheet.cells.last_cell.row)).end('up').row
    print(last_row)
    last_row=last_row-1 # Subtracting 1st header row
    if last_row==0:
        imageExcel.grabImage(b_name,"empty")
        workbook.save()
        workbook.close()
        return

    #Adding 'Aging/Updated in days' column
    if b_name in ['change','aging','updated']:
        last_col=new_sheet.api.UsedRange.Columns.Count
        new_sheet.api.Columns(last_col+1).Insert()
        if b_name=='updated':
            col_title='Updated in days'
            color_thres=2
        elif b_name in ['change','aging']:
            col_title='Aging in days'
            color_thres=5

        #Using Excel formula to calculate value
        new_sheet.range((1,last_col+1)).value=col_title
        if b_name=='change':
            col_letter='M'
            new_sheet.range('M2').formula= '=NOW()-J2'
            cell= new_sheet.range('M2')
        elif b_name=='aging':
            col_letter='L'
            new_sheet.range('L2').formula= '=NOW()-G2'
            cell= new_sheet.range('L2')
        else:
            col_letter='L'
            new_sheet.range('L2').formula= '=NOW()-H2'
            cell= new_sheet.range('L2')
        #Converting Aging/Updated date to days with 2 decimal points
        cell.number_format = '#,##0.00;[Red](#,##0.00);0.00;@'

        #Applying same formula to all 'Aging/Updated in days' rows
        last_row=new_sheet.range('A'+ str(new_sheet.cells.last_cell.row)).end('up').row
        new_sheet.range(col_letter+'2:'+col_letter+str(last_row)).api.FillDown()

        #Sorting 'Aging/Updated in days' in descending order
        range_to_sort= new_sheet.used_range
        data= range_to_sort.options(pd.DataFrame, header=1, index=False).value

        sorted_data= data.sort_values(by=col_title, ascending=False)

        new_sheet.range('A2').value=sorted_data.values.tolist()

        #For coloring Aging days greater than 5 & Updated greater than 2 with red
        range_to_sort = new_sheet.range(col_letter+'2:'+col_letter+ str(last_row)) 
        values=[cell.value for cell in range_to_sort]

        for i,value in enumerate(values):
            if new_sheet[col_letter+str(i+2)].value >=color_thres:
                new_sheet[col_letter+str(i+2)].api.Font.ColorIndex= 3

    #Adding bold font & blue fill to first row
    new_sheet.range('1:1').api.Font.Bold=True
    fill_color=(176,196,222)
    new_sheet.range('1:1').color=fill_color

    #Adding all borders
    used_range=new_sheet.used_range
    for cell in used_range:
        cell.api.Borders.Weight=2

    #Autofitting & resizing the data 
    if b_name in ['aging','updated']:
        new_sheet.range('B1').column_width =15
        new_sheet.range('C1').column_width =8
        new_sheet.range('D1').column_width =8
        range_to_sort = new_sheet.range('C2:C'+ str(last_row)) 
        values=[cell.value for cell in range_to_sort]

        #Checking if any 'State' is 'On-Hold' & setting width accordingly
        flag=0
        for i,value in enumerate(values):
            if new_sheet['C'+str(i+2)].value =='On Hold':
                flag+=1
        if flag==0:
            new_sheet.range('K1').column_width =0
        
        #Condition to highlight rows with missing data
        for i in range(int(last_row)-1):
            if (new_sheet['D'+str(i+2)].value is None or new_sheet['H'+str(i+2)].value is None or new_sheet['B'+str(i+2)].value is None or new_sheet['F'+str(i+2)].value is None):
                fill_color=(255,255,204)
                new_sheet.range(str(i+2)+':'+str(i+2)).color=fill_color
        
    elif b_name=='change':
        #Creating a change request table sorted according to 'Planned end date'
        
        old_sheet=new_wb.sheets['Sheet1']
        sheet2=old_sheet.copy(after=old_sheet)
        sheet2.name= 'Sheet2'

        range_to_sort= sheet2.used_range
        data= range_to_sort.options(pd.DataFrame, header=1, index=False).value

        sorted_data= data.sort_values(by='Planned end date')

        sheet2.range('A2').value=sorted_data.values.tolist()

        #Autofit and resize
        new_sheet.autofit()
        new_sheet.range('B1').column_width =15
        new_sheet.range('C1').column_width =0
        new_sheet.range('D1').column_width =0
        new_sheet.range('E1').column_width =0
        new_sheet.range('F1').column_width =0
        new_sheet.range('G1').column_width =0
        new_sheet.range('I1').column_width =0
        new_sheet.range('J1').column_width =0
        new_sheet.range('K1').column_width =0
        new_sheet.range('L1').column_width =0

        sheet2.autofit()
        sheet2.range('B1').column_width =15
        sheet2.range('C1').column_width =6
        sheet2.range('D1').column_width =7
        sheet2.range('E1').column_width =10
        sheet2.range('F1').column_width =6
        sheet2.range('G1').column_width =15
        
        for i in range(int(last_row)-1):
            if (sheet2['G'+str(i+2)].value is None or sheet2['K'+str(i+2)].value is None or sheet2['L'+str(i+2)].value is None or sheet2['B'+str(i+2)].value is None  or sheet2['I'+str(i+2)].value is None):
                fill_color=(255,255,204)
                sheet2.range(str(i+2)+':'+str(i+2)).color=fill_color

        range_to_sort = sheet2.range(col_letter+'2:'+col_letter+ str(last_row)) 
        values=[cell.value for cell in range_to_sort]

        for i,value in enumerate(values):
            if sheet2[col_letter+str(i+2)].value >=5:
                sheet2[col_letter+str(i+2)].api.Font.ColorIndex= 3
            else:
                sheet2[col_letter+str(i+2)].api.Font.ColorIndex= 1
        new_wb.save("C:\\Users\\2106624\\Downloads\\modified.xlsx")

        #Call to capture tables and create mail
        imageExcel.grabImage(b_name,None)
        clearClipBoard()
        print("clipB cleared")
        workbook.close()
    
    new_wb.save("C:\\Users\\2106624\\Downloads\\modified.xlsx")

    #Creating Pivot tables and modified sheets for Aging & Updated in date reports
    if b_name in ['aging','updated']:
        file_name="C:\\Users\\2106624\\Downloads\\modified.xlsx"
        mod_df= pd.read_excel(file_name, sheet_name='Sheet1')
        #Get Incident Pivot
        inc_return=excelEdit.createPivotTable('Incident',b_name)
        print(inc_return)
        #Get Task Pivot
        task_return=excelEdit.createPivotTable('Catalog Task',b_name)
        print(task_return)

        #Formatting Pivot and sheet
        def sheet_edit(sheet_val):

            new_wb=xw.Book(file_name)

            new_sheet=new_wb.sheets[sheet_val]
            new_sheet.autofit()
            new_sheet.range('B1').column_width =15
            new_sheet.range('C1').column_width =8
            new_sheet.range('D1').column_width =8
            last_row=new_sheet.range('A'+ str(new_sheet.cells.last_cell.row)).end('up').row
            range_to_sort = new_sheet.range('C2:C'+ str(last_row)) 
            number_format_range=new_sheet.range('L2:L'+ str(last_row)) 
            for cell in number_format_range:
                cell.number_format = '#,##0.00;[Red](#,##0.00);0.00;@'
            values=[cell.value for cell in range_to_sort]
            flag=0
            for i,value in enumerate(values):
                if new_sheet['C'+str(i+2)].value =='On Hold':
                    flag+=1
            if flag==0:
                new_sheet.range('K1').column_width =0
            #Condition to highlight rows with missing data
            for i in range(int(last_row)-1):
                if (new_sheet['D'+str(i+2)].value is None or new_sheet['H'+str(i+2)].value is None or new_sheet['B'+str(i+2)].value is None or new_sheet['F'+str(i+2)].value is None):
                    fill_color=(255,255,204)
                    new_sheet.range(str(i+2)+':'+str(i+2)).color=fill_color

            #Coloring days
            range_to_sort = new_sheet.range(col_letter+'2:'+col_letter+ str(last_row)) 
            values=[cell.value for cell in range_to_sort]
            for i,value in enumerate(values):
                if int(new_sheet[col_letter+str(i+2)].value) >=color_thres:
                    new_sheet[col_letter+str(i+2)].api.Font.ColorIndex= 3

            #Adding bold font & blue fill to first row
            new_sheet.range('1:1').api.Font.Bold=True
            fill_color=(176,196,222)
            new_sheet.range('1:1').color=fill_color

            #Adding all borders
            used_range=new_sheet.used_range
            for cell in used_range:
                cell.api.Borders.Weight=2

            new_wb.save(file_name)

        #Create Task and Incident Sheets
        if task_return == 'pass' or inc_return =='pass':           
            task_df=mod_df[(mod_df['State'] !='Resolved') & (mod_df['Task type']== 'Catalog Task')]
            inc_df=mod_df[(mod_df['State'] !='Resolved') & (mod_df['Task type']== 'Incident')]
            with pd.ExcelWriter(file_name, mode='a') as writer:
                if task_return=='pass':
                    task_df.to_excel(writer, sheet_name='TaskSheet',index=False)
                    print("task sheet written")

                if inc_return=='pass':
                    inc_df.to_excel(writer, sheet_name='INCSheet',index=False)
                    print("inc sheet written")
            #Call sheet_edit method
            if task_return=='pass':
                sheet_edit('TaskSheet')
            if inc_return=='pass':
                sheet_edit('INCSheet')
            #Call to capture data and generate mail
            if b_name in ['aging','updated']:
                imageExcel.grabImage(b_name,None)
                print("image grabbed")
    workbook.close()
    new_wb.close()
    app.display_alerts=False
    
    
    
#Single P1/P2 Alert function       *** Work In Progress ***
def singleAlert(b_name):
    f_path="single"
    mail(f_path,b_name)

#Mail generator for Daily reports and Multiple alerts
def mail(file_path,b_name):
    #Temporary passing codition for single alerts     *** Work In Progress ***
    path=file_path
    if path=="single":
        row_count=0
    #Diallog box for input Multiple alerts file
    else:
        file=open(file_path,'r')
        reader=csv.reader(file)
        row_count= sum(1 for row in reader)-1
        df = pd.read_csv(file_path)

    #Common Table Header for Infra & P1/P2 Reports
    temp_head="""<table>
        <tr bgcolor="#0563C1">
        <th width="100" height="25"><b>Ticket No.</th>
        <th width="100" height="25"><b>Priority</th>
        <th width="100" height="25"><b>Assignment Group</th>
        <th width="100" height="25"><b>Assigned</th>
        <th width="100" height="25"><b>State</th>
        <th width="100" height="25"><b>Resolution Code</th>
        <th width="150" height="25"><b>Short Description</th>
        <th width="100" height="25"><b>Opened</th>
        <th width="150" height="25"><b>Category</th>
        <th width="150" height="25"><b>Track</th>
        <th width="150" height="25"><b>Lead</th>
        <th width="150" height="25"><b>MIM Comments</th>
        </tr>"""

    #table header for Alerts
    table_head="""
        <br>
        <table>
        <tr bgcolor="#0563C1">
        <th width="100" height="25"><b>Ticket No.</th>
        <th width="100" height="25"><b>Priority</th>
        <th width="100" height="25"><b>State</th>
        <th width="100" height="25"><b>Assigned</th>
        <th width="100" height="25"><b>Assignment Group</th>
        <th width="100" height="25"><b>Opened</th>
        <th width="150" height="25"><b>Short Description</th>
        </tr>
        """

    #Time and Date for Infra and P1/P2 reports
    utc_now=datetime.datetime.utcnow()

    ist= pytz.timezone('Asia/Kolkata')

    ist_now=utc_now.replace(tzinfo=pytz.utc).astimezone(ist)
    est=pytz.timezone('US/Eastern')
    est_now=ist_now.astimezone(est)

    # Infra Report: Calcualting current Date
    infra_date=est_now.strftime('%d-%B-%Y')

    # P1/P2 Report: Calculating Shift
    if ist_now.hour > 8 and ist_now.hour <= 16:
        p1p2_shift='Morning'
    elif ist_now.hour > 16 and ist_now.hour <= 24:
        p1p2_shift='Afternoon'
    elif ist_now.hour > 0 and ist_now.hour <= 8:
        p1p2_shift='Night'
    else:
        p1p2_shift=''

    # Infra Report: Calculating Shift
    if est_now.hour < 12:
        day_time='Morning'
    elif est_now.hour < 17:
        day_time='Afternoon'
    else:
        day_time='Night'

    #Add data rows to all tables
    def calRows():
        # *** Work In Progress ***  Single Alert: Temporary empty alert template generator
        if(button_name=="single"):
            t_row=table_head
            t_row+="""<tr>"""
            k=0
            for k in range(7):
                t_row+=f"""<td width="100" height="25">Row 1 Column {k}</td>\n"""
            t_row+="""</tr>\n"""
            return t_row
        # Multiple Alerts: Adding rows using input from .csv file & dataframe
        elif(button_name=="multiple"):
            i=0
            temp=table_head
            if(row_count!=0):
                for i in range(row_count):
                    temp+="<tr>"
                    temp+="""<td width="100" height="25">"""+df['number'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['priority'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['assigned_to'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['state'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['assignment_group'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['opened_at'][i]+"""</td>"""
                    temp+="""<td width="150" height="25">"""+df['short_description'][i]+"""</td>"""
                    temp+="</tr>"
                return temp
            elif(row_count==0):
                temp+="""<th colspan="7">No tickets</th>"""
                return temp
        #Adding ticket data rows using input from .csv file & dataframe
        elif(button_name=="p1p2" or button_name=="infra"):
            i=0
            temp=temp_head
            if(row_count!=0):
                #Setting the fill color to Orange for 'In Progress' tickets
                state_color=[]
                state_color.clear()
                for j in range(row_count):
                    if (df['state'][j]=='In Progress' or df['state'][j]=='On Hold'):
                        state_color.append(""" bgcolor="orange"> """)
                    else:
                        state_color.append(">")
                #Adding rows to tables
                for i in range(row_count):
                    temp+="<tr>"
                    temp+="""<td width="100" height="25">"""+df['number'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['priority'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['assignment_group'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['assigned_to'][i]+"""</td>"""
                    temp+="""<td width="100" height="25" """+state_color[i]+df['state'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['close_code'][i]+"""</td>"""
                    temp+="""<td width="150" height="25">"""+df['short_description'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['opened_at'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""+df['category'][i]+"""</td>"""
                    temp+="""<td width="100" height="25">"""  """</td>"""
                    temp+="""<td width="100" height="25">"""  """</td>"""
                    temp+="""<td width="100" height="25">"""  """</td>"""
                    temp+="</tr>"
                return temp
            elif(row_count==0):
                if (button_name=='infra'):
                    temp="<b>No Tickets</b>"
                else:
                    temp+="""<th colspan="12">No tickets</th>"""
                return temp

    #Creating Outlook mail object
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    #check button
    button_name= b_name

    # Alerts: Person contacted 
    person=personContacted.get()
    if (len(person)!=0):
        greet_str=f"As discussed, {person} is checking the below mentioned P2 incident(s).<br>"
    else:
        greet_str="We received the following incident(s) please respond<br>"    

    # *** Work In Progress *** Single alerts: Mail To and Cc selection 
    if (button_name=="single"):
        mail.To= ' '
        mail.Subject = 'P2 Alert:'
        
    # Multiple alerts: Mail To and Cc Selection 
    elif (button_name=="multiple"):
        # replace with the email address you want to send the email to
        if (row_count!=0):
            if (df['assignment_group'][0]=="INFRA SUPPORT PLUTO WINDOWS"):
                mail.To = '' #Mail ID goes here..
                track='Windows'
            elif (df['assignment_group'][0]=="INFRA SUPPORT PLUTO LINUX"):
                mail.To = '' #Mail ID goes here..
                track='Linux'
            elif (df['assignment_group'][0]=="INFRA SUPPORT PLUTO DBA"):
                mail.To = '' #Mail ID goes here..
                track='DBA'
            else:
                mail.To= ' '
                track=''
        else:
            mail.To= ' '
        mail.Subject = 'Multiple P2 Alerts: '+track
        mail.CC='' #Mail ID goes here..
    
    # P1/P2 Report: Adding current shift to mail subject, To and Cc Selection  
    elif (button_name=="p1p2"):
        mail.To='' #Mail ID goes here..
        mail.CC='' #Mail ID goes here..
        mail.Subject = 'P2/P2 Tickets - '+p1p2_shift+' Shift IST'
        greet_str="Please find the list of P1/P2 Incidents from  "+p1p2_shift+"  Shift.<br>"

    # Infra Report: Adding current shift to mail subject, To and Cc Selection 
    elif (button_name=="infra"):
        mail.To='' #Mail ID goes here..
        mail.CC='' #Mail ID goes here..
        mail.Subject = 'Cognizant Command Center Ticket Status ( '+day_time+' Shift EST)'
        #Greeting line and Acitve Incidents table
        greet_str="""Greetings of the day!<br><br>Please find the below P1/P2 MIM status report:<br>
        <table width="100%">
            <tr bgcolor= "#191970">
            <th width="100" height="25" colspan="10" style= "text-align:left; color:azure" ><b>P1/P2 - Active</th>
            </tr>
            <tr bgcolor= "#B0C4DE">
            <th width="100" height="25" colspan="10" style= "text-align:left"><b>EST Date: """+infra_date+"""       Shift Covered: """+day_time+""" Shift EST</th>
            </tr>
            <tr width="100" height="25" colspan="10">
                <th colspan="3" style= "text-align:left">Projects</th>
                <th style= "text-align:center"> P1/P2 </th>
                <th style= "text-align:center">Track</th>
                <th style= "text-align:center">Status</th>
                <th style= "text-align:center">MIM Lead</th>
                <th style= "text-align:center" colspan="3">Comments</th>
            </tr>"""
        
        # Infra/ P1P2 Reports: Seperate table and data for Pluto
        Pluto=['INFRA SUPPORT PLUTO WINDOWS','INFRA SUPPORT PLUTO LINUX','INFRA SUPPORT PLUTO DBA']
        active_Pluto=df.loc[(df['assignment_group'].isin(Pluto)) & (df['state'].isin(['In Progress','On Hold'])), :]
        pluto_RC= len(active_Pluto)
        if (pluto_RC!=0):
            greet_str+="""<tr width="100" height="25" colspan="10">
            <td colspan="3" rowspan=" """+str(pluto_RC)+""" " style= "text-align:left"><b>Pluto</b><br>(Windows, Linux, DBA, App)
            </td>"""
            i=0
            for i in range(pluto_RC):
                if (active_Pluto['assignment_group'][i]=="INFRA SUPPORT PLUTO WINDOWS"):
                    track= 'Windows'
                elif (active_Pluto['assignment_group'][i]=="INFRA SUPPORT PLUTO LINUX"):
                    track= 'Linux'
                elif (active_Pluto['assignment_group'][i]=="INFRA SUPPORT PLUTO DBA"):
                    track= 'DBA'
                if(i==0):
                    row_sub=""
                else:
                    row_sub="<tr>"
                greet_str+=row_sub+"""
                <td style= "text-align:center">"""+active_Pluto['number'][i]+"""</td>
                <td style= "text-align:center">"""+track+"""</td>
                <td style= "text-align:center">"""+active_Pluto['state'][i]+"""</td>
                <td style= "text-align:center">  </td>
                <td style= "text-align:center" colspan="3">  </td>
                </tr>"""

        else:
                greet_str+="""<tr width="100" height="25" colspan="10">
                <td colspan="3" style= "text-align:left"><b>Pluto</b><br>(Windows, Linux, DBA, App)
                </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32" colspan="3">   </td>
                </tr>"""
        
        # Infra/ P1P2 Reports: Seperate table and data for Non-hosting
        active_NH=df.loc[(df['assignment_group']== 'APP SUPPORT INFRA NETWORKING GLOBAL') & (df['state'].isin(['In Progress','On Hold'])), :]
        NH_RC= len(active_NH)
        if (NH_RC!=0):
            greet_str+="""<tr width="100" height="25" colspan="10">
            <td colspan="3" rowspan=" """+str(NH_RC)+""" " style= "text-align:left"><b>Network - BT</b><br>(Network)
            </td>"""
            i=0
            track='Network'
            for i in range(NH_RC):
                if(i==0):
                    row_sub=""
                else:
                    row_sub="<tr>"
                greet_str+=row_sub+"""
                <td style= "text-align:center">"""+active_NH['number'][i]+"""</td>
                <td style= "text-align:center">"""+track+"""</td>
                <td style= "text-align:center">"""+active_NH['state'][i]+"""</td>
                <td style= "text-align:center">   </td>
                <td style= "text-align:center" colspan="3">   </td>
                </tr>"""
            greet_str+="</table>"
        else:    
            greet_str+="""<tr width="100" height="25" colspan="10">
                <td colspan="3" style= "text-align:left"><b>Network - BT</b><br>(Network)
                </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32">   </td>
                <td style= "text-align:center" bgcolor="#32CD32" colspan="3">   </td>
                </tr>
                </table>"""


    # Define the HTML for the table
    html = """<html>
    <head><style>
    table, tr, td, th{
        border: 1.5px solid black;
        border-collapse:collapse;
        text-align:center;
    }
    </style></head>
    <body>
    Hi All,<br><br>
    """+greet_str

    if(button_name=='p1p2' or button_name=='infra'):
        if (button_name=='p1p2'):
            title1="""Pluto:"""
            title2="""Non-Hosting:"""
        else:
            title1="""Pluto - P1 and P2 status:"""
            title2="""Non-Hosting Status:"""
        html+="""<br>
            <u><h3 style="color: fuchsia">"""+title1+"""</h3></u>"""
        if (button_name=='infra'):
            infraDF=df.loc[(~df['state'].isin(['In Progress','On Hold'])) & (df['assignment_group']!='APP SUPPORT INFRA NETWORKING GLOBAL'), :]
            row_count= len(infraDF)
            df=infraDF
        else:
            p1p2DF=df.loc[df['assignment_group']!='APP SUPPORT INFRA NETWORKING GLOBAL', :]
            row_count=len(p1p2DF)
            df=p1p2DF
        html+=calRows()+"""
            </table>
            <br>
            <u><h3 style="color: #002060">"""+title2+"""</h3></u>
            """
        nonHost_Rows=df.loc[df['assignment_group']== 'APP SUPPORT INFRA NETWORKING GLOBAL', :]
        row_count= len(nonHost_Rows)
        df=nonHost_Rows
        html+=calRows()+"""
        <br>
        """
    else:
        html+=calRows()
    
    html+="""
    </table>
    </body>
    </html><br><br>
    <b>Regards,<br>
    Cognizant Command Center.<br>
    CC Email :- <br>
    CC/MIM Phone :- </b>"""
    
    mail.HTMLBody = html
    mail.Display()


# Code for Image reference below to add Organization logo
#cog_img=customtkinter.CTkImage(dark_image=Image.open(""),size=(200,55))
#ImgLabel= customtkinter.CTkLabel(root,text=" " ,image=cog_img)
#ImgLabel.pack()
#bms_img=customtkinter.CTkImage(dark_image=Image.open(""),size=(30,50))
#ImgLabel= customtkinter.CTkLabel(root,text=" " ,image=bms_img)
#ImgLabel.pack(side="top")


#Alerts Frame
frame= customtkinter.CTkFrame(master=root)
frame.pack(pady=20,padx=60, fill="both",expand=True, side="left")

label= customtkinter.CTkLabel(master=frame, text="INCIDENT ALERTS", font=("Calibri Light",20))
label.pack(pady=12,padx=10)

ticketNo=customtkinter.CTkEntry(master=frame, placeholder_text="Ticket No.",font=("Calibri",15))
ticketNo.pack(pady=12,padx=10)

personContacted=customtkinter.CTkEntry(master=frame, placeholder_text="Person Contacted",font=("Calibri",15))
personContacted.pack(pady=12,padx=10)

button=customtkinter.CTkButton(master=frame, text="Generate",font=("Calibri",16,"bold") , command=lambda: singleAlert('single'))
button.pack(pady=12,padx=10)

#Multiple Alerts Frame within Alerts Frame
frame2= customtkinter.CTkFrame(master=frame)
frame2.pack(pady=15,padx=40, fill="both",expand=True)

label= customtkinter.CTkLabel(master=frame2, text="MULTIPLE ALERTS", font=("Calibri Light",20))
label.pack(pady=12,padx=10)

csvFile=customtkinter.CTkButton(master=frame2, text="Select .csv file",font=("Calibri",16,"bold") , command=lambda: openFile('multiple'))
csvFile.pack(pady=12,padx=10)

#frame for reports
frame3=frame= customtkinter.CTkFrame(master=root)
frame3.pack(pady=20,padx=60, fill="both",expand=True, side="right")
label= customtkinter.CTkLabel(master=frame3, text="REPORTS", font=("Calibri Light",20))
label.pack(pady=12,padx=10)

p1p2Button=customtkinter.CTkButton(master=frame3, text="P1/P2 Report",font=("Calibri",16,"bold") , command=lambda: openFile('p1p2'))
p1p2Button.pack(pady=12,padx=10)

InfraButton=customtkinter.CTkButton(master=frame3, text="Infra Report",font=("Calibri",16,"bold") , command=lambda: openFile('infra'))
InfraButton.pack(pady=12,padx=10)

#Frame for SMO within Reports Frame
frame4=frame= customtkinter.CTkFrame(master=frame3)
frame4.pack(pady=15,padx=40, fill="both",expand=True)
label= customtkinter.CTkLabel(master=frame4, text="SMO", font=("Calibri Light",20))
label.pack(pady=12,padx=10)

ChangeButton=customtkinter.CTkButton(master=frame4, text="Change Request",font=("Calibri",16,"bold") , command=lambda: openFile('change'))
ChangeButton.pack(pady=12,padx=10)

AgingButton=customtkinter.CTkButton(master=frame4, text="Aging",font=("Calibri",16,"bold") , command=lambda: openFile('aging'))
AgingButton.pack(pady=12,padx=10)

UpdatedButton=customtkinter.CTkButton(master=frame4, text="Updated",font=("Calibri",16,"bold") ,command=lambda: openFile('updated'))
UpdatedButton.pack(pady=12,padx=10)

creditLabel=customtkinter.CTkLabel(master=root, text="Created by Rahul Kamble", font=("Calibri",12))
creditLabel.pack(side="bottom")

root.mainloop()
