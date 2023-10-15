import xlwings as xw
import PIL.ImageGrab as ImageGrab
import os
import csv
import pandas as pd

import subprocess
import pymem
import datetime
import time
import win32com.client as win32
import win32clipboard
import pytz

def clearClipBoard():
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()


# Save Excel data as image 
def grabImage(b_name,token):

    file_name="C:\\Users\\2106624\\Downloads\\modified.xlsx"
    wb=xw.Book(file_name)  
    def chk_process(process_name):
        progs=str(subprocess.check_output('tasklist'))
        if process_name in progs:
            return True
        else:
            return False

    if chk_process('OUTLOOK.EXE'):    
        subprocess.run(["TASKKILL","/F","/IM","OUTLOOK.EXE"])

    if b_name== 'change':
        if token == "empty":
            change1="""<b> *No tickets* </b>"""
            change2=""" """
        else:
            while True:
                try:
                    if 'Sheet2' in wb.sheet_names:
                        sheet=wb.sheets['Sheet2']
                        range=sheet.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","Change2.png")
                        img_file=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        change1=f"""<img src='{img_file}' alt="No Tickets" width="100%" style="height:auto;" >"""
                    else:
                        change1="""<b> *No tickets* </b>"""
                    break
                except Exception:
                    print("ImgGrab exception")

            while True:
                try:
                    if 'Sheet1' in wb.sheet_names:
                        sheet=wb.sheets['Sheet1']
                        range=sheet.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","Change1.png")
                        img_file1=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        change2=f"""<img src='{img_file1}' alt="No Tickets" width="50%" style="height:auto;" >"""
                    else:
                        change2="""<b> *No tickets* </b>"""
                    break
                except Exception:
                    print("ImgGrab exception")


    if b_name in ['aging','updated']:
        # Checking if sheet names exist then saving images
        if token == "empty":
            iPiv,tPiv="""<b> *No tickets* </b>"""
            iSheet,tSheet=""" """
        else: 
            while True:
                try:       
                    if 'PivotIncident' in wb.sheet_names:
                        sheet=wb.sheets['PivotIncident']
                        range=sheet.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","INC_Piv.png")
                        img_inc_piv=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        iPiv= f"""<img src='{img_inc_piv}' alt=" " width="50%" style="height:auto;" >"""
                    else:
                        iPiv= """<b>No Tickets</b>"""
                    break
                except Exception:
                    print("ImgGrab exception")            

            while True:
                try:
                    if 'INCSheet' in wb.sheet_names:
                        sheet1=wb.sheets['INCSheet']
                        range=sheet1.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","INC_sheet.png")
                        img_inc_sheet=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        iSheet= f"""<img src='{img_inc_sheet}' alt=" " width="100%" style="height:auto;" >"""
                    else:
                        iSheet= """ """
                    break
                except Exception:
                    print("ImgGrab exception")  

            while True:
                try:    
                    if 'PivotTask' in wb.sheet_names:
                        sheet2=wb.sheets['PivotTask']
                        range=sheet2.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","Task_Piv.png")
                        img_task_piv=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        tPiv= f"""<img src='{img_task_piv}' alt=" "width="50%" style="height:auto;" >"""            
                    else:
                        tPiv= """<b>No Tickets</b>"""
                    break
                except Exception:
                    print("ImgGrab exception")  

            while True:
                try: 
                    if 'TaskSheet' in wb.sheet_names:
                        sheet3=wb.sheets['TaskSheet']
                        range=sheet3.used_range
                        range.api.Copy()
                        img= ImageGrab.grabclipboard()
                        file= os.path.join("C:\\Users\\2106624\\Downloads\\","Task_sheet.png")
                        img_task_sheet=os.path.normpath(file).replace("\\","\\\\")
                        img.save(file)
                        tSheet= f"""<img src='{img_task_sheet}' alt=" " width="100%" style="height:auto;" >"""
                    else:
                        tSheet= """ """
                    break
                except Exception:
                    print("ImgGrab exception")
    wb.close()

    # Calculating date for reports    
    utc_now=datetime.datetime.utcnow()
    ist= pytz.timezone('Asia/Kolkata')
    ist_now=utc_now.replace(tzinfo=pytz.utc).astimezone(ist)
    smo_date=ist_now.strftime('%d-%B-%Y')

    # Mail object
    outlook=win32.Dispatch('outlook.application')
    mail=outlook.CreateItem(0)
    mail.To="BMSPlutoInfraLinux@cognizant.com; BMSPlutoInfraWindows@cognizant.com; BMSPlutoInfraSMO@cognizant.com; COGBMSCommandCenter@cognizant.com; BMSPlutoInfraDBA@cognizant.com; MG-BMS-PLUTO-LINUX@bms.com; MG-BMS-PLUTO-WINDOWS@bms.com; MG-BMS-PLUTO-DBA@bms.com; MG-BMS-PLUTO-SMO@bms.com; MG-BMS-PLUTO-COMMANDCENTER@bms.com; Nicholas.Pereira@cognizant.com"
    mail.HTMLBody="""<span style="background-color: #FFFF00">** This is a test mail, generated programmatically including all its contents **</span><br><br>"""
    mail.Subject="Test Automated: "
    if b_name=='change':
        mail.Subject+="Pluto Change Request closure before Planned end date"
        mail.HTMLBody+=f"""<html><body>Hi Team,<br><br>Please ensure that you close the <b>Change request</b>
        before the <b>Planned end date and time.</b><br><br>{change1}<br><br>{change2}<br><br>"""

    elif b_name=='aging':
        mail.Subject+="Pluto Aging report as of "+smo_date
        mail.HTMLBody+=f"""<html><body>Hi Team,<br><br>Please find below the snapshots for the aging report for Incident and Catalog Tasks 
        as of """+smo_date+f""". Kindly make sure notes tab is updated with all the information with each day progress, and even if you 
        are waiting for other team to complete the task.<br><br><u><b>Incident</b></u><br><br>{iPiv}<br><br>
        {iSheet}<br><br><u><b>Catalog Task</b></u><br><br>{tPiv}
        <br><br>{tSheet}<br><br>"""
    
    else:
        mail.CC="Eswar.Singh@cognizant.com; Souvik.Mitra@bms.com; Arijit.Sinha@cognizant.com"
        mail.Subject+="Updated date report as of "+smo_date
        mail.HTMLBody+=f"""<html><body>Hi Team,<br><br>Please find below the snapshots for the last updated report for Incident and Catalog Tasks. Also, find the details of the Incident and Catalog Tasks that have not been updated for more than 1 day.
        <br>As mentioned kindly update the Incidents, Change task and Catalog Task on a daily basis and also update the Configuration item without failure.
        <br><br><u><b>Incident</b></u><br><br>{iPiv}<br><br>
        {iSheet}<br><br><u><b>Catalog Task</b></u><br><br>{tPiv}
        <br><br>{tSheet}<br><br>Please see below status on tickets received from last 24 hrs.<br><br>"""

    mail.HTMLBody+="""
        <br><b>Regards,<br>
        Cognizant Command Center.<br>
        CC Email :- COGBMSCommandCenter@cognizant.com<br>
        CC/MIM Phone :- +91 887-092-2960 /+91 770-806-927</b>
    </body></html>"""
    mail.Display()

