import xlwings as xw
import win32com.client as win32
import pandas as pd
import pyautogui

def smo_Reports():
    #Opening excel file 
    workbook=xw.Book("C:\\Users\\2106624\\Downloads\\task (1).xlsx", read_only=True)
    worksheet=workbook.sheets.active

    #Reading data from old excel file to new
    data= worksheet.used_range.value
    new_wb=xw.Book()
    new_sheet=new_wb.sheets.active
    new_sheet.range('A1').value = data

    #Adding 'Aging in days' column
    last_col=new_sheet.api.UsedRange.Columns.Count
    new_sheet.api.Columns(last_col+1).Insert()
    new_sheet.range((1,last_col+1)).value='Aging in days'
    new_sheet.range('M2').formula= '=NOW()-J2'
    cell= new_sheet.range('M2')
    cell.number_format = '#,##0.00;[Red](#,##0.00);0.00;@'

    #calculating and adding values to 'Aging in days'
    last_row=new_sheet.range('L'+ str(new_sheet.cells.last_cell.row)).end('up').row
    new_sheet.range('M2:M'+str(last_row)).api.FillDown()
    #color= new_sheet.range("M2:M" + str(last_row))


    #Sorting 'Aging in days' in descending order
    range_to_sort= new_sheet.used_range
    data= range_to_sort.options(pd.DataFrame, header=1, index=False).value

    sorted_data= data.sort_values(by='Aging in days', ascending=False)

    new_sheet.range('A2').value=sorted_data.values.tolist()

    #for coloring aging days greater than 5 red
    range_to_sort = new_sheet.range("M2:M" + str(last_row)) 
    values=[cell.value for cell in range_to_sort]

    for i,value in enumerate(values):
        if new_sheet['M'+str(i+2)].value >=5:
            new_sheet['M'+str(i+2)].api.Font.ColorIndex= 3

    #Adding bold font & blue fill to first row
    new_sheet.range('1:1').api.Font.Bold=True
    fill_color=(176,196,222)
    new_sheet.range('1:1').color=fill_color

    #Adding all borders
    used_range=new_sheet.used_range
    for cell in used_range:
        cell.api.Borders.Weight=2

    #Autofitting & resizing the data 
    new_sheet.autofit()
    #new_sheet.range('A1').column_width =15
    new_sheet.range('B1').column_width =15
    new_sheet.range('C1').column_width =6
    new_sheet.range('D1').column_width =7
    new_sheet.range('E1').column_width =10
    new_sheet.range('F1').column_width =6
    new_sheet.range('G1').column_width =15
    #new_sheet.range('H1').column_width =15
    #new_sheet.range('I1').column_width =15
    #new_sheet.range('J1').column_width =15
    #new_sheet.range('K1').column_width =15
    #new_sheet.range('L1').column_width =15


    #for row in new_sheet.cells:
    #    for cell in row:
    #        cell_value=str(cell.value)
    #        if '\n' in cell_value:
    #            cell.wrap()
    #        elif len(cell_value)> len(str(cell.api.Value)):
    #            cell.wrap()

    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey("ctrl","c")

    sheet2= new_wb.sheets.add(name='Sheet2')

    pyautogui.hotkey("ctrl","v")

    #range_to_sort = sheet2.range("A2:M" + str(last_col))
    range_to_sort= sheet2.used_range
    data= range_to_sort.options(pd.DataFrame, header=1, index=False).value

    sorted_data= data.sort_values(by='Planned end date')

    sheet2.range('A2').value=sorted_data.values.tolist()
    #sorted_values=sorted(values)
    #for i,value in enumerate(sorted_values):
    #    sheet2['L'+str(i+2)].value=value

    sheet2.autofit()
    sheet2.range('B1').column_width =15
    sheet2.range('C1').column_width =6
    sheet2.range('D1').column_width =7
    sheet2.range('E1').column_width =10
    sheet2.range('F1').column_width =6
    sheet2.range('G1').column_width =15


    new_wb.save("C:\\Users\\2106624\\Downloads\\modified.xlsx")

    workbook.close()


    outlook=win32.Dispatch('outlook.application')
    mail=outlook.CreateItem(0)
    mail.To="example"
    mail.Subject="Pluto Change Request closure before Planned end date"
    mail.HTMLBody="""<html><body>Hi Team,<br>Please ensure that you close the <b>Change request</b>
    before the <b>Planned end date and time.</b><br>
        <br><b>Regards,<br>
        Cognizant Command Center.<br>
        CC Email :- <br>
        CC/MIM Phone :- </b>
    </body></html>"""
    mail.Display(True)
