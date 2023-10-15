import pandas as pd
import xlwings as xw
import numpy as np

# Aging/Updated date Reports: Pivot Table creation
def createPivotTable(task_type,b_name):
    #Path to the pre-modified file from GUI.py file
    file_name="C:\\Users\\2106624\\Downloads\\modified.xlsx"
    new_wb=xw.Book(file_name)

    if task_type == 'Incident':
        sheet_value='PivotIncident'
    else:
        sheet_value='PivotTask'

    # Task Type & State filter conditions
    pivot_df= pd.read_excel(file_name, sheet_name='Sheet1')
    filtered_df=pivot_df[(pivot_df['State'] !='Resolved') & (pivot_df['Task type']== task_type)].copy()

    if filtered_df.empty:
        no_str='No Tickets'
        return no_str
    else:
        no_str='pass'

    #Pivot fields as seen in Excel
    column=["Range"]
    rows=["Assignment group","Assigned to"]
    value=["Number"]

    # Grouping Days from 0 to 10 by 5
    if b_name == 'aging':
        aging_bin= [0,5,10, float('inf')]
        aging_label=['0-5','5-10','>10']
        filtered_df.loc[:,'Range']= pd.cut(filtered_df['Aging in days'], bins=aging_bin, labels=aging_label)
    
    # Grouping Days from 0 to 10 by 2
    elif b_name== 'updated':
        aging_bin= [0,2,4,6,8,10,float('inf')]
        aging_label=['0-2','2-4','4-6','6-8','8-10','>10']
        filtered_df.loc[:,'Range']= pd.cut(filtered_df['Updated in days'], bins=aging_bin, labels=aging_label)
        print("0-2 filtered")

    #Final Pivot Dataframe
    final_pivot=pd.pivot_table(filtered_df, value,index=rows, columns=column, aggfunc=np.count_nonzero, margins=True, margins_name='Grand Total')
    print("final pivot created")

    # Ticket count for individual Assignment Groups
    grouped= filtered_df.groupby('Assignment group')['Number'].nunique()
    grouped_dict=grouped.to_dict()
    final_pivot['Group Count']=final_pivot.index.get_level_values('Assignment group').map(grouped_dict)
    print("Grouped according to AG")

    # Closing xlwings functions to use ExcelWriter
    new_wb.save(file_name) 
    new_wb.close()

    # Adding Pivot sheet to excel
    with pd.ExcelWriter(file_name, mode='a') as writer:
        if sheet_value in writer.book.sheetnames:
            writer.book.remove(writer.book[sheet_value])
        final_pivot.to_excel(writer,sheet_name=sheet_value)
        print("added final sheet")

    new_wb=xw.Book(file_name)
    pivot_sheet=new_wb.sheets[sheet_value]

    pivot_sheet.autofit()
    
    #Adding all borders
    used_range=pivot_sheet.used_range
    for cell in used_range:
        cell.api.Borders.Weight=2

    last_row=pivot_sheet.range('A'+ str(pivot_sheet.cells.last_cell.row)).end('up').row
    fill_color=(176,196,222)
    pivot_sheet.range('1:2').color=fill_color
    pivot_sheet.range('A3:B3').color=fill_color
    pivot_sheet.range(str(last_row)+":"+str(last_row)).color=fill_color
    fill_color=(196,222,176)
    pivot_sheet.range("A4:A"+str((last_row)-1)).color=fill_color

    new_wb.save(file_name)
    new_wb.close()
    return no_str


