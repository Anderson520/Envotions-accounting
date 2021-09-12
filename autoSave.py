from win32com import client
import win32api
import pathlib
import pandas as pd
import pdb
import time

df=pd.read_excel('data.xlsx', sheet_name=['employee'])
employee = []
employee_cnt = len(df['employee'].values)
for ii in range(0,employee_cnt):
    employee.append(df['employee'].values[ii][0].strip())


excel_file = "薪資表.xlsx"

excel_path = str(pathlib.Path.cwd() / excel_file)


excel = client.DispatchEx("Excel.Application")
excel.Visible = 0

wb = excel.Workbooks.Open(excel_path)
#pdb.set_trace()
#ws = wb.Worksheets[12]

for id in range(1, len(wb.Worksheets)):
    try:
        name = wb.Worksheets[id].name.strip()
        print(f"分頁:{name}")
        if name in employee:
            pdf_file = name+".pdf"
            print(f"Save {pdf_file}")
            pdf_path = str(pathlib.Path.cwd() / pdf_file)
            wb.Worksheets[id].SaveAs(pdf_path, FileFormat=57)
    
    except Exception as e:
        print("Failed to convert")
        print(str(e))

wb.Close()
excel.Quit()
