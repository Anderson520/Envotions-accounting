import win32com.client as win32
import psutil
import os
import subprocess
import pandas as pd
import os
import pdb

def send_notification(email, subject, message, file):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = subject
    mail.GetInspector 
    fpath = f"{os.getcwd()}\{file}"
    print(f"File path: {fpath}")
    mail.Attachments.Add(fpath)
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + message + mail.HTMLbody[index + 1:] 
    #mail.Display(True)
    mail.Send()


# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
def open_outlook(path):
    try:
        subprocess.call([path])
        os.system(path);
    except:
        print("請確認你的Outlook有打開")

# Checking if outlook is already opened. If not, open Outlook.exe and send email
for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0

df=pd.read_excel('data.xlsx', sheet_name=['config', 'employee', 'subject', 'body'])
employee_cnt = len(df['employee'].values)
path = f"{df['config'].values[0][0]}"
#pdb.set_trace()
for ii in range(0,employee_cnt):
    print(f"姓名: {df['employee'].values[ii][0]}")
    print(f"郵箱: {df['employee'].values[ii][1]}")
    print(f"Subject: {df['subject'].values[0][0]}")
    mail = f"{df['employee'].values[ii][1]}"
    subject = f"{df['subject'].values[0][0]}"
    body = f"{df['body'].values[0][0]}"
    file = f"{df['employee'].values[ii][0]}.pdf"
    if (flag == 1):
        send_notification(mail, subject, body, file)
    else:
        open_outlook(path)
        send_notification(mail, subject, body, file)
    print("====")