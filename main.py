import datetime
import subprocess
import time
import openpyxl
import win32com.client as win32

def open_excel_at_time(excel_file, target_time):
    while True:
        current_time = datetime.datetime.now().time()
        if current_time >= target_time:
            subprocess.Popen(['start', excel_file], shell=True)
            break
        time.sleep(1)

def create_outlook_draft(recipient, cc_recipient, subject, body):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.To = recipient
    mail.CC = cc_recipient
    mail.Save()

if __name__ == '__main__':
    # Excelファイルのパスを指定してください
    excel_file = r'C:\Users\齊藤未来\xlautomate\data\bbdata.xlsx'

    # 実行したい時刻を指定してください (24時間形式)
    target_time = datetime.time(15, 0)  # 例: 15時0分

    # メールの設定
    recipient = 'recipient@example.com'  # 宛先のメールアドレス
    cc_recipient = 'cc_recipient@example.com'  # CCのメールアドレス
    subject = 'メールの件名'
    body_template = '本日は{}、{}となります'

    open_excel_at_time(excel_file, target_time)

    while True:
        try:
            workbook = openpyxl.load_workbook(excel_file)
            sheet = workbook.active
            cell_a1_value = sheet['A2'].value
            cell_a2_value = sheet['B2'].value
            if cell_a1_value and cell_a2_value:
                body = body_template.format(cell_a1_value, cell_a2_value)
                create_outlook_draft(recipient, cc_recipient, subject, body)
                break
        except:
            pass

        time.sleep(1)
