import openpyxl
import datetime
import win32com.client as win32
import time

# Excelファイルのパス
excel_file_path = r"C:\Users\齊藤未来\xlautomate\data\bbdata.xlsx"

# 毎日の実行時刻
target_time = datetime.time(15, 42, 0)  # 例: 9時0分0秒

# Excelファイルからセルの値を取得
def read_excel_cell(file_path, sheet_name, cell_address):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    cell_value = sheet[cell_address].value
    wb.close()
    return cell_value

# メールの送信
def send_email(subject, body, to_recipients, cc_recipients=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.To = to_recipients
    if cc_recipients:
        mail.CC = cc_recipients
    mail.Send()

# 指定時間にメールを送信する関数
def send_email_at_target_time():
    current_time = datetime.datetime.now().time()
    if current_time >= target_time:
        # ExcelファイルからA2セルとB2セルの値を取得
        value_a2 = read_excel_cell(excel_file_path, 'Sheet1', 'A2')
        value_b2 = read_excel_cell(excel_file_path, 'Sheet1', 'B2')

        # メール本文を作成
        body = f"本日は{value_a2}と{value_b2}です。"

        # メール送信
        send_email("件名", body, "to@example.com", "cc@example.com")

# プログラム実行
while True:
    current_datetime = datetime.datetime.now()
    if current_datetime.time() >= target_time:
        # 指定時間以降に実行された場合

        # 送信するメールを送信時刻の翌日に設定
        send_time = current_datetime + datetime.timedelta(days=1)
        send_time = send_time.replace(hour=target_time.hour, minute=target_time.minute, second=target_time.second)

        # 送信時刻まで待機
        time.sleep((send_time - current_datetime).total_seconds())

        # メール送信
        send_email_at_target_time()

    # 次の日の実行時刻まで待機
    tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)
    tomorrow = tomorrow.replace(hour=target_time.hour, minute=target_time.minute, second=target_time.second)
    time.sleep((tomorrow - datetime.datetime.now()).total_seconds())