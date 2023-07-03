import datetime
import subprocess
import time

def open_excel_at_time(excel_file, target_time):
    while True:
        current_time = datetime.datetime.now().time()
        if current_time >= target_time:
            subprocess.Popen(['start', excel_file], shell=True)
            break
        time.sleep(1)

if __name__ == '__main__':
    # Excelファイルのパスを指定してください
    excel_file = 'C:\\path\\to\\your\\excel_file.xlsx'

    # 実行したい時刻を指定してください (24時間形式)
    target_time = datetime.time(15, 00)  # 例: 15時0分

    open_excel_at_time(excel_file, target_time)
