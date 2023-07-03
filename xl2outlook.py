import xlwings as xw
import win32com.client as win32

# エクセルファイルのパスを設定
excel_file_path = r'C:\Users\齊藤未来\xlautomate\data\bbdata.xlsx'

# Excelアプリケーションを起動してブックを開く
app = xw.App(visible=False)
workbook = app.books.open(excel_file_path)

# シートを選択してA2とB2のセルの値を取得
sheet = workbook.sheets['Sheet1']
cell_a2_value = sheet.range('A2').value
cell_b2_value = sheet.range('B2').value

# エクセルアプリケーションを終了
app.quit()

# Outlookの下書きを作成
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.Subject = 'メールの件名をここに入力してください'
mail.Body = f'本日は{cell_a2_value}と{cell_b2_value}です。'

# Outlookウィンドウを表示
mail.Display(True)
