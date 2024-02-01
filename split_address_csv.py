import pandas as pd

# CSVファイルのパス
csv_file_path = r"C:\Users\mikus\Downloads\BLASTMAILerror_craftbank (1).csv"

# CSVファイルを読み込む（Shift_JISエンコーディングを指定）
df = pd.read_csv(csv_file_path, encoding='shift_jis')

# 新しいデータフレームを作成
new_df = df.copy()

# E-Mail列の各セルに対して処理を行う
for index, row in df.iterrows():
    # 実際の列名に変更する
    email_addresses = str(row['E-Mail'])

    # カンマで分割
    split_addresses = email_addresses.split(',')

    # E-Mail列にアドレスを格納（カンマで分割した1番目）
    new_df.at[index, 'E-Mail'] = split_addresses[0] if len(split_addresses) > 0 else ''

    # email2列にアドレスを格納（カンマで分割した2番目以降）
    new_df.at[index, 'email2'] = ",".join(split_addresses[1:]) if len(split_addresses) > 1 else ''

# 処理したデータを新しいCSVファイルに保存
output_csv_path = r"C:\Users\mikus\Downloads\error1_split.csv"
new_df.to_csv(output_csv_path, index=False, encoding='shift_jis')

print(f"処理が完了しました。結果は {output_csv_path} に保存されています。")
