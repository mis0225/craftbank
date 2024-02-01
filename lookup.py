"""
人物ID統合済み"配信中"ファイルの"ID"列をと"会社名"列をキーにして"PD全データ"ファイルの人物ID列と組織名を参照し、
マッチしたら"PD"ファイルの該当する"組織ID"列のデータを、
新しい"ID"列に作成してその列にデータを格納する
"""


import pandas as pd

# "配信中"の人物ID紐づけ済みファイルを読み込む
haisihin_file_path = r"C:\Users\mikus\Downloads\東京許可業者.xlsx"
haishin_file = pd.read_excel(haisihin_file_path)

# "PD"ファイルを読み込む
pd_file_path = r"C:\Users\mikus\Downloads\people-15402717-1894.xlsx"
pd_file = pd.read_excel(pd_file_path)

# 新しい "Company_ID" 列を作成
haishin_file['職人酒場'] = ''

# "配信中"ファイルの"人物ID"列をループで処理
for index, row in haishin_file.iterrows():
    # 処理している行番号を表示
    print(f"Processing row {index + 1}/{len(haishin_file)}")

    # "配信中"ファイルの人物ID列、会社名列の値を取得
    id_to_match = row['商号名称']

    # "PD"ファイルでマッチングする行を取得
    matched_row = pd_file[(pd_file['人物 - 組織'] == id_to_match)]

    # マッチングした行が存在する場合、"PD"ファイルの"id"列のデータを、新しい "ID" 列にコピー
    if not matched_row.empty:
        haishin_file.at[index, '職人酒場'] = matched_row['参加'].values[0]

# 結果を"配信中"ファイルに保存
haishin_file.to_excel(r"C:\Users\mikus\Downloads\5_addsakaba.xlsx", index=False)