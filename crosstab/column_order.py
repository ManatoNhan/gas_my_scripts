# !pip install gspread pandas google-auth

import gspread
from google.colab import auth
from google.auth import default
import pandas as pd

# Google Colabで認証を行う
auth.authenticate_user()
creds, _ = default()

# gspreadクライアントを作成
gc = gspread.authorize(creds)

# 必要なパラメータ
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1e3SdHODewRGvyELYLSDGjJpcbSYiVyyiUSoiWAzzA3Q/edit?gid=1602405071#gid=1602405071'  # スプレッドシートのURL
sheet_name = 'アンケートraw'  # 使用するシート名
column_group1_start = 'E'  # 最初の列グループの開始列
column_group1_end = 'E'    # 最初の列グループの終了列
column_group2_start = 'H'  # 二番目の列グループの開始列
column_group2_end = 'Q'    # 二番目の列グループの終了列

# 列名をインデックスに変換する関数
def col2num(col):
    num = 0
    for c in col:
        if 'A' <= c <= 'Z':
            num = num * 26 + (ord(c) - ord('A')) + 1
    return num

# スプレッドシートを開く
spreadsheet = gc.open_by_url(spreadsheet_url)
worksheet = spreadsheet.worksheet(sheet_name)

# データを取得
data = worksheet.get_all_values()

# データをDataFrameに変換
df = pd.DataFrame(data[1:], columns=data[0])

# グループの列インデックスを取得
group1_start_idx = col2num(column_group1_start) - 1
group1_end_idx = col2num(column_group1_end)
group2_start_idx = col2num(column_group2_start) - 1
group2_end_idx = col2num(column_group2_end)

# グループの列を選択
column_group1 = df.iloc[:, group1_start_idx:group1_end_idx]
column_group2 = df.iloc[:, group2_start_idx:group2_end_idx]

# グループ1のすべての選択肢をフラット化
group1_flat = column_group1.values.flatten()
# グループ2のすべての選択肢をフラット化
group2_flat = column_group2.values.flatten()

# ユニークな選択肢を取得し、ソート
group1_unique = sorted(pd.Series(group1_flat).dropna().unique(), key=lambda x: (x != 'その他', x))
group2_unique = sorted(pd.Series(group2_flat).dropna().unique(), key=lambda x: (x != 'その他', x))

# クロス集計用のDataFrameを初期化
cross_tab = pd.DataFrame(0, index=group1_unique, columns=group2_unique)

# クロス集計を行う
for i in range(len(df)):
    group1_items = column_group1.iloc[i].dropna()
    group2_items = column_group2.iloc[i].dropna()
    for item1 in group1_items:
        for item2 in group2_items:
            cross_tab.at[item1, item2] += 1

# 新しいシート名を作成
new_sheet_name = f"{column_group1_start}_{column_group1_end}x{column_group2_start}_{column_group2_end}_cross_tab"

# 新しいシートを追加してクロス集計結果を書き込む
try:
    new_worksheet = spreadsheet.add_worksheet(title=new_sheet_name, rows=str(cross_tab.shape[0] + 1), cols=str(cross_tab.shape[1] + 1))
except gspread.exceptions.APIError:
    print(f"シート名 '{new_sheet_name}' は既に存在します。別の名前を使用してください。")
    new_worksheet = None

if new_worksheet:
    # ヘッダー行を設定
    header = [""] + cross_tab.columns.tolist()
    rows = cross_tab.reset_index().values.tolist()
    new_worksheet.update([header] + rows)

    print(f"クロス集計結果は新しいシート '{new_sheet_name}' に保存されました。")
else:
    print("クロス集計結果の保存に失敗しました。")
