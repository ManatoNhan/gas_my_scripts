# ===========================================
# ライブラリインストール（初回実行時のみ）
# ===========================================
!pip install gspread pandas google-auth

# ===========================================
# パラメータ設定（ここを編集してください）
# ===========================================
SPREADSHEET_URL = ""
SHEET_NAME = ""

# ===========================================
# 以下、関数定義（変更不要）
# ===========================================

import pandas as pd
import gspread
from google.auth import default
from google.colab import auth
import re
from collections import defaultdict

def authenticate_google():
    """Google認証を行う"""
    auth.authenticate_user()
    creds, _ = default()
    return gspread.authorize(creds)

def load_survey_data(spreadsheet_url, sheet_name):
    """スプレッドシートからアンケートデータを読み込む"""
    gc = authenticate_google()

    # URLからスプレッドシートIDを抽出
    sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]
    workbook = gc.open_by_key(sheet_id)
    worksheet = workbook.worksheet(sheet_name)

    # データを取得してDataFrameに変換
    data = worksheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])  # 1行目をヘッダーとして使用

    # 3行目以降を集計対象とする（2行目はタイトル行のため除外）
    df_for_analysis = df.iloc[1:].reset_index(drop=True)  # 2行目以降（元の3行目以降）

    return df, df_for_analysis, workbook

def identify_question_columns(df):
    """質問カラムを識別し、単一回答と複数回答に分類する"""
    question_cols = [col for col in df.columns if col.startswith('q_')]

    single_answer = []
    multiple_answer = defaultdict(list)

    # 質問番号ごとにグループ化
    question_groups = defaultdict(list)
    for col in question_cols:
        parts = col.split('_')
        if len(parts) >= 2 and parts[1].isdigit():
            q_num = parts[1]  # 質問番号
            question_groups[q_num].append(col)

    # 各質問グループを分析
    for q_num, cols in question_groups.items():
        if len(cols) == 1:
            # 1つしかない場合は単一回答
            single_answer.append(cols[0])
        else:
            # 複数ある場合は複数回答として処理
            base_q = f"q_{q_num}"

            # 数字のみのもの（q_3_1, q_3_2など）と
            # 長い質問文のもの（q_3_長い質問文）を両方含める
            for col in cols:
                multiple_answer[base_q].append(col)

    return single_answer, multiple_answer

def get_question_title(df, question_col):
    """質問のタイトル（質問カラム名_2行目の内容）を取得"""
    if len(df) > 0:  # 2行目があることを確認
        # 2行目（インデックス0、元データの2行目）から質問タイトルを取得
        title_row = df.iloc[0]  # dfは既に2行目以降なので、iloc[0]が元の2行目
        if pd.notna(title_row[question_col]) and title_row[question_col] != '':
            # 質問番号_質問タイトルの形式で返す
            parts = question_col.split('_')
            if len(parts) >= 2:
                q_num = f"q_{parts[1]}"
                return f"{q_num}_{title_row[question_col]}"
            else:
                return f"{question_col}_{title_row[question_col]}"
    return question_col

def process_single_answer_crosstab(df, question_col):
    """単一回答のクロス集計を処理"""
    # 空でない回答のみを対象
    valid_responses = df[df[question_col].notna() & (df[question_col] != '')]

    if len(valid_responses) == 0:
        return pd.DataFrame()

    # 回答の種類数をカウント
    unique_answers = valid_responses[question_col].nunique()

    # 回答数が20を超える場合はFA判定でスキップ
    if unique_answers > 20:
        print(f"  {question_col}: 回答数{unique_answers}個のためFA判定でスキップ")
        return pd.DataFrame()

    # クロス集計
    crosstab = pd.crosstab(
        [valid_responses['gender'], valid_responses['age_range']],
        valid_responses[question_col],
        margins=True
    )

    # インデックスを整理
    crosstab.index.names = ['性別', '年代']

    return crosstab

def process_gender_crosstab(df, question_col):
    """性別×回答のクロス集計を処理"""
    # 空でない回答のみを対象
    valid_responses = df[df[question_col].notna() & (df[question_col] != '')]

    if len(valid_responses) == 0:
        return pd.DataFrame()

    # 回答の種類数をカウント
    unique_answers = valid_responses[question_col].nunique()

    # 回答数が20を超える場合はFA判定でスキップ
    if unique_answers > 20:
        print(f"  {question_col}: 回答数{unique_answers}個のためFA判定でスキップ")
        return pd.DataFrame()

    # クロス集計（性別×回答）
    crosstab = pd.crosstab(
        valid_responses['gender'],
        valid_responses[question_col],
        margins=True
    )

    return crosstab

def process_age_crosstab(df, question_col):
    """年代×回答のクロス集計を処理"""
    # 空でない回答のみを対象
    valid_responses = df[df[question_col].notna() & (df[question_col] != '')]

    if len(valid_responses) == 0:
        return pd.DataFrame()

    # 回答の種類数をカウント
    unique_answers = valid_responses[question_col].nunique()

    # 回答数が20を超える場合はFA判定でスキップ
    if unique_answers > 20:
        print(f"  {question_col}: 回答数{unique_answers}個のためFA判定でスキップ")
        return pd.DataFrame()

    # クロス集計（年代×回答）
    crosstab = pd.crosstab(
        valid_responses['age_range'],
        valid_responses[question_col],
        margins=True
    )

    return crosstab

def process_multiple_answer_crosstab(df, question_group, question_cols):
    """複数回答のクロス集計を処理"""
    # カラムを適切な順序でソート
    def sort_multiple_cols(cols):
        def get_col_sort_key(col):
            parts = col.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])  # q_3_1 -> 1
                else:
                    return 0  # q_3_長い質問文 -> 0
            return 999
        return sorted(cols, key=get_col_sort_key)

    sorted_question_cols = sort_multiple_cols(question_cols)
    results = []

    for col in sorted_question_cols:
        # 空でない回答のみを対象
        valid_responses = df[df[col].notna() & (df[col] != '')]

        if len(valid_responses) == 0:
            continue

        # 各選択肢ごとの集計
        for idx, row in valid_responses.iterrows():
            results.append({
                '性別': row['gender'],
                '年代': row['age_range'],
                '回答内容': row[col],
                '選択肢': col
            })

    if not results:
        return pd.DataFrame()

    result_df = pd.DataFrame(results)

    # クロス集計
    crosstab = pd.crosstab(
        [result_df['性別'], result_df['年代']],
        [result_df['選択肢'], result_df['回答内容']],
        margins=True
    )

    # カラムを数値順に並び替え
    if isinstance(crosstab.columns, pd.MultiIndex):
        # レベル0（選択肢）でソート
        def sort_col_key(col_tuple):
            col_name = col_tuple[0]  # 選択肢名
            parts = col_name.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])
                else:
                    return 0  # 長い質問文は0番
            return 999

        # カラムをソート
        sorted_cols = sorted(crosstab.columns, key=sort_col_key)
        crosstab = crosstab.reindex(columns=sorted_cols)

    return crosstab

def process_multiple_answer_gender_crosstab(df, question_group, question_cols):
    """複数回答の性別×回答のクロス集計を処理"""
    # カラムを適切な順序でソート
    def sort_multiple_cols(cols):
        def get_col_sort_key(col):
            parts = col.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])  # q_3_1 -> 1
                else:
                    return 0  # q_3_長い質問文 -> 0
            return 999
        return sorted(cols, key=get_col_sort_key)

    sorted_question_cols = sort_multiple_cols(question_cols)
    results = []

    for col in sorted_question_cols:
        # 空でない回答のみを対象
        valid_responses = df[df[col].notna() & (df[col] != '')]

        if len(valid_responses) == 0:
            continue

        # 各選択肢ごとの集計
        for idx, row in valid_responses.iterrows():
            results.append({
                '性別': row['gender'],
                '回答内容': row[col],
                '選択肢': col
            })

    if not results:
        return pd.DataFrame()

    result_df = pd.DataFrame(results)

    # クロス集計（性別×選択肢×回答内容）
    crosstab = pd.crosstab(
        result_df['性別'],
        [result_df['選択肢'], result_df['回答内容']],
        margins=True
    )

    # === カラムを数値順に並び替え ===
    if isinstance(crosstab.columns, pd.MultiIndex):
        def sort_col_key(col_tuple):
            col_name = col_tuple[0]  # 選択肢名
            parts = col_name.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])  # q_3_1 -> 1
                else:
                    return 0  # 長い質問文は0番
            return 999

        # カラムを完全に再ソート
        sorted_cols = sorted(crosstab.columns.tolist(), key=sort_col_key)
        crosstab = crosstab.reindex(columns=sorted_cols)
        print(f"  性別×回答: カラムをソートしました")

    return crosstab

def process_multiple_answer_age_crosstab(df, question_group, question_cols):
    """複数回答の年代×回答のクロス集計を処理"""
    # カラムを適切な順序でソート
    def sort_multiple_cols(cols):
        def get_col_sort_key(col):
            parts = col.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])  # q_3_1 -> 1
                else:
                    return 0  # q_3_長い質問文 -> 0
            return 999
        return sorted(cols, key=get_col_sort_key)

    sorted_question_cols = sort_multiple_cols(question_cols)
    results = []

    for col in sorted_question_cols:
        # 空でない回答のみを対象
        valid_responses = df[df[col].notna() & (df[col] != '')]

        if len(valid_responses) == 0:
            continue

        # 各選択肢ごとの集計
        for idx, row in valid_responses.iterrows():
            results.append({
                '年代': row['age_range'],
                '回答内容': row[col],
                '選択肢': col
            })

    if not results:
        return pd.DataFrame()

    result_df = pd.DataFrame(results)

    # クロス集計（年代×選択肢×回答内容）
    crosstab = pd.crosstab(
        result_df['年代'],
        [result_df['選択肢'], result_df['回答内容']],
        margins=True
    )

    # === カラムを数値順に並び替え ===
    if isinstance(crosstab.columns, pd.MultiIndex):
        def sort_col_key(col_tuple):
            col_name = col_tuple[0]  # 選択肢名
            parts = col_name.split('_')
            if len(parts) >= 3:
                if parts[2].isdigit():
                    return int(parts[2])  # q_3_1 -> 1
                else:
                    return 0  # 長い質問文は0番
            return 999

        # カラムを完全に再ソート
        sorted_cols = sorted(crosstab.columns.tolist(), key=sort_col_key)
        crosstab = crosstab.reindex(columns=sorted_cols)
        print(f"  年代×回答: カラムをソートしました")

    return crosstab

def convert_to_proper_types(value):
    """値を適切な型に変換する関数"""
    if pd.isna(value) or value == '':
        return ''

    # まず文字列に変換
    str_value = str(value)

    # 数値の場合は数値として返す
    try:
        # 整数の場合
        if str_value.isdigit() or (str_value.startswith('-') and str_value[1:].isdigit()):
            return int(str_value)
        # 小数の場合
        float_value = float(str_value)
        # 整数に変換できる場合は整数として返す
        if float_value.is_integer():
            return int(float_value)
        else:
            return float_value
    except ValueError:
        # 数値変換できない場合は文字列として返す
        return str_value

def create_summary_sheet(workbook, results, sheet_name_prefix="クロス集計結果"):
    """結果をスプレッドシートに書き込む"""
    try:
        # 既存のシートがあれば削除
        try:
            existing_sheet = workbook.worksheet(sheet_name_prefix)
            workbook.del_worksheet(existing_sheet)
            print(f"既存の '{sheet_name_prefix}' シートを削除しました")
        except gspread.WorksheetNotFound:
            print(f"'{sheet_name_prefix}' シートは存在しないため、新規作成します")

        # 新しいシートを作成
        worksheet = workbook.add_worksheet(title=sheet_name_prefix, rows=1000, cols=20)
        print(f"新しいシート '{sheet_name_prefix}' を作成しました")

        current_row = 1
        all_data = []  # 一括更新用のデータ配列

        for question, crosstab in results.items():
            if crosstab.empty:
                print(f"  {question}: データが空のためスキップ")
                continue

            print(f"  {question}: データを書き込み中...")

            # 質問タイトルを追加
            all_data.append([f'【{question}】'])
            all_data.append([''])  # 空行

            # ヘッダー行を準備
            if isinstance(crosstab.columns, pd.MultiIndex):
                # 複数レベルのカラムの場合
                header_rows = []
                for level in range(crosstab.columns.nlevels):
                    header_row = [''] + [str(col[level]) if isinstance(col, tuple) else str(col)
                                        for col in crosstab.columns]
                    header_rows.append(header_row)
                all_data.extend(header_rows)
            else:
                # 単一レベルのカラムの場合
                headers = [''] + [str(col) for col in crosstab.columns]
                all_data.append(headers)

            # データ行を追加（型変換を適用）
            for idx, row in crosstab.iterrows():
                if isinstance(idx, tuple):
                    idx_str = ' / '.join(str(i) for i in idx)
                else:
                    idx_str = str(idx)

                # 各値を適切な型に変換
                row_data = [idx_str] + [convert_to_proper_types(val) for val in row.values]
                all_data.append(row_data)

            # 空行を追加
            all_data.append([''])
            all_data.append([''])

        # データを一括で書き込み
        if all_data:
            # 最大列数を計算
            max_cols = max(len(row) for row in all_data)
            # 各行を同じ長さに調整
            for row in all_data:
                while len(row) < max_cols:
                    row.append('')

            # 値の型変換処理を追加
            print("  データの型変換を実行中...")
            processed_data = []
            for row in all_data:
                processed_row = []
                for cell in row:
                    processed_row.append(convert_to_proper_types(cell))
                processed_data.append(processed_row)

            # A1から一括更新
            range_name = f'A1:{chr(65 + max_cols - 1)}{len(processed_data)}'

            # gspreadのupdate関数で値の型を指定
            worksheet.update(
                values=processed_data,
                range_name=range_name,
                value_input_option='USER_ENTERED'  # ユーザー入力として扱う（数値は数値として認識）
            )
            print(f"  データを範囲 {range_name} に一括書き込みしました")

        print(f"結果を '{sheet_name_prefix}' シートに保存完了しました")

    except Exception as e:
        print(f"シート作成中にエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

def run_survey_crosstab(spreadsheet_url, sheet_name):
    """メイン処理：アンケートクロス集計を実行"""
    print("データ読み込み中...")
    df, df_analysis, workbook = load_survey_data(spreadsheet_url, sheet_name)

    print("質問カラムを識別中...")
    single_answer, multiple_answer = identify_question_columns(df_analysis)

    print(f"単一回答質問: {single_answer}")
    print(f"複数回答質問: {list(multiple_answer.keys())}")

    results = {}

    # 単一回答の処理
    print("\n単一回答のクロス集計処理中...")
    for question in single_answer:
        print(f"  処理中: {question}")

        # 質問タイトルを取得（元のdfから2行目を取得）
        question_title = get_question_title(df, question)

        # 性別×回答
        gender_crosstab = process_gender_crosstab(df_analysis, question)
        if gender_crosstab is not None and not gender_crosstab.empty:
            results[f"{question_title} (性別×回答)"] = gender_crosstab

        # 年代×回答
        age_crosstab = process_age_crosstab(df_analysis, question)
        if age_crosstab is not None and not age_crosstab.empty:
            results[f"{question_title} (年代×回答)"] = age_crosstab

        # 性別×年代×回答
        crosstab = process_single_answer_crosstab(df_analysis, question)
        if crosstab is not None and not crosstab.empty:
            results[f"{question_title} (性別×年代×回答)"] = crosstab

    # 複数回答の処理
    print("\n複数回答のクロス集計処理中...")
    for question_group, question_cols in multiple_answer.items():
        print(f"  処理中: {question_group}")

        # 質問タイトルを取得（長い質問文を優先的に選択）
        question_title = question_group  # デフォルト値
        if question_cols:
            # 長い質問文（数字でないもの）を探す
            long_question_col = None
            for col in question_cols:
                parts = col.split('_')
                if len(parts) >= 3 and not parts[2].isdigit():
                    long_question_col = col
                    break

            # 長い質問文があればそれを使用、なければ最初の選択肢を使用
            title_col = long_question_col if long_question_col else question_cols[0]
            question_title = get_question_title(df, title_col)

        # 性別×回答
        gender_crosstab = process_multiple_answer_gender_crosstab(df_analysis, question_group, question_cols)
        if gender_crosstab is not None and not gender_crosstab.empty:
            results[f"{question_title} (性別×回答)"] = gender_crosstab

        # 年代×回答
        age_crosstab = process_multiple_answer_age_crosstab(df_analysis, question_group, question_cols)
        if age_crosstab is not None and not age_crosstab.empty:
            results[f"{question_title} (年代×回答)"] = age_crosstab

        # 性別×年代×回答
        crosstab = process_multiple_answer_crosstab(df_analysis, question_group, question_cols)
        if crosstab is not None and not crosstab.empty:
            results[f"{question_title} (性別×年代×回答)"] = crosstab

    # 結果をスプレッドシートに保存
    print("\n結果をスプレッドシートに保存中...")

    # 結果を質問番号順にソート
    def get_question_sort_key(title):
        """タイトルから質問番号とサブ番号を抽出してソート用のキーを返す"""
        try:
            # "q_0_..." から質問番号を抽出
            if title.startswith('q_'):
                # アンダースコアで分割してカラム名部分を抽出
                title_parts = title.split(' (')[0]  # " (性別×回答)" の部分を除去
                col_parts = title_parts.split('_')

                if len(col_parts) >= 2 and col_parts[1].isdigit():
                    main_num = int(col_parts[1])

                    # サブ番号の判定
                    if len(col_parts) >= 3:
                        if col_parts[2].isdigit():
                            # q_3_1, q_3_2 などの数字サブ番号
                            sub_num = int(col_parts[2])
                        else:
                            # q_3_長い質問文 の場合は0番として扱う
                            sub_num = 0
                    else:
                        sub_num = 0

                    return (main_num, sub_num)
        except Exception as e:
            print(f"ソートキー抽出エラー: {title}, {e}")
        return (999, 999)  # 抽出できない場合は最後に配置

    # 質問番号順にソートした結果を作成
    print("\n=== ソート前のタイトル一覧 ===")
    for title in results.keys():
        sort_key = get_question_sort_key(title)
        print(f"タイトル: {title}")
        print(f"ソートキー: {sort_key}")
        print("---")

    sorted_results = dict(sorted(results.items(), key=lambda x: get_question_sort_key(x[0])))

    print("\n=== ソート後のタイトル順序 ===")
    for i, title in enumerate(sorted_results.keys()):
        print(f"{i+1}. {title}")

    create_summary_sheet(workbook, sorted_results)

    print("処理完了！")
    return results

# ===========================================
# 実行（変更不要）
# ===========================================
results = run_survey_crosstab(SPREADSHEET_URL, SHEET_NAME)

# 結果の確認（オプション）
for question, crosstab in results.items():
    print(f"\n=== {question} ===")
    print(crosstab)
