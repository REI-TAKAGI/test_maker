import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os

# 最新の模擬試験ファイルを読み込む
latest_exam_file = max(
    [f for f in os.listdir() if f.startswith('模擬試験_') and f.endswith('.xlsx')],
    key=os.path.getctime
)
df_exam = pd.read_excel(latest_exam_file)

# 問題データを読み込む
df_question = pd.read_excel('問題データ.xlsx')

# 問題文を検索して回答を取得する関数
def get_answer(row):
    question = row.iloc[0]  # 問題はA列
    answer = df_question.loc[df_question['問題'] == question, '正解'].values
    return answer[0] if len(answer) > 0 else None

# 模擬試験と問題データをマージして採点結果を計算する
df_exam['正解'] = df_exam.apply(get_answer, axis=1)
df_exam['採点結果'] = df_exam.apply(lambda row: '〇' if row['回答'] == row['正解'] else '×' if row['回答'] is not None else None, axis=1)

# 採点結果の列に正解がある場合のみ、その行の点数を計算する
score = df_exam[df_exam['採点結果'] == '〇'].shape[0]

# 採点結果を出力するExcelファイル名を設定する
current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f'採点済み_{latest_exam_file}'

# 採点結果を出力するExcelファイルに書き込む
with pd.ExcelWriter(output_filename) as writer:
    df_exam.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet['B75'] = score

print(f'採点結果が {output_filename} として出力されました。')
