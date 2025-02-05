import pandas as pd
import random
from datetime import datetime

# 問題データを読み込む
df = pd.read_excel('問題データ.xlsx', sheet_name='Sheet1')

# 問題文のみを取得し、解答欄のシートを作成する
questions = df['問題']
answers_sheet = pd.DataFrame({'問題': questions})

# ランダムに選んだ70問を取得
selected_questions = answers_sheet.sample(n=70)

# 現在の日付と時刻を取得
current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")

# 模擬試験_日付時刻.xlsxに書き込む
output_filename = f'模擬試験_{current_datetime}.xlsx'
selected_questions.to_excel(output_filename, index=False)

print(f'模擬試験が {output_filename} として出力されました。')

# 採点処理は別のスクリプトで行う
# 採点スクリプトは作成されたエクセルファイルを読み込み、回答と正解を比較して採点し、結果を出力する

