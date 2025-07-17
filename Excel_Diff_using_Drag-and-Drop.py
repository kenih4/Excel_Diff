import pandas as pd
import sys
import win32com.client
import os
import time
import datetime
print("len(sys.argv) =  ",len(sys.argv))
print("sys.argv = ", sys.argv)
# コマンドライン引数からファイル名を取得
if len(sys.argv) < 3:
    print("使い方: python compare_excel.py <file1> <file2>")
    input("Enterを押すと終了します...")
    sys.exit(1)
file1 = sys.argv[1]
file2 = sys.argv[2]
sheet_name = 'SACLA'

print(file1)
print(file2)
input("Enterを押すと続行します...")

# Excel読み込み
df1 = pd.read_excel(file1, sheet_name=sheet_name, header=None)
df2 = pd.read_excel(file2, sheet_name=sheet_name, header=None)
input("DEBUG    Enterを押すと続行します...")
# サイズを揃える
max_rows = max(df1.shape[0], df2.shape[0])
max_cols = max(df1.shape[1], df2.shape[1])
df1 = df1.reindex(index=range(max_rows), columns=range(max_cols))
df2 = df2.reindex(index=range(max_rows), columns=range(max_cols))

# 差分をリストにまとめる
diffs = []
for row in range(max_rows):
    for col in range(max_cols):
        val1 = df1.iat[row, col]
        val2 = df2.iat[row, col]
        if pd.isna(val1) and pd.isna(val2):
            continue
        if val1 != val2:
            cell = f"{chr(65 + col)}{row + 1}"  # A1形式
            diffs.append([sheet_name, cell, val1, val2])

input("DEBUG 2    Enterを押すと続行します...")

# 差分をExcelに出力
if diffs:
    output_file = "C:\\Users\\kenic\\Dropbox\\gitdir\\Excel_Diff\\差分リスト.xlsx"
    diff_df = pd.DataFrame(diffs, columns=["シート名", "セル", "ファイル1の値", "ファイル2の値"])
    diff_df.to_excel(output_file, index=False)
    print(diff_df)
    print("差分が見つかってしまいました。\n差分リスト.xlsx を出力しました！")
    input("Enterを押すと開きます。")
    
    try:
        """
        excelapp = win32com.client.Dispatch('Excel.Application')        #Excelアプリケーションを起動する
        excelapp.Visible = 1                                            #Excelウインドウを表示する
        wb = excelapp.Workbooks.Open("差分リスト.xlsx", ReadOnly=False)
        input("Excelが開かれました。\nEnterを押すと終了します。")
        wb.Close(SaveChanges=False)
        excelapp.Quit()
        """
        print(f'Excelファイル "{output_file}" に出力しました。')
        if abs(time.time() - os.path.getmtime(output_file))<10:
#           input("正常に「.xlsx」が作成されました。\nPress Enter to Exit...")
            os.startfile(output_file)
        else:
            print(f"異常：作成されたはずの.xlsxのタイムスタンプが古いです。 最終更新時刻: {datetime.datetime.fromtimestamp(os.path.getmtime(output_file))}")
        
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")
else:
    print("差分は見つかりませんでした。")
