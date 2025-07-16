import pandas as pd
from openpyxl import Workbook
import win32com.client                                          #Win32comモジュールを呼び出す
import time

# Usage:
# python Excel_Diff.py
#
# ファイルパス（自分のファイルに書き換えてください）
#file1 = 'C:\\Users\\kenic\\Dropbox\\gitdir\\VBA運転集計用\\要調査Jupyterで出力したのと自分のPythonで出力したので異なる\\計画時間my.xlsx'
#file2 = 'C:\\Users\\kenic\\Dropbox\\gitdir\\VBA運転集計用\\要調査Jupyterで出力したのと自分のPythonで出力したので異なる\\計画時間byJupyter.xlsx'
#file2 = 'C:\\Users\\kenic\\OneDrive\\Desktop\\集計のBK\\25-6作成前\\SACLA運転状況集計まとめ.xlsm'
#file2 = 'C:\\Users\\kenic\\OneDrive\\Desktop\\集計のBK\\25-3町田さん作成\\SACLA運転状況集計まとめ.xlsm'
file1 = 'C:\\Users\\kenic\\OneDrive\\Desktop\\集計のBK\\25-6完\\SACLA運転状況集計まとめ.xlsm'
file2 = '\\\\saclaopr18.spring8.or.jp\\common\\運転状況集計\\最新\\SACLA\\SACLA運転状況集計まとめ.xlsm'
sheet_name = 'Fault集計'
#sheet_name = 'まとめ '

# Excel読み込み
df1 = pd.read_excel(file1, sheet_name=sheet_name, header=None)
df2 = pd.read_excel(file2, sheet_name=sheet_name, header=None)

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

# 差分をExcelに出力
if diffs:
    diff_df = pd.DataFrame(diffs, columns=["シート名", "セル", "ファイル1の値", "ファイル2の値"])
    diff_df.to_excel("差分リスト.xlsx", index=False)
    print("差分が見つかってしまいました。\n差分リスト.xlsx を出力しました！")
    
    try:
        excelapp = win32com.client.Dispatch('Excel.Application')        #Excelアプリケーションを起動する
        excelapp.Visible = 1                                            #Excelウインドウを表示する
        wb = excelapp.Workbooks.Open("C:\\Users\\kenic\\Dropbox\\gitdir\\Excel_Diff\\差分リスト.xlsx",ReadOnly=False)
        input("Excelが開かれました。\nEnterを押すと終了します。")
        wb.Close(SaveChanges=False)
        excelapp.Quit()
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")
    finally:
        pass
        excelapp.Application.Quit()
    
else:
    print("差分は見つかりませんでした。")
