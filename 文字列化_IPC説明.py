import sys
import os
import xlwings as xw
import openpyxl

# スクリプト or .exe のあるディレクトリを取得
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS  # pyinstallerで一時展開されたフォルダ
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# IPCファイルのフルパスを組み立て
ipc_path = os.path.join(base_dir, "IPC_Ver2025-ALLsection.xlsx")

# IPC説明辞書の構築（A列=IPCコード, B列=説明）
ipc_dict = {}
ipc_wb = openpyxl.load_workbook(ipc_path, data_only=True)
ipc_ws = ipc_wb.active

for row in ipc_ws.iter_rows(min_row=1):
    code_cell = row[0]  # A列
    desc_cell = row[1]  # B列
    if code_cell.value and desc_cell.value:
        code = str(code_cell.value).strip().replace(" ", "")
        desc = str(desc_cell.value).strip()
        ipc_dict[code] = desc
ipc_wb.close()

# 対象のExcelファイルを開く（この部分は必要に応じてファイル選択に変えてもOK）
app = xw.App(visible=False)
wb = app.books.open("整理済みファイル_値に変換.xlsm")
ws = wb.sheets[0]  # 最初のシート

# 見出し行（3行目）から列位置を特定
headers = ws.range("3:3").value
ipc_col = None
desc_col = None

for col_idx, header in enumerate(headers):
    if header == "筆頭ＩＰＣ":
        ipc_col = col_idx + 1
        desc_col = ipc_col + 1
        break

if ipc_col and desc_col:
    last_row = ws.used_range.last_cell.row
    for row in range(4, last_row + 1):
        ipc_code = str(ws.cells(row, ipc_col).value or "").strip().replace(" ", "")
        matched_desc = ""
        for key in sorted(ipc_dict.keys(), key=len, reverse=True):
            if ipc_code.startswith(key):
                matched_desc = ipc_dict[key]
                break
        ws.cells(row, desc_col).value = matched_desc
else:
    print("『筆頭ＩＰＣ』列が見つかりませんでした。")

# 保存＆終了
wb.save("整理済みファイル_IPC説明追加.xlsm")
wb.close()
app.quit()

print("完了しました：整理済みファイル_IPC説明追加.xlsm に保存されました。")
