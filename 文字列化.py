import xlwings as xw

# 対象の見出し（部分一致のものを含む）
exact_match_headers = [
    "出願年", "特許", "特開", "JP", "WO", "CN", "US", "EP",
    "KR", "TW", "DE", "IN", "AU", "FR", "BR",
    "筆頭ＩＰＣ", "筆頭FIメイングループ"
]
partial_match_keyword = "抄録リンク"

# ブックを開く（マクロ有効ファイルもOK）
app = xw.App(visible=False)
wb = app.books.open("整理済みファイル.xlsm")
ws = wb.sheets[0]

# 見出し行を取得（3行目）
headers = ws.range("3:3").value

for col_idx, header in enumerate(headers):
    if header is None:
        continue
    if header in exact_match_headers or partial_match_keyword in str(header):
        col_letter = xw.utils.col_name(col_idx + 1)
        # 4行目以降（最大行まで）の値取得
        cell_range = f"{col_letter}4:{col_letter}{ws.used_range.last_cell.row}"
        values = ws.range(cell_range).value

        # 見た目の値で上書き
        if isinstance(values, list):
            for row_offset, val in enumerate(values):
                ws.range(f"{col_letter}{row_offset+4}").value = val
        else:
            ws.range(f"{col_letter}4").value = values

# 保存＆終了
wb.save("整理済みファイル_値に変換.xlsm")
wb.close()
app.quit()

