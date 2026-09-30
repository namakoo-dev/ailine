import openpyxl
from datetime import date
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "10月"
ws.append(["日付", "借方勘定科目", "借方金額", "貸方勘定科目", "貸方金額", "摘要"])
for row in [
    (date(2026, 10, 1), None, 1320, "現金", 1320, "JR 客先訪問"),
    (date(2026, 10, 2), None, 4200, "現金", 4200, "コピー用紙 アスクル"),
    (date(2026, 10, 3), None, 8800, "普通預金", 8800, "NTT 電話代"),
    (date(2026, 10, 6), None, 3300, "現金", 3300, "Amazon USBケーブル"),
    (date(2026, 10, 9), None, 440, "普通預金", 440, "振込手数料"),
    (date(2026, 10, 14), None, 25000, "現金", 25000, "会食 さくら商店様"),
]:
    ws.append(row)
for r in range(2, 8):
    ws.cell(r, 1).number_format = "yyyy/mm/dd"
wb.save("仕訳_10月.xlsx")
print("ok")
