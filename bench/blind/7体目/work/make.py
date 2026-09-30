import openpyxl
from openpyxl.styles import Font, Alignment
from datetime import date

# 1. 請求一覧（タイトル・結合セル・見出し4行目・合計行）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "2026年9月"
ws["A1"] = "請求一覧（2026年9月分）"
ws.merge_cells("A1:I1")
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "作成：2026/9/30　担当：山田"
hdr = ["請求番号", "請求日", "取引先コード", "取引先名", "税抜金額", "消費税", "税込金額", "入金予定日", "入金状況"]
for i, h in enumerate(hdr, 1):
    ws.cell(4, i, h).font = Font(bold=True)
rows = [
    ("INV-0901", date(2026, 9, 1), "0012", "株式会社丸山工業", 150000, "2026/10/31", "入金済"),
    ("INV-0902", date(2026, 9, 3), "0045", "有限会社さくら商店", 48000, "2026/10/31", "未入金"),
    ("INV-0903", date(2026, 9, 5), "0103", "田中歯科クリニック", 220000, "2026/10/31", "入金済"),
    ("INV-0904", date(2026, 9, 10), "0012", "株式会社丸山工業", 33000, "2026/10/31", "未入金"),
    ("INV-0905", date(2026, 9, 12), "0207", "合同会社ブルーリーフ", 99000, "2026/11/30", "未入金"),
    ("INV-0906", date(2026, 9, 15), "0045", "有限会社さくら商店", 12500, "2026/10/31", "一部入金"),
    ("INV-0907", date(2026, 9, 20), "0310", "鈴木会計事務所", 55000, "2026/10/31", "入金済"),
    ("INV-0908", date(2026, 9, 25), "0103", "田中歯科クリニック", 88000, "2026/11/30", "未入金"),
]
r = 5
for no, d, code, name, amt, due, st in rows:
    ws.cell(r, 1, no)
    ws.cell(r, 2, d).number_format = "yyyy/mm/dd"
    c = ws.cell(r, 3, code); c.number_format = "@"
    ws.cell(r, 4, name)
    ws.cell(r, 5, amt).number_format = "#,##0"
    ws.cell(r, 6, f"=ROUNDDOWN(E{r}*0.1,0)").number_format = "#,##0"
    ws.cell(r, 7, f"=E{r}+F{r}").number_format = "#,##0"
    ws.cell(r, 8, due)
    ws.cell(r, 9, st)
    r += 1
ws.cell(r, 4, "合計")
for col in "EFG":
    ws[f"{col}{r}"] = f"=SUM({col}5:{col}{r-1})"
wb.save("請求一覧.xlsx")

# 2. 取引先マスタ
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "マスタ"
ws.append(["取引先コード", "取引先名", "担当者", "締日", "支払サイト"])
for row in [("0012", "株式会社丸山工業", "山田", "月末", "翌月末"),
            ("0045", "有限会社さくら商店", "佐藤", "20日", "翌月末"),
            ("0103", "田中歯科クリニック", "山田", "月末", "翌々月末"),
            ("0207", "合同会社ブルーリーフ", "佐藤", "月末", "翌々月末"),
            ("0310", "鈴木会計事務所", "高橋", "15日", "翌月末")]:
    ws.append(row)
for rr in range(2, 7):
    ws.cell(rr, 1).number_format = "@"
wb.save("取引先マスタ.xlsx")

# 3. 入金明細（通帳から）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "入金"
ws["A1"] = "普通預金 入金明細（9月）"
ws.append([])
ws.append(["入金日", "請求番号", "振込名義", "入金額"])
for row in [(date(2026, 9, 28), "INV-0901", "ｶ)ﾏﾙﾔﾏｺｳｷﾞｮｳ", 165000),
            (date(2026, 9, 29), "INV-0903", "ﾀﾅｶｼｶ", 242000),
            (date(2026, 9, 29), "INV-0906", "ﾕ)ｻｸﾗｼｮｳﾃﾝ", 10000),
            (date(2026, 9, 30), "INV-0907", "ｽｽﾞｷｶｲｹｲ", 60500)]:
    ws.append(row)
wb.save("入金明細.xlsx")

# 4. 仕訳帳
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "仕訳帳"
ws.append(["日付", "借方勘定科目", "借方金額", "貸方勘定科目", "貸方金額", "摘要"])
for row in [
    (date(2026, 9, 1), "旅費交通費", 1280, "現金", 1280, "JR 客先訪問"),
    (date(2026, 9, 2), "消耗品費", 3980, "現金", 3980, "コピー用紙 アスクル"),
    (date(2026, 9, 3), "通信費", 8800, "普通預金", 8800, "NTT 電話代"),
    (date(2026, 9, 5), "会議費", 2400, "現金", 2400, "喫茶 打合せ"),
    (date(2026, 9, 8), "旅費交通費", 560, "現金", 560, "地下鉄"),
    (date(2026, 9, 10), "水道光熱費", 12300, "普通預金", 12300, "東京電力"),
    (date(2026, 9, 12), "消耗品費", 1650, "現金", 1650, "トナー"),
    (date(2026, 9, 15), "接待交際費", 18000, "現金", 18000, "会食 丸山工業様"),
    (date(2026, 9, 20), "旅費交通費", 14200, "普通預金", 14200, "新幹線 大阪出張"),
    (date(2026, 9, 25), "支払手数料", 440, "普通預金", 440, "振込手数料"),
]:
    ws.append(row)
for rr in range(2, 12):
    ws.cell(rr, 1).number_format = "yyyy/mm/dd"
wb.save("仕訳帳.xlsx")
print("ok")
