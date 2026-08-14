"""実務帳票 battery の検体生成スクリプト（測定タスク用・決定論的）。

ailine.py / helpers / refs は一切変更しない。ここで作るのは bench/realworld/*.xlsx の
入力データだけ。再実行可能（既存ファイルを上書きする）。

検体:
  A title_rows.xlsx    見出しが1行目でなく3行目にある帳票（結合タイトル行あり）
  B large.xlsx          3,000行のデータ（MAX_ROWS=1000 の死角を実測するため）
  C formulas.xlsx        数式セル(D列=B*C, 7行目=SUM, F1=TODAY()の揮発性セル)を持つ帳票
  D merged_head.xlsx     2段見出し（結合セルで上期/下期、2行目に月名。7月を含む）
"""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent


def make_title_rows() -> Path:
    """A: 1行目=結合タイトル、2行目=作成日、3行目=見出し、4行目以降=データ10行。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"

    ws["A1"] = "◯◯株式会社 売上台帳"
    ws.merge_cells("A1:E1")
    ws["A1"].font = Font(bold=True, size=14)

    ws["A2"] = "作成日: 2026/08/01"

    headers = ["商品", "金額", "在庫", "売上", "原価"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h)

    products = ["りんご", "みかん", "ぶどう", "もも", "なし",
                "いちご", "メロン", "バナナ", "キウイ", "さくらんぼ"]
    # 決定論的な値（行番号から導出）
    for i, name in enumerate(products):
        row = 4 + i
        kingaku = 1000 + i * 137          # 金額
        zaiko = 50 - i * 3                # 在庫
        uriage = kingaku * (10 - i)       # 売上
        genka = int(kingaku * 0.6)        # 原価
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=kingaku)
        ws.cell(row=row, column=3, value=zaiko)
        ws.cell(row=row, column=4, value=uriage)
        ws.cell(row=row, column=5, value=genka)

    out = HERE / "title_rows.xlsx"
    wb.save(out)
    return out


def make_large() -> Path:
    """B: 1行目見出し（商品コード/数量/単価）、2〜3001行目データ3,000行。決定論的な値。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"

    ws["A1"] = "商品コード"
    ws["B1"] = "数量"
    ws["C1"] = "単価"

    N = 3000
    for i in range(1, N + 1):
        row = 1 + i
        code = f"P{i:05d}"
        suryo = (i % 50) + 1                    # 数量: 1..50 の周期
        tanka = 100 + (i * 37) % 9000            # 単価: 100..9099 の規則的な分布
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=suryo)
        ws.cell(row=row, column=3, value=tanka)

    out = HERE / "large.xlsx"
    wb.save(out)
    return out


def make_formulas() -> Path:
    """C: 1行目見出し（商品/数量/単価/金額/備考）、2〜6行目データ5行。
       D列は式 =B{row}*C{row}、7行目は =SUM(D2:D6)、F1 は =TODAY()（揮発性）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"

    headers = ["商品", "数量", "単価", "金額", "備考"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)

    products = ["ノート", "鉛筆", "消しゴム", "定規", "ハサミ"]
    for i, name in enumerate(products):
        row = 2 + i
        suryo = 10 + i * 5
        tanka = 100 + i * 50
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=suryo)
        ws.cell(row=row, column=3, value=tanka)
        ws.cell(row=row, column=4, value=f"=B{row}*C{row}")
        ws.cell(row=row, column=5, value="")

    # 最下行に SUM 式
    ws.cell(row=7, column=1, value="合計")
    ws.cell(row=7, column=4, value="=SUM(D2:D6)")

    # 揮発性セル（TODAY()）
    ws["F1"] = "=TODAY()"

    out = HERE / "formulas.xlsx"
    wb.save(out)
    return out


def make_merged_head() -> Path:
    """D: 2段見出し。1行目は「上期」(B1:C1結合)「下期」(D1:E1結合)、2行目に月名
       （4月/5月/7月/8月 — 7月を必ず含める）。3行目以降にデータ数行。A列は商品名。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"

    ws["A1"] = "商品"
    ws["B1"] = "上期"
    ws.merge_cells("B1:C1")
    ws["D1"] = "下期"
    ws.merge_cells("D1:E1")

    months = ["4月", "5月", "7月", "8月"]
    for c, m in enumerate(months, start=2):
        ws.cell(row=2, column=c, value=m)
    ws.cell(row=2, column=1, value="")  # A2 は空（A列見出しは商品のみ1行目）

    products = ["りんご", "みかん", "ぶどう", "もも", "なし"]
    for i, name in enumerate(products):
        row = 3 + i
        ws.cell(row=row, column=1, value=name)
        for j, m in enumerate(months):
            col = 2 + j
            val = 100 + i * 17 + j * 23
            ws.cell(row=row, column=col, value=val)

    out = HERE / "merged_head.xlsx"
    wb.save(out)
    return out


def main():
    paths = [make_title_rows(), make_large(), make_formulas(), make_merged_head()]
    for p in paths:
        print(f"生成: {p}")


if __name__ == "__main__":
    main()
