"""filetypes — ブック様ファイルの拡張子判定を1箇所に集約した登録簿。
   ★ ailine を import しない（tests/test_line_budget.py の移植可能性番人と同じ線）。
   stdlib のみ。

   利用箇所ごとに「拡張子っぽい」の意味が違う（scan が拾う対象／openpyxl が実際に
   開ける形式／依頼文の2冊目候補らしさ、はそれぞれ別の判定）。無理に1集合へ潰さず
   名前つきの複数集合として置く ── 各利用箇所は元のリテラルと完全に同じ値をここから
   引くだけで、挙動は1ミリも変えない。
"""
from __future__ import annotations

# ailine.py の `_looks_like_second_book_path`（M3 `ailine run <A.xlsx> <B.xlsx> "<依頼>"`
# の2冊目トークン判定）が「表計算らしい拡張子」として認める集合。
# ★ .csv / .tsv も含む（このリファクタでは .csv の扱いを一切変えない ── 現状の集合を
# そのまま移しただけ）。
BOOKLIKE_SUFFIXES = {".xlsx", ".xls", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt",
                     ".ods", ".ots", ".csv", ".tsv"}

# ailine_core/multifile.py の `classify_folder_contents`（`ailine scan <folder>` が
# フォルダ直下で候補として拾う拡張子）。★ ここは「拾う対象」であって「開ける」保証では
# ない ── .xls も候補には入るが、開けるかどうかは後段（OPENPYXL_READABLE_SUFFIX 側）
# で別に判定される。
SCAN_CANDIDATE_SUFFIXES = {".xlsx", ".xls"}

# ailine_core/multifile.py（open_base_workbook・evaluate_file）・extract_multi.py
# （evaluate_and_extract）・stack.py（evaluate_and_stack）が共通で使う、
# openpyxl(data_only=True) で実際に読める唯一の拡張子。基準ブック選定・各ファイルの
# 照合のいずれも「開ける形式か」の判定はこの1値との一致で行う（.xls はここで弾かれ
# 「旧形式(.xls)」として報告される）。
OPENPYXL_READABLE_SUFFIX = ".xlsx"
