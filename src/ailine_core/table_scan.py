"""表の走査 ── シートのどこまでが使われているか、見出しから列を引く、座標を書く。

★ なぜ在るか（2026-09-03・分割の一歩目）: 単一ファイル（17,000 行超）を割る前に、
**事後条件の層が要る荷物**を先に外へ出す。この 7 本は「op ごとの事後条件」から
14〜2 箇所ずつ呼ばれ、かつ**その外からも呼ばれる**ので、事後条件層と一緒に動かすと
`ailine_core → ailine` の循環になる。だから先にここへ置く。

★ 何を置くかの線: **シートを走査して位置や広がりを求めるもの**だけ。
値の意味づけ（合計行か・数値か・空欄か）は置かない ── それは
`total_row` / `primitives` / `column_type` の仕事で、混ぜると
「表のどこを見ているか」と「その値をどう解釈するか」が 1 つのモジュールに同居する。

★ 走査と物理の使用範囲は**別物**（`extent_gap` が在る理由）:
  走査は「見出しから連続して値がある範囲」、物理は「openpyxl が持っている範囲」。
  食い違うのは、離れた場所に値が残っている時 ── その差を数えて開示するために
  両方を持っている。**片方だけ見て『全部見た』と言わない**ための組。
"""
from __future__ import annotations

from openpyxl.utils import get_column_letter

def _cell_ref(row: int, col: int) -> str:
    """**1 起点**の行/列 (r, c) を、人が読める A1 形式（例: C2）にする。

    ★★ この repo には**起点が 2 つある**（2026-09-03 に数えて確かめた）:
      ・**0 起点** … LibreOffice Basic 側。`getCellByPosition(列, 行)`、
        ヘルパの `headerRow` 引数、LLM に見せる「列は 0 起点で 0..N」の説明、
        `used_range`（ailine.py:643 に 0起点→1起点 の変換が在る）
      ・**1 起点** … openpyxl 側と検算側。`_col_index_by_header` の戻り値、
        スナップショットの `r,c`（`for r in range(1, nrow + 1)`）

    ★ この関数は **1 起点の側**にある。docstring は長く「0起点」と書かれていたが、
      呼び元 8 箇所すべてが 1 起点を渡しており、**実装と呼び元は整合していた** ──
      説明だけが Basic 側の世界のまま取り残されていた。
    ★ 0 を渡すと `get_column_letter` が ValueError を投げる（列 0 は存在しない）。
      それでよい ── **黙って `B0` のような嘘の座標を作るより落ちる方が安全**。
      この振る舞いは tests/test_cell_ref_is_one_based.py が凍結している。
    """
    return f"{get_column_letter(col)}{row}"

def _col_index_by_header(ws, name: str, header_row: int = 1):
    """見出し行(既定は物理1行目・header_row で指定可)を左から走査して name に一致する
       列の1起点インデックスを返す。無ければ None。★ W3: header_row 省略時は旧挙動と同一。
       ★ W3: header_row>1 の子見出し行で空欄の列は、book_columns() と同じ規則で
       真上の行を遡って引き継ぐ（多段見出しの先頭列対策。無いと『7月』等キー列より
       手前の空欄列で走査が誤って打ち切られる）。"""
    c = 1
    while True:
        v = ws.cell(row=header_row, column=c).value
        if v in (None, "") and header_row > 1:
            for up in range(header_row - 1, 0, -1):
                uv = ws.cell(row=up, column=c).value
                if uv not in (None, ""):
                    v = uv
                    break
        if v in (None, ""):
            return None
        if str(v) == name:
            return c
        c += 1

def _scan_last_row(ws, key_col: int = 1, header_row: int = 1) -> int:
    """key_col(1起点)を上から走査した最終データ行（見出し行を除く）。データが無ければ
       header_row。★ W3: header_row 省略時は旧挙動（見出し=1行目・データ開始=2行目）と同一。"""
    r = header_row + 1
    while ws.cell(row=r, column=key_col).value not in (None, ""):
        r += 1
    return r - 1

def _used_extent(ws) -> tuple:
    """シートの**使用範囲**（行数, 列数）。中身のあるセルが 1 つも無ければ (0, 0)。

    ★ なぜ在るか（盲検の契約レビュー・2026-08-24）: 「雛形は 1 セルも変わっていない」を
      走査（A 列・1 行目を上/左から見て最初の空で打ち切り）で測っていたため、
      **A 列を余白にした典型的な請求書雛形**（B2 に「請 求 書」）では比較セル数が 0 になり、
      雛形が何セル壊れても pass していた。**分母ゼロの主張は主張ではない。**
    ★ 雛形は「表」ではなく「紙の見た目」なので、走査でなく使用範囲で測るのが正しい。
      openpyxl の max_row/max_column は書式だけのセルも含んで大きく出ることがあるので、
      値のあるセルの最大位置を自分で取る。
    """
    last_r = last_c = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                if cell.row > last_r:
                    last_r = cell.row
                if cell.column > last_c:
                    last_c = cell.column
    return last_r, last_c

def _scan_last_col(ws, header_row: int = 1) -> int:
    """見出し行(既定は物理1行目)を左から走査した最終列（1起点）。
       ★ W3: _col_index_by_header と同じ規則で、子見出し行(header_row>1)の空欄列は
       真上の行を遡って引き継ぐ（多段見出しの先頭列対策）。"""
    def _effective(c: int):
        v = ws.cell(row=header_row, column=c).value
        if v in (None, "") and header_row > 1:
            for up in range(header_row - 1, 0, -1):
                uv = ws.cell(row=up, column=c).value
                if uv not in (None, ""):
                    return uv
        return v
    c = 1
    while _effective(c) not in (None, ""):
        c += 1
    return c - 1

def extent_gap(ws, header_row: int = 1, key_col: int = 1) -> dict:
    """走査で得た範囲と、**物理の使用範囲**の食い違いを数える。

    ★ 2026-08-25（塊②）: 中核 op の盲検が実測した 2 件の根。
      行: `_scan_last_row` は key_col を上から見て**最初の空で止まる**。末尾に
          その列が空の行があると、処理からも分母からも消える（「3行中1行が一致」の
          真の分母は 5 だった）。★ 表の**途中**の空きは ⚠ が出るのに、
          **末尾だけ鳴らない** ── 警告条件が「下方向に中身があるか」なので原理的に発火しない。
      列: 並べ替えの範囲は見出し由来（`len(headers)-1`）なので、見出しの無い列は
          範囲外に落ちる。実測では D 列だけ動かず、**全行で備考が別商品の物に付け替わった**。

    ★ 物理の使用範囲は「値のあるセルの最大位置」で測る（_used_extent と同じ規則。
      openpyxl の max_row/max_column は書式だけのセルを含んで大きく出るため使わない）。
    戻り値の rows_* / cols_* はいずれも**見出し行を除いたデータの数**。
    """
    phys_r, phys_c = _used_extent(ws)
    rows_physical = max(0, phys_r - header_row)
    cols_physical = max(0, phys_c)
    rows_scanned = max(0, _scan_last_row(ws, key_col=key_col, header_row=header_row) - header_row)
    cols_scanned = max(0, _scan_last_col(ws, header_row=header_row))
    return {
        "rows_scanned": rows_scanned, "rows_physical": rows_physical,
        "rows_missing": max(0, rows_physical - rows_scanned),
        "cols_scanned": cols_scanned, "cols_physical": cols_physical,
        "cols_missing": max(0, cols_physical - cols_scanned),
    }

def data_extent(ws, header_row: int = 1) -> tuple:
    """見出しより下のデータの**物理の**広がり (最終行, 最終列)。

    ★★ 2026-08-27（Namakoo が実測・今日 3 箇所目）: `_scan_last_row` は 1 列目を上から
      見て**最初の空で止まる**。表の途中に空行が 1 本あるだけで、その下が全部消える。
      ・位置の解決が「みかんが見つかりません」になった
      ・事後条件が「検証対象が 0 件」になった
      ・分母が消えた（塊①で直したのと同じ形）
      ★ 3 箇所で同じ穴を開けたので、**器官を 1 つにする**（それぞれに書き写さない）。
    """
    phys_r, phys_c = _used_extent(ws)
    last = max(phys_r, _scan_last_row(ws, header_row=header_row))
    cols = max(phys_c, _scan_last_col(ws, header_row=header_row), 1)
    return last, cols
