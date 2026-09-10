"""form_grid — 帳票を「座標 → 中身」の格子として読む層。**意味を持たない**。

★★ なぜ要るか（2026-09-11）:
  実物の請求書は「表」ではなく**帳票**だった（実測 28 冊）。取引先は列でなくセルに在り、
  明細は別ブロック、集計は下の帯。だから既存の `table_scan`（表として読む層）とは
  別の入口が要る。

★ ここは**値の意味づけを 1 つも持たない**（`cellmap` / `table_scan` と同じ規律）。
  「どのセルが請求元か」は上の層（規則）が決める。ここが返すのは格子と、その形だけ。

★★ 実物から出た読み取りの癖（この層が吸収するもの・すべて実測）:

    ① 結合セルは**左上だけが値を持つ**。展開せずに読むと、隣を見た瞬間に空になる。
       実物の帳票は結合が 43〜174 箇所（misoca / spreadoffice）。
    ② ★ 展開したまま数えると、**同じ値が結合の幅ぶん重複する**。
       「独立した 2 つの根拠が一致した」を数える時にこれを畳まないと、
       **自分の写しと一致して「裏が取れた」**ことになる（2026-09-10 に実測で踏んだ）。
       → だから `at()` は必ず**アンカー（左上）の番地**を返す。
    ③ 走査窓は 20 行では足りない。実物の明細見出しは 15〜32 行目に在る。
       ★ ただし `STRUCT_HEADER_SCAN_ROWS`（表の層の定数）は動かさない ──
         広げると官公庁の統計表で誤検出が出る（実測 0/6 → 1/6）。
         **窓はこの層の引数として渡す**（定数を触らない）。
    ④ 実物の金額はほぼ式。`data_only=True` はキャッシュを読むので、
       人が計算しない道具で保存した冊は全部 `None` になる。
       ★ この層は「値が無い」と「式が在るのにキャッシュが無い」を**区別して返す**
         （黙って 0 にしない）。

★ ailine を import しない（ailine_core の作法）。
"""
from __future__ import annotations

from dataclasses import dataclass

#: 走査窓の既定（実測: 実物の帳票は明細見出しが 15〜32 行目・帯は 34〜51 行目）。
#: ★ 表の層の定数とは別物。ここを変えても DSL 経路には影響しない。
DEFAULT_ROWS = 60
DEFAULT_COLS = 24


def column_letter(col: int) -> str:
    """1 起点の列番号を A1 記法の列名にする（openpyxl に依存しないで持つ）。"""
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


@dataclass(frozen=True)
class Cell:
    """格子の 1 マス。

    ★ `at` は**アンカーの番地**（結合なら左上）。写しを数えないための鍵。
    """
    row: int
    col: int
    value: object
    at: str                    #: "B13"（★ 結合なら左上の番地）
    anchor: tuple              #: (行, 列) ── 同一性の判定はこれで行う
    from_merge: bool           #: 結合の展開で得た値か（左上そのものなら False）


class Grid:
    """1 シートの格子。★ 意味は持たない。

    使い方:
        g = Grid.read(ws, rows=60, cols=24)
        g.cell(13, 2).value        # B13 の中身（結合なら左上の値）
        g.text_cells()             # 文字が入っているマス
        g.number_cells()           # 数が入っているマス
    """

    def __init__(self, cells: dict, rows: int, cols: int,
                 formula_without_cache: tuple = ()):
        self._cells = cells
        self.rows = rows
        self.cols = cols
        #: ★ 式が在るのに計算結果が保存されていないマスの番地。
        #:   「値が無い」と混ぜない（黙って 0 にしないため）。
        self.formula_without_cache = formula_without_cache

    # ── 読み込み ────────────────────────────────────────────
    @classmethod
    def read(cls, ws, rows: int = DEFAULT_ROWS, cols: int = DEFAULT_COLS,
             ws_formula=None) -> "Grid":
        """worksheet から格子を作る。

        Args:
            ws: `data_only=True` で開いた worksheet（値が見える）。
            rows / cols: 走査窓。★ 定数でなく引数（上の癖③）。
            ws_formula: 同じシートを `data_only=False` で開いたもの（省略可）。
                ★ 渡すと「式は在るがキャッシュが無い」マスを見分けられる（癖④）。
        """
        max_r = min(int(getattr(ws, "max_row", 0) or 0), rows)
        max_c = min(int(getattr(ws, "max_column", 0) or 0), cols)

        # ★ 結合の展開表を先に作る（アンカーの値を範囲全体へ配る）
        anchor_of: dict = {}
        for rng in getattr(ws, "merged_cells", None).ranges if getattr(ws, "merged_cells", None) else []:
            a = (rng.min_row, rng.min_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    anchor_of[(r, c)] = a

        cells: dict = {}
        no_cache = []
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                raw = ws.cell(row=r, column=c).value
                anchor = anchor_of.get((r, c), (r, c))
                from_merge = False
                value = raw
                if raw in (None, "") and anchor != (r, c):
                    value = ws.cell(row=anchor[0], column=anchor[1]).value
                    from_merge = True
                if ws_formula is not None and value in (None, ""):
                    f = ws_formula.cell(row=anchor[0], column=anchor[1]).value
                    if isinstance(f, str) and f.startswith("="):
                        no_cache.append(column_letter(anchor[1]) + str(anchor[0]))
                if value in (None, ""):
                    continue
                cells[(r, c)] = Cell(
                    row=r, col=c, value=value,
                    at=column_letter(anchor[1]) + str(anchor[0]),
                    anchor=anchor, from_merge=from_merge)
        return cls(cells, max_r, max_c, tuple(sorted(set(no_cache))))

    # ── 引き方 ──────────────────────────────────────────────
    def cell(self, row: int, col: int):
        """1 マス（無ければ None）。"""
        return self._cells.get((row, col))

    def all_cells(self) -> list:
        """中身のあるマスを、左上から順に。"""
        return [self._cells[k] for k in sorted(self._cells)]

    def text_cells(self) -> list:
        """文字が入っているマス。"""
        return [c for c in self.all_cells() if isinstance(c.value, str) and c.value.strip()]

    def number_cells(self) -> list:
        """数が入っているマス（★ bool は数として扱わない）。"""
        return [c for c in self.all_cells()
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]

    def anchors(self, cells) -> set:
        """マスの並びを**アンカーの集合**に畳む。

        ★ これが癖②の処方。「独立した根拠が 2 つ」を数える前に必ず通す ──
          通さないと、結合の写しを 2 つ目の根拠として数えてしまう。
        """
        return {c.anchor for c in cells}

    def right_of(self, cell, span: int = 5) -> list:
        """同じ行の右隣から span マス（★ 自分の結合範囲は跨いで返す）。"""
        out = []
        for c in range(cell.col + 1, min(self.cols, cell.col + span) + 1):
            got = self._cells.get((cell.row, c))
            if got is None:
                continue
            if got.anchor == cell.anchor:      # ★ 自分自身の写しは返さない
                continue
            out.append(got)
        return out

    def below_of(self, cell, span: int = 2) -> list:
        """同じ列の下 span マス（★ 自分の結合範囲は跨ぐ）。"""
        out = []
        for r in range(cell.row + 1, min(self.rows, cell.row + span) + 1):
            got = self._cells.get((r, cell.col))
            if got is None or got.anchor == cell.anchor:
                continue
            out.append(got)
        return out
