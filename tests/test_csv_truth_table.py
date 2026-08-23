# CSV 型検疫の真理値表（手で凍結・実装より先）── 2026-08-22
#
# ここに書いた期待値は「実装がこう返した」ではなく「俺がこう決めた」。
# 実装がこの表に合わないなら、直すのは実装であってこの表ではない。
# 表を変えるのは設計判断の変更であり、DESIGN-20260821-multifile.md の
# CSV 検疫 設計 v2 と同時に変更し、理由を残すこと。
#
# ## 分類の定義（凍結）
#
# kind は 4 分類・列単位・多数決禁止:
#   "number"      全非空セルが ^-?\d{1,15}(\.\d+)?$ に合致（カンマ列は列一貫時のみ+開示）
#   "date"        全非空セルが yyyy/m/d または yyyy-mm-dd の形で暦として成立
#                 （2 形式の混在は許す ── どちらも曖昧さがないため）
#   "string"      上記に確信をもって入らないものをバイト忠実に保持。
#                 一票拒否権 a〜g もここに落ちる。開示（reasons）は出すが ⚠ ではない。
#                 「壊さず持つ」は正しい判断であり、疑いではない。
#   "undecidable" データ自身が矛盾・破損を示唆する形。⚠ を出し、その列に ✓ を出さない。
#                 例: 暦として不成立な日付様式・カンマ桁区切りの不整合・
#                     確信クラス同士の衝突（日付と数値が同居）
#
# ## 一票拒否権（1 セルでも該当すれば列全体が number/date になれない）
#   a leading_zero        ^-?0\d （"0"・"0.5" は該当しない）
#   b (桁固定記号は正規表現不合致で自然に string へ ── 専用 veto は置かない)
#   c formula_head        先頭 = または @ は無条件で文字列強制（数式化け実測 'f' の遮断）。
#                         先頭 + は数値規則に合致しないので自然に string。
#                         先頭 - は有効な数値なら number（-1 を疑わない）
#   d digit_overflow      整数部 16 桁以上（Excel の 15 桁精度で壊れる）
#   e surrounding_space   前後空白（trim してから判定しない ── バイトを変えない）
#   f fullwidth_digit     全角数字を含む
#   g accounting_negative △ 先頭または (数値) 形（会計負数の可能性 ── 解釈せず保持）
#
# ## セル単位の隔離（列 kind とは独立の名指し）
#   control_char   C0 制御文字（openpyxl write が落ちる実測）→ セルを名指し
#   overlong       32,768 文字以上（Excel 上限 32,767）→ セルを名指し
#
# ## 空セルの扱い
#   空文字セルは投票しない。全空の列は string + empty_column。

import pytest

csv_quarantine = None
try:
    from ailine_core import csv_quarantine  # noqa: F401
except ImportError:
    pass

needs_impl = pytest.mark.xfail(
    csv_quarantine is None,
    reason="csv_quarantine 未実装（真理値表は凍結済み・実装が来たら自動で実測に切り替わる）",
    strict=True,
)

# (見出し, セル列, 期待 kind, reasons に必ず含む語)
TRUTH = [
    # --- 0 落ち（商品価値の中核: 先頭ゼロを守る）---
    ("先頭ゼロ 0123", ["0123"], "string", ["leading_zero"]),
    ("先頭ゼロ 00456", ["00456"], "string", ["leading_zero"]),
    ("先頭ゼロ 007", ["007"], "string", ["leading_zero"]),
    ("先頭ゼロ 0012345", ["0012345"], "string", ["leading_zero"]),
    ("負号つき先頭ゼロ -0123", ["-0123"], "string", ["leading_zero"]),
    ("郵便7桁（ゼロ無しのみ）は規則どおり数値", ["1000001"], "number", []),
    ("郵便列にゼロ始まり混在で列ごと文字列", ["1000001", "0600000"], "string", ["leading_zero"]),
    ("電話番号はハイフンで自然に文字列", ["03-1234-5678", "090-0000-0001"], "string", []),
    # --- 境界: ゼロそのものと小数 ---
    ("単独の 0 は数値", ["0"], "number", []),
    ("0.5 は数値（小数点前の 0 は veto でない）", ["0.5"], "number", []),
    ("-1 は数値（負号を疑わない）", ["-1"], "number", []),
    # --- 桁溢れ ---
    ("15 桁ちょうどは数値", ["123456789012345"], "number", []),
    ("16 桁 JAN は文字列", ["4901234567890123"], "string", ["digit_overflow"]),
    ("18 桁は文字列", ["123456789012345678"], "string", ["digit_overflow"]),
    # --- 数式注入（実測: openpyxl が data_type 'f' にする）---
    ("=SUM は文字列強制", ["=SUM(1,2)"], "string", ["formula_head"]),
    ("=1+1 は文字列強制", ["=1+1"], "string", ["formula_head"]),
    ("@SUM は文字列強制", ["@SUM(A1)"], "string", ["formula_head"]),
    ("=cmd 注入形も文字列強制", ["=cmd|'/c calc'!A1"], "string", ["formula_head"]),
    ("+1 は数値規則不合致で文字列", ["+1"], "string", []),
    # --- 前後空白（trim しない）---
    ("空白で包まれた 0123", [" 0123 "], "string", ["surrounding_space"]),
    ("末尾空白の 123", ["123 "], "string", ["surrounding_space"]),
    # --- 全角 ---
    ("全角数字", ["１２３"], "string", ["fullwidth_digit"]),
    ("全角数字とカンマ", ["１，２３４"], "string", ["fullwidth_digit"]),
    ("長音符は自然に文字列", ["ー123"], "string", []),
    # --- カンマ桁区切り（列一貫時のみ+開示）---
    ("一貫したカンマ列は数値+開示", ["1,234", "2,345"], "number", ["comma_grouped"]),
    ("1000 未満混在でも桁区切りとして一貫", ["1,234", "567"], "number", ["comma_grouped"]),
    ("小数つきカンマも一貫なら数値", ["1,234.56", "2,000.00"], "number", ["comma_grouped"]),
    ("カンマ形式と裸の 4 桁の混在は判定不能", ["1,234", "1234"], "undecidable", ["comma_inconsistent"]),
    ("桁位置の壊れたカンマは判定不能（欧州小数点の疑い）", ["12,34"], "undecidable", ["comma_inconsistent"]),
    # --- 会計負数（解釈せず保持）---
    ("△負数は文字列保持", ["△1,234"], "string", ["accounting_negative"]),
    ("括弧負数は文字列保持", ["(1,234)"], "string", ["accounting_negative"]),
    # --- 日付 ---
    ("yyyy/m/d は日付", ["2026/1/2", "2026/12/31"], "date", []),
    ("yyyy-mm-dd は日付", ["2026-01-02"], "date", []),
    ("2 形式の混在は許す（どちらも曖昧さがない）", ["2026/1/2", "2026-01-02"], "date", []),
    ("閏日の成立年は日付", ["2028/2/29"], "date", []),
    ("閏日の不成立年は判定不能", ["2026/2/29"], "undecidable", ["calendar_invalid"]),
    ("13 月 45 日は判定不能", ["2026/13/45"], "undecidable", ["calendar_invalid"]),
    ("8 桁数字は数値のまま+日付の可能性を開示", ["20260102"], "number", ["eight_digit_maybe_date"]),
    ("和暦 R8.1.2 は対象外宣言つき文字列", ["R8.1.2"], "string", ["wareki_out_of_scope"]),
    ("和暦 令和8年1月2日 も同様", ["令和8年1月2日"], "string", ["wareki_out_of_scope"]),
    ("年なし 1/2 は解釈しない文字列", ["1/2"], "string", []),
    ("英語月 Jan-26 は文字列", ["Jan-26"], "string", []),
    # --- 確信クラスの衝突 ---
    ("日付と数値の同居は判定不能", ["2026/1/2", "123"], "undecidable", ["mixed_confident"]),
    # --- 欠損・記号 ---
    ("TRUE は文字列（第一波に真偽クラスなし）", ["TRUE"], "string", []),
    ("プレースホルダ - 混在の数値列は文字列保持（Σ 対象外の開示で守る）",
     ["123", "-", "456"], "string", []),
    ("#N/A などエラー字句は文字列", ["#N/A", "#DIV/0!"], "string", ["excel_error_token"]),
    ("NULL は文字列", ["NULL"], "string", []),
    ("空セルは投票しない", ["123", "", "456"], "number", []),
    ("全空列は文字列+開示", ["", ""], "string", ["empty_column"]),
    # --- 科学記法ふう（Excel なら 1E5→100000 に壊す領域）---
    ("E5 は文字列", ["E5"], "string", []),
    ("1E5 は文字列（指数解釈しない）", ["1E5"], "string", []),
    ("NaN は文字列", ["NaN"], "string", []),
    ("Infinity は文字列", ["Infinity"], "string", []),
]


@needs_impl
@pytest.mark.parametrize(
    "cells,kind,reasons", [t[1:] for t in TRUTH], ids=[t[0] for t in TRUTH]
)
def test_type_truth_table(cells, kind, reasons):
    v = csv_quarantine.classify_column(cells)
    assert v.kind == kind
    for r in reasons:
        assert r in v.reasons, f"開示 {r!r} が無い: {v.reasons}"


@needs_impl
def test_undecidable_is_the_only_warn_kind():
    # ⚠ は kind=="undecidable" の列だけ。文字列保持は開示であって疑いではない。
    warn = csv_quarantine.classify_column(["2026/13/45"])
    hold = csv_quarantine.classify_column(["0123"])
    assert warn.warn is True
    assert hold.warn is False


# --- セル単位の隔離（列 kind と独立）---

@needs_impl
def test_control_char_cells_are_named():
    flags = csv_quarantine.control_char_cells(["ok", "a\x01b", "fine"])
    assert [(i, cp) for i, cp in flags] == [(1, "U+0001")]


@needs_impl
def test_overlong_cells_are_named():
    flags = csv_quarantine.overlong_cells(["x" * 32768, "short"])
    assert [i for i, _ in flags] == [0]


# --- 文字コード判定（BOM 最優先 → UTF-8 → cp932・実測根拠: 短文 UTF-8 の
#     cp932 黙読 41%・逆方向 0%）---

ENC_TRUTH = [
    ("BOM は最優先", "ヘッダ,列".encode("utf-8-sig"), "utf-8-sig", False),
    ("純 ASCII は UTF-8 扱い", b"a,b,c\r\n1,2,3", "utf-8", False),
    # 実測 2026-08-22: この UTF-8 バイト列は cp932 で復号不能 → 曖昧でない
    ("UTF-8 日本語は UTF-8", "名前,金額\n田中,100".encode("utf-8"), "utf-8", False),
    # 実測 2026-08-22: この cp932 バイト列は UTF-8 で復号不能
    ("cp932 専用バイトは cp932", "名前,金額\n田中,100".encode("cp932"), "cp932", False),
    # 実測 2026-08-22: 「ち」の UTF-8 バイト e381a1 は cp932 でも「縺｡」に化けて
    # 復号できてしまう（黙読帯の実物）── UTF-8 続行 + 開示（M2 変異の主戦場）
    ("両方可の短文は UTF-8 続行+開示", "ち".encode("utf-8"), "utf-8", True),
]


@needs_impl
@pytest.mark.parametrize(
    "raw,enc,ambiguous", [t[1:] for t in ENC_TRUTH], ids=[t[0] for t in ENC_TRUTH]
)
def test_encoding_truth_table(raw, enc, ambiguous):
    d = csv_quarantine.detect_encoding(raw)
    assert d.encoding == enc
    assert d.ambiguous is ambiguous


# レビュー実測で発見した穴の追補 2026-08-23（rule d: 桁溢れの一票拒否権がカンマ桁区切りに
# 届いていなかった・SEALED-20260823-jisaku-ultra.md 致命⑤）

@needs_impl
def test_comma_grouped_16_digits_is_vetoed_truth_table():
    """"1,234,567,890,123,456"（整数部16桁）── カンマ形でも digit_overflow で string。"""
    v = csv_quarantine.classify_column(["1,234,567,890,123,456"])
    assert v.kind == "string"
    assert "digit_overflow" in v.reasons


@needs_impl
def test_comma_grouped_15_digits_still_number_truth_table():
    """誤爆防止: 整数部15桁ちょうど（境界の1つ内側）のカンマ列は従来どおり number。"""
    v = csv_quarantine.classify_column(["123,456,789,012,345"])
    assert v.kind == "number"
    assert "comma_grouped" in v.reasons
