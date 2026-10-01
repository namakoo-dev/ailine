# 転記先と参照表のシートは依頼文と突き合わせる（2026-10-01・依頼の項の台帳で B だった項目）。
#
# ★ 形: target_sheet / source_sheet は「実在するか」しか見ていなかった。参照表が複数ある冊で、
#   依頼が名指ししないシートを LLM が選ぶと、別の表から引いた値で列を上書きして ✓ が出る
#   （宣言↔実体は合っている ── 依頼の項が欠けていた）。
#
# 契約:
#   ① 依頼文が名指ししていれば黙る（--sheet で選んだ転記先も名指しと数える）
#   ② 名指しが無く、実表で候補が 1 つ（対象列を持つ参照表が 1 枚だけ等）なら機械が決める
#      ── LLM と違えば書き換え、出典を _sources に出す
#   ③ 名指しが無く候補が 2 つ以上なら ⚠（_warnings ＝ ✓→△ の材料）で開示し、他の候補を名指しする
#   ④ 他のシート名の断片（『単価表』と『単価表_旧』）を証拠にしない
#   ⑤ シートが 1 枚の冊・依頼文が空なら黙る

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402

ORDER = ["コード", "数量"]
MASTER = ["コード", "商品名", "単価"]


def _meta(sheets: dict, cli=False):
    m = {"sheets": list(sheets), "headers": {k: list(v) for k, v in sheets.items()},
         "header_rows": {k: 1 for k in sheets}}
    if cli:
        m["_sheet_source"] = "cli"
    return m


def _lookup(task, sheets, target="注文", source="商品表", col="単価", key="コード", cli=False):
    ok, r, _i, err = ailine.verify_dsl_args(
        "LOOKUP_FILL", {"target_sheet": target, "source_sheet": source, "key_col": key,
                        "target_col": col},
        _meta(sheets, cli), task=task, target_sheet=target)
    assert ok, err
    return r


def _sheet_warnings(r):
    return [w for w in r.get("_warnings", []) if "シートを指す語が見当たりません" in w]


THREE = {"注文": ORDER, "商品表": MASTER, "旧単価表": ["コード", "単価"]}


# --- ① 名指しされていれば黙る ------------------------------------------------------------

def test_named_source_and_target_are_silent():
    r = _lookup("商品表から単価を注文シートに転記して", THREE)
    assert r["source_sheet"] == "商品表" and not _sheet_warnings(r)


def test_target_chosen_with_the_sheet_option_counts_as_named():
    r = _lookup("商品表から単価を転記して", THREE, cli=True)
    assert r["target_sheet"] == "注文" and not _sheet_warnings(r)


# --- ② 候補が 1 つなら機械が決める ---------------------------------------------------------

def test_two_sheet_book_without_names_is_decided_by_the_table():
    r = _lookup("単価を転記して", {"注文": ORDER, "商品表": MASTER})
    assert (r["target_sheet"], r["source_sheet"]) == ("注文", "商品表")
    assert not _sheet_warnings(r) and "source_sheet" not in r.get("_sources", {})


def test_the_only_sheet_holding_the_column_wins_over_the_guess():
    r = _lookup("注文シートに単価を転記して", {"注文": ORDER, "商品表": MASTER, "メモ": ["備考"]},
                source="メモ")
    assert r["source_sheet"] == "商品表"
    assert "商品表" in r["_sources"]["source_sheet"]
    assert not _sheet_warnings(r)


# --- ③ 候補が複数なら ⚠ ------------------------------------------------------------------

def test_two_reference_tables_without_a_name_are_disclosed():
    r = _lookup("注文シートに単価を転記して", THREE)
    ws = _sheet_warnings(r)
    assert len(ws) == 1 and "参照表" in ws[0]
    assert "『商品表』" in ws[0] and "『旧単価表』" in ws[0]
    assert r["source_sheet"] == "商品表"          # 書き換えない（推測で別の表へ移さない）


def test_neither_named_in_a_three_sheet_book_discloses_both():
    ws = _sheet_warnings(_lookup("単価を転記して", THREE))
    assert any("転記先" in w for w in ws) and any("参照表" in w for w in ws)


# --- ④ 断片は証拠にしない ------------------------------------------------------------------

def test_a_fragment_of_another_sheet_name_is_not_evidence():
    sheets = {"注文": ORDER, "単価表": ["コード", "単価"], "単価表_旧": ["コード", "単価"]}
    r = _lookup("注文シートに単価表から単価を転記して", sheets, source="単価表_旧")
    assert _sheet_warnings(r)
    r2 = _lookup("注文シートに単価表から単価を転記して", sheets, source="単価表")
    assert not _sheet_warnings(r2)


# --- ⑤ 黙る形 --------------------------------------------------------------------------------

def test_empty_task_and_a_single_sheet_book_are_silent():
    """★ 依頼文が空の経路は列の関所が先に断るので、判定そのものを直接呼んで縛る。"""
    args = {"target_sheet": "注文", "source_sheet": "商品表", "key_col": "コード",
            "target_col": "単価"}
    assert argcheck._lookup_sheets_from_request("", dict(args), _meta(THREE), THREE) == []
    one = {"注文": ORDER + ["単価"]}
    assert argcheck._lookup_sheets_from_request(
        "単価を転記して", {**args, "source_sheet": "注文"}, _meta(one), one) == []
    # 陽性対照: 同じ呼び方で、名指しの無い 3 枚の冊は鳴る
    assert argcheck._lookup_sheets_from_request("単価を転記して", dict(args), _meta(THREE), THREE)


# --- 変異: 配線を外すと黙る -----------------------------------------------------------------

def test_without_the_wiring_the_guess_would_pass_silently(monkeypatch):
    monkeypatch.setattr(argcheck, "_lookup_sheets_from_request", lambda *a, **k: [])
    r = _lookup("注文シートに単価を転記して", THREE)
    assert r["source_sheet"] == "商品表" and not _sheet_warnings(r)   # ← 直す前の挙動


def test_if_every_sheet_counted_as_named_the_guess_would_pass(monkeypatch):
    monkeypatch.setattr(argcheck, "_sheet_named_by_the_request", lambda *a, **k: True)
    r = _lookup("単価を転記して", THREE)
    assert r["target_sheet"] == "注文" and not _sheet_warnings(r)
