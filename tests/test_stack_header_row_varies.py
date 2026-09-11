# -*- coding: utf-8 -*-
"""縦積み: 冊ごとに見出し行の位置が違っても積めること（2026-09-11）。

★★ 実装より先に書いた赤い検体。

  実測（部署別 30 ファイル・2026-09-11）: **20 冊が積めなかった**。
  理由はどれも同じで `欠け: 日付, 部署, 科目, 金額, 備考 ／ 余り: 開発課 経費明細`。

  原因: `cmd_stack` は見出し行を**基準 1 冊から 1 回だけ推定**し、その行番号を
  全ファイルに当てている。実物のフォルダは、同じ様式でも飾り行の数が違う ──
  30 冊の見出し行は 1 行目 10 冊 / 2 行目 10 冊 / 3 行目 10 冊 に散っていた。

  ★ 検体を書いた者（規則を知らない者）が、あらかじめ「落とし方」に
    **『基準 1 冊の番地に固定した実装』**と名指ししていた罠でもある。

★★ ただし緩めてはいけない線が 1 本ある:
  「行を探しに行く」を部分一致で許すと、**本当に列が違う冊**がどこかの行で
  たまたま一致して積まれてしまう（`test_unmatched_file_is_skipped_named_and_counted`
  が守っている契約が骨抜きになる）。
  → 採用するのは**基準の見出しが全部そろった行だけ**。下の負の被覆で縛る。
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
HDRS = ["日付", "部署", "科目", "金額"]


def _book(path: Path, headers, rows=(), *, decor=0, title="経費明細"):
    """飾り行を `decor` 行だけ上に置いた冊を作る（実物の形）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    for k in range(decor):
        ws.append([f"{title}" if k == 0 else "2026 年 8 月分"])
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def _stack(folder: Path, out: Path, *extra):
    return subprocess.run(
        [sys.executable, "-m", "ailine", "stack", str(folder), "--out", str(out),
         "--json", *extra],
        capture_output=True, text=True, timeout=180, encoding="utf-8",
        errors="replace", cwd=str(REPO))


def _rows_of(out: Path) -> list:
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    got = [[c.value for c in row] for row in ws.iter_rows()]
    wb.close()
    return got


@pytest.fixture()
def folder(tmp_path):
    """同じ様式・飾り行だけ 0/1/2 行と違う 3 冊。"""
    d = tmp_path / "src"
    _book(d / "a_営業一課.xlsx", HDRS,
          [("2026-08-01", "営業一課", "用紙代", 1000)], decor=0)
    _book(d / "b_経理課.xlsx", HDRS,
          [("2026-08-02", "経理課", "運送費", 2000)], decor=1, title="経理課 経費明細")
    _book(d / "c_総務課.xlsx", HDRS,
          [("2026-08-03", "総務課", "保守料", 3000)], decor=2, title="総務課 経費明細")
    return d


def test_files_whose_header_row_differs_are_still_stacked(folder, tmp_path):
    """★ 飾り行の数が違うだけの冊が、全部積めること。

    ★ 実装前はここが赤だった ── 3 冊中 2 冊が
      「欠け: 日付, 部署, 科目, 金額 ／ 余り: 経理課 経費明細」で落ちる。
    """
    out = tmp_path / "積み上げ.xlsx"
    r = _stack(folder, out)
    assert r.returncode == 0, f"落ちた: rc={r.returncode}\n{r.stdout}\n{r.stderr}"
    assert out.exists(), r.stdout

    rows = _rows_of(out)
    body = [x for x in rows[1:] if x and x[0] not in (None, "")]
    assert len(body) == 3, f"★ 3 冊ぶん積めていない（{len(body)} 行）:\n{r.stdout}"

    amounts = {x[3] for x in body}
    assert amounts == {1000, 2000, 3000}, f"★ 金額が揃わない: {amounts}"

    depts = {x[1] for x in body}
    assert depts == {"営業一課", "経理課", "総務課"}, f"★ 部署が揃わない: {depts}"


def test_the_decoration_line_never_becomes_a_data_row(folder, tmp_path):
    """★ 負の被覆 ── 飾り行が明細として積まれていないこと。

    ★ 「見出し行を探す」を入れると、今度は**飾り行を 1 行目のデータとして積む**
      向きの事故が生える。消えたものは diff に出ないが、増えたものは出る。
    """
    out = tmp_path / "積み上げ.xlsx"
    _stack(folder, out)
    flat = [str(v) for row in _rows_of(out) for v in row if v is not None]
    # ★ 先に「積めている」ことを要求する ── これが無いと、出力が空でもこの試験は通る
    #   （出ないことは信号ではない・docs/開発手法.md）。
    assert any("営業一課" in v for v in flat), f"★ そもそも積めていない: {flat[:12]}"
    assert any("総務課" in v for v in flat), f"★ 飾り 2 行の冊が積めていない: {flat[:12]}"
    assert not any("経費明細" in v for v in flat), f"★ 飾り行を積んだ: {flat[:12]}"
    assert not any("2026 年 8 月分" == v for v in flat), "★ 飾り行を積んだ"


def test_a_file_with_genuinely_different_columns_is_still_refused(folder, tmp_path):
    """★★ 緩めてはいけない線 ── 列が本当に違う冊は、どの行でも採用しない。

    ★ 「行を探しに行く」を部分一致で許すと、無関係な冊がどこかの行で
      たまたま一致して積まれる。採用するのは**基準の見出しが全部そろった行だけ**。
    """
    _book(folder / "z_別物.xlsx", ["氏名", "電話", "住所"],
          [("山田", "03-0000-0000", "東京")], decor=2, title="連絡先一覧")
    out = tmp_path / "積み上げ.xlsx"
    r = _stack(folder, out)

    assert "z_別物.xlsx" in r.stdout, f"★ 積めなかった冊を名指ししていない:\n{r.stdout}"
    flat = [str(v) for row in _rows_of(out) for v in row if v is not None]
    assert not any("山田" in v for v in flat), f"★ 列が違う冊を積んだ: {flat[:12]}"


def test_the_json_counts_every_file_in_the_denominator(folder, tmp_path):
    """★ 分母は入力の冊数。積めた数を分母にしない。"""
    out = tmp_path / "積み上げ.xlsx"
    r = _stack(folder, out)
    payload = None
    for line in r.stdout.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                continue
    assert payload is not None, f"--json の出力が読めない:\n{r.stdout}"
    assert payload["denominator"] == 3, payload
    assert payload["stacked_files"] == 3, payload


# ── ★ 開示は「黙って別の行を読まない」こと ─────────────────────
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _run_route(route: str, folder: Path, out: Path, monkeypatch, capsys):
    """経路ごとに 1 回走らせて、人に見えた文字を返す。

    ★★ `run` は 7B を使わない ── 翻訳は `translate_task` の monkeypatch
      （tests/test_run_folder.py と同じ作法・製品コードにテスト用の口を彫らない）。
      ★ ここを subprocess + 実 LLM で書いたら、素の環境（ollama を港 9 に向ける）で
        落ちた。手元に在るものに依存して CI で落ちる形 ── この repo が何度も踏んだ
        「居るから見えない」。測りたいのは翻訳の質ではなく**開示の配線**。
    """
    if route == "stack":
        rc = ailine.main(["stack", str(folder), "--out", str(out)])
    else:
        monkeypatch.setattr(
            ailine, "translate_task",
            lambda model, task, book_meta, temperature=0.1: {
                "plan": [{"op": "EXTRACT",
                          "args": {"column": "金額", "cmp": "gte", "value": 1}}]})
        rc = ailine.main(["run", str(folder), "金額が 1 以上の行を集めて"])
    return rc, capsys.readouterr().out


@pytest.mark.parametrize("route", ["stack", "run"])
def test_using_a_different_header_row_is_disclosed_on_every_route(
        folder, tmp_path, monkeypatch, capsys, route):
    """★★ 基準と違う行を見出しとして読んだら、**どの経路でも**そう言うこと。

    ★★ なぜ 1 本で 2 経路を縛るか:
      この集約は `cmd_stack` と `cmd_run_folder` の 2 箇所に写し取られていて、
      表示も `render_stack_report` と `cmd_run_folder` の自前 `say` の 2 箇所に在る。
      片方だけ直すのが既定で起きる（この repo は 2026-08-21 以降この形を何度も踏んだ）。
      → 経路を parametrize して、**片方を黙らせたら必ず赤くなる**ようにする。

    ★ 黙って別の行を読むのは、この repo でいちばん高くつく失敗の形。
      飾り行の数が冊ごとに違うのは実物では普通だが、**そう言わずに読む**のは別問題。
    """
    out = tmp_path / f"出力_{route}.xlsx"
    # ★ skip にしない。skip は番人ではない ── その経路が裸のまま緑に見える。
    rc, text = _run_route(route, folder, out, monkeypatch, capsys)
    assert rc == 0, f"★ {route} が落ちた: rc={rc}\n{text}"

    assert "見出しは 2 行目" in text, (
        f"★ {route}: 基準と違う行を読んだのに黙っている\n{text}")
    assert "基準は 1 行目" in text, f"★ {route}: 基準の行を示していない\n{text}"


def test_the_json_names_the_files_that_could_not_be_stacked(folder, tmp_path):
    """★ `--json` が積めなかった冊を**名指し**すること。

    ★ 2026-09-11 まで、--json は denominator と stacked_files の**数だけ**で、
      どの冊が落ちたかを言えなかった（人向けの表示は名指ししていた）。
      自動化する側は「4 冊中 1 冊」とだけ知らされ、残り 3 冊を聞けない。
      ★ 同じ事実に出口が 2 つあって、片方にだけ名前が載っていた。
    """
    _book(folder / "z_別物.xlsx", ["氏名", "電話", "住所"],
          [("山田", "03-0000-0000", "東京")], decor=2, title="連絡先一覧")
    out = tmp_path / "積み上げ.xlsx"
    r = _stack(folder, out)
    payload = None
    for line in r.stdout.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                continue
    assert payload is not None, r.stdout
    names = {f["name"] for f in payload.get("skipped", ())}
    assert "z_別物.xlsx" in names, f"★ 落ちた冊が --json に名指しで無い: {payload}"
    assert all(f.get("reason") for f in payload["skipped"]), "★ 理由が空"
    used = {f["name"]: f["used_row"] for f in payload.get("header_row_fallbacks", ())}
    assert used.get("b_経理課.xlsx") == 2, f"★ 使った見出し行が --json に無い: {payload}"
