"""本番の走行は、試験の走行から学ばない（2026-09-22）。

★★ 実測した事故: 本番の台帳 `~/.ailine/history.jsonl` は 19,046 行のうち
  **18,292 行（96.0%）が試験の走行**だった。私の記録には「4,011 行・22%」と
  書いてあり、**その数字が間違っていた**（手で数えず機械で数え直して分かった）。

★★ 実害の経路が 2 本、実測で生きていた:
  ① 思い出し（`op_that_worked_before`）が見る窓 500 行は **100% が試験**だった。
     道具が見せる例（23 種）と同じ依頼文の ok 記録が **1,370 行**あり、
     買い手が例をそのまま打つと「前回 <日付> にこの操作で通りました」と出る ──
     **買い手はその走行をやっていない。** 4 体目の買い手が導入を断った理由は
     「作業記録が嘘をつく」だった。同じ族。
  ② 同じ窓を「この道具が過去にそこへ書いたか」の上書き判定も使う。窓が試験で
     埋まると、**道具が自分で作った .out が記録から落ちて「人のファイル」扱い**になる
     ── 5 体目の指摘①と同じ形。

★ 台帳そのものは退避で掃除した（消していない・`history-test.jsonl` に移した）。
  ここが縛るのは**再発防止**の方 ── 誰かが隔離を忘れて走らせても、印が付く。

★ なぜ `AILINE_HOME` で分けないか: 汚染は「隔離を忘れて**既定の** home に書いた」
  ときに起きた。その印では捕まらない。`PYTEST_CURRENT_TEST` は pytest が環境に立てる
  **宣言**で、subprocess にも継承される ── 推測でなく宣言で分けられる唯一の材料。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402


def test_a_run_inside_a_test_is_marked(tmp_path):
    """★ 試験から始まった走行には印が付く（隔離を忘れていても）。"""
    p = tmp_path / "h.jsonl"
    ailine.append_history({"ts": "2026-09-22T00:00:00+00:00", "task": "並べ替えて",
                           "ok": True, "book": str(tmp_path / "b.xlsx")}, path=p)
    row = json.loads(p.read_text(encoding="utf-8").splitlines()[0])
    assert row.get("from_test") is True, row
    assert row["task"] == "並べ替えて", "既存の欄を触っていないこと"


def test_a_real_run_does_not_see_test_rows(tmp_path, monkeypatch):
    """★★ 本体 ── 本番の走行からは、印の付いた行が見えない。"""
    p = tmp_path / "h.jsonl"
    ailine.append_history({"ts": "2026-09-22T00:00:01+00:00", "task": "試験の行",
                           "ok": True, "book": "x.xlsx"}, path=p)
    with p.open("a", encoding="utf-8") as f:                 # 印の無い（＝本番の）行
        f.write(json.dumps({"ts": "2026-09-22T00:00:02+00:00", "task": "本番の行",
                            "ok": True, "book": "y.xlsx"}, ensure_ascii=False) + chr(10))

    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)  # ★ 本番の走行を真似る
    assert not ailine.run_started_from_a_test()
    tasks = [e.get("task") for e in ailine.read_history(path=p, max_n=50)]
    assert tasks == ["本番の行"], tasks


def test_a_test_run_still_sees_its_own_rows(tmp_path):
    """★ 試験の中からは見える ── 塞ぐと、思い出し・上書き判定の番人が死ぬ。

    ★ 契約は「本番は試験を見ない」であって「印の行を消す」ではない。
    """
    p = tmp_path / "h.jsonl"
    ailine.append_history({"ts": "2026-09-22T00:00:03+00:00", "task": "試験の行",
                           "ok": True, "book": "x.xlsx"}, path=p)
    tasks = [e.get("task") for e in ailine.read_history(path=p, max_n=50)]
    assert tasks == ["試験の行"], tasks


def test_the_recall_does_not_cite_a_test_run(tmp_path, monkeypatch):
    """★ 実害の経路①を名指しで縛る ── 思い出しが試験の走行を根拠にしないこと。"""
    p = tmp_path / "h.jsonl"
    ailine.append_history({"ts": "2026-09-22T00:00:04+00:00", "task": "見出しを太字にして",
                           "ok": True, "book": "x.xlsx", "op": "BOLD"}, path=p)
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    op, on = ailine.op_that_worked_before("見出しを太字にして",
                                          ailine.read_history(path=p, max_n=500))
    assert op is None and on is None, (op, on)


def test_old_rows_without_the_mark_still_pass(tmp_path, monkeypatch):
    """★ 印の無い古い行は通る（過去の台帳を壊さない）。"""
    p = tmp_path / "h.jsonl"
    p.write_text(json.dumps({"ts": "2026-08-01T00:00:00+00:00", "task": "古い行",
                             "ok": True, "book": "z.xlsx"}, ensure_ascii=False) + chr(10),
                 encoding="utf-8")
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    assert [e.get("task") for e in ailine.read_history(path=p, max_n=5)] == ["古い行"]
