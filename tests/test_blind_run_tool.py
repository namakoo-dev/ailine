# -*- coding: utf-8 -*-
"""盲検の分母を確保する道具そのものの番人（2026-09-18）。

★ 出所: docs/FROZEN-20260917-盲検の合格線.md「4 体目を回す前にやること ①分母の確保」。
★ 道具の中身は tests/blind_run_core.py・入口は scripts/blind_run.py。
  ★ 中身を tests/ に置くのは、素の環境の番人が「宣言外の import」を全部止めるから
    ── 2026-09-18 にこの検体が**それで 4 本落ちた**（refresh_records・wiring_board に続く 3 度目）。

★★ この道具は**作った当日に嘘を 2 件出した**（記録として残す）:
  ・断りの回に「断ったのに書いた」── 数えていたのは `home/history.jsonl`。
    合格線の「1 バイトも書かない」は**利用者の冊**の話で、道具の台帳ではない。
    しかも履歴に断りを残すのは 2026-09-18 に足したばかりの**正しい**振る舞い。
  ・揺れの床に 1 件 ── バックアップのファイル名に**時刻**が入るので 2 回流せば必ず別名。
    揺れではなく時計だった。
  ★ 測定器が自分の粗さで欠陥を作る形。だからこの番人は **①緩めすぎて何も拾えない**
    ことと **②拾ってはいけない物を拾う**ことの**両方**を縛る。

★★ そして「拾えること」は**陽性対照**で確かめる ── 破れを仕込んで、正しい条文で
  名指しできるか。仕込まずに「0 件です」と言う道具は、番人として無意味。
"""
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
TOOL = REPO / "scripts" / "blind_run.py"

sys.path.insert(0, str(REPO / "tests"))


def _sheet(tmp_path, steps) -> str:
    """採点表を、仕込んだ流し直しの結果に対して出す。"""
    import blind_run_core as blind_run
    d = tmp_path / "003"
    d.mkdir(parents=True)
    (d / "replay_a.json").write_bytes(
        json.dumps(steps, ensure_ascii=False).encode("utf-8"))
    blind_run.RUNS = tmp_path
    import io
    import contextlib
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        blind_run.cmd_sheet("003")
    return buf.getvalue()


def _step(**kw):
    base = {"argv": ["run", "b.xlsx", "なにか"], "exit": 0, "stdout": "", "stderr": "",
            "wrote": [], "changed": [], "vanished": []}
    base.update(kw)
    return base


def test_the_sheet_catches_every_machine_condition(tmp_path):
    """★★ 陽性対照 ── 仕込んだ 4 つの破れを、**正しい条文で**名指しすること。

    ★ これが無いと「0 件でした」としか言わない道具になる（緩めすぎの検出）。
    """
    said = _sheet(tmp_path, [
        _step(exit=0, stdout="？ できませんでした"),                 # 断り4
        _step(exit=3, wrote=["b.xlsx"], stdout="？ だめです --copy"),  # 断り5
        _step(exit=0, vanished=["b.out.xlsx"]),                      # 到達3
        _step(exit=3, stdout="？ できませんでした。"),                 # 断り3
    ])
    assert "断りの顔で exit 0（断り4）" in said, said
    assert "断ったのに書いた（断り5）" in said, said
    assert "消えた物" in said and "（到達3）" in said, said
    assert "逃げ道も通る書き方も示していない（断り3）" in said, said


def test_the_sheet_is_quiet_when_nothing_is_broken(tmp_path):
    """★ 陰性対照: 正しく通った回・正しく断った回を破れと呼ばない。"""
    said = _sheet(tmp_path, [
        _step(exit=0, wrote=["b.out.xlsx"], stdout="✓ 機械検証済みの内容です"),
        _step(exit=3, stdout="？ できません（例:『金額が5000以上の行を抜き出して』）"),
        _step(exit=7, stdout="？ 止めました ── --overwrite を付けてください"),
    ])
    assert "機械で拾える破れ: 0 件" in said, said


def test_the_tool_only_watches_the_users_books(tmp_path):
    """★★ 道具の台帳を「利用者の成果物」と数えないこと（当日の嘘 2 件の根）。

    ★ `home/` 配下（履歴・バックアップ・別名簿）は利用者から見えない物。
      ここを数えると、**断りの回に必ず「断ったのに書いた」が立つ**（履歴が増えるので）。
    """
    import blind_run_core as blind_run
    w = tmp_path / "w"
    (w / "home" / "backups").mkdir(parents=True)
    (w / "b.xlsx").write_bytes(b"book")
    (w / "home" / "history.jsonl").write_bytes(b"{}")
    (w / "home" / "backups" / "x.xlsx").write_bytes(b"bak")
    seen = blind_run._digests(w)
    assert "b.xlsx" in seen, seen
    assert not [k for k in seen if k.startswith("home/")], (
        f"道具の台帳を利用者の成果物として数えている: {sorted(seen)}")


def test_the_sheet_says_what_it_did_not_judge(tmp_path):
    """★★ 「機械が合格と言った」を独り歩きさせない ── 見ていない 4 条を必ず書く。"""
    said = _sheet(tmp_path, [_step()])
    assert "人が読む" in said, said
    assert "合否はこの表では決まらない" in said, said
    for word in ("中身が正しい", "本当の理由を言う", "文言が正確"):
        assert word in said, f"見ていない条を挙げていない: {word}"


def test_freeze_refuses_without_the_initial_books(tmp_path):
    """★ 触られた後の冊では流し直せない ── 初期の冊が無ければ**名指しで止まる**。"""
    r = subprocess.run([sys.executable, str(TOOL), "freeze", "999", str(tmp_path)],
                       cwd=str(REPO), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    assert r.returncode == 3, r.stdout
    assert "初期の冊" in r.stdout, r.stdout


def test_the_tool_is_documented_as_partial():
    """★ 道具の口上が「5 条だけ」と名乗っていること（凍結文書と食い違わせない）。

    ★ 見るのは**中身**（tests/blind_run_core.py）── 入口（scripts/blind_run.py）は
      素の環境の遮断を避けるための薄い殻で、口上はそこには置かない。
    ★ 2026-09-18 に分割した時、この検体だけ入口を見たまま残って赤くなった
      ── 移したら**見る先も一緒に移す**（片方だけ直さない）。
    """
    core = (REPO / "tests" / "blind_run_core.py").read_text(encoding="utf-8")
    assert "機械で決まるのは **5 つだけ**" in core, "部分的であることを名乗っていない"
    assert "FROZEN-20260917" in core, "凍結文書を指していない"
    shell = TOOL.read_text(encoding="utf-8")
    assert len(shell.splitlines()) < 60, (
        f"入口が薄くない（{len(shell.splitlines())} 行）── 中身が入口へ戻っている")
