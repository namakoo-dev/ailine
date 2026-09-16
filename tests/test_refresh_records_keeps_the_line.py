# -*- coding: utf-8 -*-
"""記録を作り直す道具が、**触ってよい線**を越えないこと（2026-09-16）。

★ 出所（Namakoo「これって ailine をいじる度に更新しないといけないよね」）:
  2026-09-16 の 1 日で、記録ファイルを触った commit は行数 5 回 / README 10 回 / 図 3 回。
  数え直すだけの作業が push を 2 回止めた。だから 1 コマンドにまとめた
  （scripts/refresh_records.py）。

★★ ただし記録は 2 種類あり、扱いを逆にすると害になる:
  ① 導出でしかないもの（行数・試験の本数・依存の図）── 作り直してよい
  ② 測定を記録したもの（効果の行列・翻訳精度）── **絶対に自動更新しない**
     自動で揃えたら記録は常に実測と一致し、二度と警告しなくなる ── 恒真。
     同じ日に、依存の図が「生成器と番人が同じ盲点を共有していたせいで
     13 日間 48% 間違ったまま緑」だったのを見つけている。

★ この番人は「①を直せること」と「②を直さないこと」を**両方向**で見る。
  片方だけだと、道具が何もしなくても・全部書き換えても緑になりうる。
"""
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
TOOL = REPO / "scripts" / "refresh_records.py"
BUDGET = REPO / "tests" / "ailine_py_line_budget.txt"
RECORD = REPO / "tests" / "battery_recorded.json"


def _run(*args):
    return subprocess.run([sys.executable, str(TOOL), *args], cwd=str(REPO),
                           capture_output=True, text=True, encoding="utf-8", errors="replace")


def _touchable():
    """この道具が書き換えうるファイル（退避の対象）。

    ★★ 初版はこれが無く、修復を試す試験の中の `--write` が **README まで書き換えていた**
      ── そのせいで別の試験が偶然緑になった（1 回目は赤、2 回目は緑）。
      作業木を静かに変える試験は、次に走る試験の結果を汚す。触りうるものは全部退避する。
    """
    out = [BUDGET, REPO / "docs" / "依存関係.md"]
    out += [p for p in sorted(REPO.rglob("*.md"))
            if ".git" not in str(p) and "node_modules" not in str(p)]
    return [p for p in out if p.exists()]


class _Frozen:
    """触りうるファイルを丸ごと退避し、抜けるときに**バイト単位で**戻す。"""

    def __enter__(self):
        self._snap = {p: p.read_bytes() for p in _touchable()}
        return self

    def __exit__(self, *exc):
        for p, b in self._snap.items():
            if p.read_bytes() != b:
                p.write_bytes(b)
        return False


def test_it_reports_clean_when_records_match():
    """★ 空回りの検出 ── 揃っているのに何か言うなら、下の試験の意味が無くなる。"""
    r = _run()
    assert r.returncode == 0, f"揃っているのに古いと言っている:\n{r.stdout[-600:]}"


def test_it_detects_and_repairs_a_derived_record():
    """① ずらしたら気づき、--write で**元のバイトに**戻すこと。"""
    with _Frozen():
        before = BUDGET.read_bytes()
        BUDGET.write_bytes(b"99999\n")
        assert _run().returncode == 1, "行数がずれているのに気づかない"
        w = _run("--write")
        assert w.returncode == 0, w.stdout[-400:]
        assert BUDGET.read_bytes() == before, "作り直した結果が元と違う"


def test_it_never_touches_a_measured_record():
    """② ★★ 実機を回して初めて出る数字には手を出さないこと。

    ★ ここが破れると、記録は常に実測と一致し、**二度と警告しなくなる**。
      壊したまま残るのが正しい振る舞い（人が測り直して書く）。
    """
    before = RECORD.read_bytes()
    d = json.loads(before.decode("utf-8"))
    d["matrix"]["intended"] = 1          # ★ わざと嘘の数にする
    RECORD.write_bytes((json.dumps(d, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))
    try:
        with _Frozen():
            _run("--write")
        after = json.loads(RECORD.read_bytes().decode("utf-8"))
        assert after["matrix"]["intended"] == 1, (
            "測定の記録を道具が書き換えた ── 記録が常に実測と一致するようになり、"
            "番人が二度と鳴らなくなる（恒真）")
    finally:
        RECORD.write_bytes(before)


def test_the_tool_says_what_it_will_not_touch():
    """★ 口上を縛る ── 触らないものを画面に出すこと（次に読む者が線を知れるように）。"""
    r = _run()
    assert "触らないもの" in r.stdout, "触らない線を画面に出していない"
    assert "MATRIX" in r.stdout
    assert "恒真" in r.stdout, "なぜ触らないのか（恒真になる）を書いていない"


def test_the_measured_marks_list_is_not_empty():
    """★ 空回りの検出 ── 守る対象の名簿が空なら、②の試験は自明に通る。"""
    sys.path.insert(0, str(REPO / "scripts"))
    import refresh_records
    assert len(refresh_records.MEASURED_MARKS) >= 3, refresh_records.MEASURED_MARKS
    assert "MATRIX" in refresh_records.MEASURED_MARKS
