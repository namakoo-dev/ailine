# 翻訳精度の数字を、文書と実測で縛る ── 2026-08-29。
#
# ★★ 見つかった食い違い: README は「op 一致 51/52 = 98.1%」と書いていたが、
#   同じ凍結検体を今日 2 回走らせると **49/52 = 94.2%** だった。
#   OPS_DOC に 16 行足した回の低下（98.1%→94.2%・当時 実測済み）が、
#   **文書に反映されないまま残っていた**。
#   ★ これは面接で口に出す数字だ。手で守れない数字は機械が守る
#     （試験の本数・主ファイルの行数と同じ扱いにする）。
#
# 二段構え:
#   ① 非 local: 文書の数字と**記録**（tests/battery_recorded.json）が一致すること
#   ② local:    記録と**実測**が一致すること（実物の LLM が要るので実機側）

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

from _doc_numbers import assert_all_agree

REPO = Path(__file__).resolve().parent.parent
RECORD = Path(__file__).resolve().parent / "battery_recorded.json"


def _record():
    return json.loads(RECORD.read_text(encoding="utf-8"))


def test_the_readme_matches_the_record():
    """① 文書 vs 記録。★ 数字を直すときは記録も直す（片方だけ動かせない形にする）。"""
    rec = _record()
    want = f"{rec['op_correct']}/{rec['op_total']} = {rec['op_correct'] / rec['op_total'] * 100:.1f}%"
    # ★ 2026-08-31: README だけでなく、印のある文書を**全部**見る（tests/_doc_numbers.py）。
    assert_all_agree("BATTERY_OP", want)


def test_the_runbook_does_not_quote_a_stale_number():
    """★ 手順書は当日そのまま読み上げる紙 ── 古い数字が残っていたら赤くする。"""
    rec = _record()
    pct = f"{rec['op_correct'] / rec['op_total'] * 100:.1f}%"
    text = (REPO / "demo" / "手順.md").read_text(encoding="utf-8")
    nums = set(re.findall(r"op 一致 ([0-9]+\.[0-9]%)", text))
    assert nums <= {pct}, f"手順書に古い数字が残っている: {nums - {pct}}（今は {pct}）"


def test_the_record_names_how_it_was_measured():
    """★ 数字だけを残さない ── いつ・どのモデルで・何回で出たかを一緒に置く
       （後から誰が見ても、同じ条件で測り直せる形にしておく）。"""
    rec = _record()
    for key in ("measured_on", "model", "runs", "misassert", "misassert_total"):
        assert rec.get(key) not in (None, ""), key
    assert rec["runs"] >= 2, "1 回の測定を記録にしない（LLM は揺れる）"



def test_the_documents_do_not_drift_around_the_ratio():
    """★ 比だけでなく**件数と断りの数**も印で縛る（2026-09-04）。

    ★ なぜ足したか: MATRIX の印は**比の文字列しか守っていなかった**。
      190 → 221 に増やした回、印の中は番人が直させたのに、その**周りの地の文**が
      4 箇所ずれていた ── README の「断り 3」、手順書の「下の 190 件」「残りの 1 件」、
      そして**想定問答の「93 件中 92 件（98.9%）」**（93 件時代のまま）。
      ★ 想定問答は当日そのまま読み上げる紙で、いちばんずれてはいけない所だった。
      「在っても、その事故の形では鳴らない」の実例なので、鳴る形にする。
    """
    m = _matrix()
    assert_all_agree("MATRIX_CASES", str(m["cases"]), at_least=4)
    assert_all_agree("MATRIX_REFUSED", str(m["refused"]), at_least=3)


def test_the_interview_script_is_bound_to_the_record_too():
    """★ 想定問答にも MATRIX の印が在ること（3 文書の中に居ることを名指しで確かめる）。

    ★ at_least だけだと「どこか 3 つ」で通ってしまい、**読み上げる紙が抜けても**
      鳴らない。抜けてはいけない紙は名指しで要求する。
    """
    from _doc_numbers import marked
    files = {str(p) for p, _ in marked("MATRIX")}
    assert any("想定問答" in f for f in files), (
        "demo/想定問答.md に MATRIX の印が無い ── 当日読み上げる紙が縛られていない")


# ★ 「N 件中」と書いてよい (文書, N) の対。
#   ★★ 数字だけの白名簿にしない ── 2026-09-04 の変異試験で見つけた穴:
#     想定問答のために 93 を許したら、**README に「93 件中」を書いても通った**。
#     ある紙のための許可が、全部の紙に効いてしまう。だから対で持つ。
#   ★ 白名簿は宣言駆動 ── 増やすときは、なぜ歴史として残すのかをここに書く。
_CASES_IN_TEXT_ALLOWED = {
    ("PREREG_translation_v7.md", 2),   # 事前登録（別の測定・2 件の予測）
    ("想定問答.md", 10),                # 「1B だと 10 件中 6 件が的外れ」（モデル比較）
    ("想定問答.md", 93),                # Q8 のモデル比較表 ── 2026-08-30 に 93 件で測った
                                        #   **過去の測定**。小さいモデルはその後測り直して
                                        #   いないので分母は当時のまま残す（測り直さずに
                                        #   分母だけ書き換えると、測っていない数字を主張する）
}


def test_no_document_quotes_a_superseded_case_count():
    """★ 「N 件中」の N が、記録の件数でも白名簿でもなければ赤くする（2026-09-04）。

    ★ なぜ足したか: MATRIX の印は**印の中しか守らない**。同じ日に 2 度、
      印の外で古い数字が見つかった ──
      ① README「断り 3」/ 手順書「下の 190 件」「残りの 1 件」
      ② 想定問答「93 件中 92 件（98.9%）」（41 行目・印を置いて解決）と、
         **同じ紙の 206 行目にもう 1 つ**「93 件中 98.9%」（こちらは見落としていた）。
      ★ 1 箇所直して安心した所に 2 つ目が在った。だから「印を置く」だけでなく
        **古い分母を名指しで禁じる**側も要る。
    """
    import re
    rec = _matrix()
    cur = int(rec["cases"])
    bad = []
    for path in sorted(REPO.rglob("*.md")):
        if any(x in str(path) for x in (".git", "node_modules", "CHANGELOG")):
            continue
        text = path.read_text(encoding="utf-8")
        # 印の中は別の番人が見ているので外す
        text = re.sub(r"<!-- [A-Z_]+ -->.*?<!-- /[A-Z_]+ -->", "", text, flags=re.S)
        for m in re.finditer(r"(\d+)\s*件中", text):
            n = int(m.group(1))
            if n == cur:
                continue
            if any(f in str(path) and n == k for f, k in _CASES_IN_TEXT_ALLOWED):
                continue
            bad.append((str(path.relative_to(REPO)), m.group(0)))
    assert not bad, (
        f"古い分母が文書に残っている: {bad}（今は {rec['cases']} 件）── "
        "歴史として残すなら _CASES_IN_TEXT_ALLOWED に理由つきで足すこと")


def test_no_document_claims_zero_failures_when_there_are_failures():
    """★ 記録に失敗が在る間は、どの文書も「壊した 0 / 壊していない 100%」と書けない。

    ★ なぜ足したか（2026-09-04）: 段2.5 で**語彙の中の穴**が 2 件見つかり、
      記録の failed が 0 でなくなった。この時いちばん危ないのは、
      **見出しの主張だけが古いまま残る**こと ── しかもそれは
      「読み上げる紙」に載っている。数字でなく**主張**を記録に縛る。
    ★ 逆向き（failed が 0 に戻ったのに 99.1% と書いたまま）は
      MATRIX の印が捕まえる。
    """
    import re
    rec = _matrix()
    if int(rec["failed"]) == 0:
        return
    banned = ("壊した 0", "壊していない 100%", "失敗 0")
    # ★ 過去の別の測定として残す文言（★ 数字だけでなく**文書と文言の対**で許す ──
    #   2026-09-04 の変異試験で「数字だけの白名簿は全部の紙に効く」穴を踏んだ）。
    allowed = {("想定問答.md", "壊した 0→9")}   # 語彙を広げた実験（0→9）の対比
    bad = []
    for path in sorted(REPO.rglob("*.md")):
        if any(x in str(path) for x in (".git", "node_modules", "CHANGELOG")):
            continue
        text = path.read_text(encoding="utf-8")
        for word in banned:
            for m2 in re.finditer(re.escape(word), text):
                around = text[m2.start():m2.start() + 24]
                if any(f in str(path) and around.startswith(w) for f, w in allowed):
                    continue
                bad.append((str(path.relative_to(REPO)), word))
                break
    assert not bad, (
        f"記録の失敗が {rec['failed']} 件なのに「壊していない」と書いた文書がある: {bad}")

@pytest.mark.local
def test_the_record_still_matches_the_machine():
    """② 記録 vs 実測。実物の ollama が要るので実機側。
       ★ ここが赤くなったら、直すのは**記録と文書**（実測が正）。"""
    r = subprocess.run(
        [sys.executable, str(REPO / "bench" / "translation_dsl_battery_run.py"),
         _record()["model"]],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=2400, cwd=str(REPO / "bench"),
        env={**__import__("os").environ, "PYTHONPATH": str(REPO / "src")})
    assert r.returncode == 0, r.stdout[-600:]
    m = re.search(r"op 分類: (\d+)/(\d+)", r.stdout)
    assert m, r.stdout[-600:]
    rec = _record()
    got, total = int(m.group(1)), int(m.group(2))
    assert total == rec["op_total"], f"検体の数が変わった（凍結のはず）: {total}"
    # ★ LLM は揺れる ── 1 件のぶれは許すが、それ以上ずれたら記録を測り直す
    assert abs(got - rec["op_correct"]) <= 1, (
        f"実測 {got}/{total}・記録 {rec['op_correct']}/{total} ── "
        "記録と文書を測り直して直すこと（実測が正）")


# --- ★★ 効果で測る検体（bench/basic_ops_matrix.py）の数字も同じ二段構えで縛る ---------
#
# ★ 2026-08-29: README は「84 件・97.6%」と地の文で書いていた。上の翻訳精度と違って
#   印も記録も無く、**手で守る数字**だった。同じ日に検体を 93 件へ増やしたので、
#   その場で古くなる ── 手で守れない数字は機械が守る（この repo の規範）。


def _matrix():
    return _record()["matrix"]


def test_the_readme_matches_the_matrix_record():
    """① 文書 vs 記録。"""
    m = _matrix()
    want = f"{m['intended']}/{m['cases']} = {m['intended'] / m['cases'] * 100:.1f}%"
    # ★★ 2026-09-02: 印が README にしか無く、demo/手順.md と docs/なぜこの形か.md に
    #   93 件時代の数字が**残っていた**（番人は全 .md を走査するのに、印が無い所は
    #   見えない ── 在っても鳴らない）。3 文書に印を置き、**3 箇所以上**を要求する。
    assert_all_agree("MATRIX", want, at_least=3)


def test_the_matrix_record_names_how_it_was_measured():
    m = _matrix()
    for key in ("measured_on", "model", "cases", "intended", "refused", "failed"):
        assert m.get(key) is not None, key
    assert m["intended"] + m["refused"] + m["failed"] == m["cases"], m


@pytest.mark.local
def test_the_matrix_record_still_matches_the_machine():
    """② 記録 vs 実測（実物の LLM と LibreOffice が要る・13 分ほど掛かる）。"""
    r = subprocess.run(
        [sys.executable, str(REPO / "bench" / "basic_ops_matrix.py")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=3600, cwd=str(REPO),
        env={**__import__("os").environ, "PYTHONPATH": str(REPO / "src")})
    assert r.returncode == 0, r.stdout[-800:]
    m = re.search(r"合計 (\d+) 件: ✓ (\d+)\s+？断り (\d+)\s+× 失敗 (\d+)", r.stdout)
    assert m, r.stdout[-800:]
    rec = _matrix()
    cases, ok = int(m.group(1)), int(m.group(2))
    assert cases == rec["cases"], f"検体の数が変わった: {cases}（記録は {rec['cases']}）"
    assert abs(ok - rec["intended"]) <= 1, (
        f"実測 {ok}/{cases}・記録 {rec['intended']}/{cases} ── "
        "記録と文書を測り直して直すこと（実測が正）")
