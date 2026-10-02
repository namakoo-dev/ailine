# -*- coding: utf-8 -*-
"""断りが示した道を**実際に歩く**（2026-09-19・導通試験）。

★★ なぜ要るか（Namakoo「到達率を極大にするのは大事だけど、**漏らした依頼を適切に判断し
  ユーザを正解に導いてやること**までやるから」）: 合格線の 3 条目「通る道を示す」は、
  機械が **escape（逃げ道の語が本文に在る）/ example（例が在る）** しか見ていなかった。
  **その道が通るか**は誰も確かめていない ── 今日ずっと潰してきた「宣言 vs 実体」の隙間。

★★ 判定の芯は 5 つ（Namakoo 承認・2026-09-19）:

    walked      示した道を歩いたら着いた
    by_design   意図して行き止まり（パスの打ち間違い等・unlock が「無い」）
    vague       ★ 理由が粗くて道を示せていない（直せる見込みが高い）
    no_path     本当に道が無い（機能・語彙が無い）
    path_fails  ★★ 歩いたが着かなかった ── **通らない道を示した**（最悪）

  ★ 測れなかった回を**別に持つ**（判定と混ぜない ── 見ていないものを見たことにしない）:

    未調査        この器では歩けない（2 冊と実表が要る等）
    未記入        歩き方を台帳に書いていない
    引き金が引けない  断りが出ずに到達した（検体が古い ── 直すのは検体の側）
    歩けなかった   ★★ 機械が塞がっていた（exit 6）── 2026-09-20 に足した。
                 旧版はこれを path_fails と読み、**測れなかったのに「通らない」と
                 主張していた**（一度に 10 件が偽の赤）。

  ★ `path_fails` は `no_path` より重い。道が無いのは不足だが、通らない道を示すのは誤情報。
  ★ 道具を作る前に 1 件を手で歩いた時点で `path_fails` が 1 件出ている
    （`--op ADD_ROW` → 「その列のデータ行を全部書き換えます」＝ ADD_ROW には嘘）。

★ verdict は台帳に**書かない** ── 歩いた結果から決める（名簿を写して名簿と比べない）。
★ ここは LLM を回さない群（A 群）だけを歩く。翻訳を固定して引き金を引くので
  **素の環境でも毎回走る**（揺れが混ざらないので判定が安定する）。
  B 群（実機が要る 10 件）は -m local 側の仕事。

★ なぜ tests/ に在るか: 素の環境の番人は scripts/ 同士の import を弾く（3 度踏んだ）。
"""
from __future__ import annotations

import argparse
import contextlib
import io
import json
import os
import re
import shlex
import shutil
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _home_isolation import is_isolated, isolated_home  # noqa: E402

#: ★ 本物のホーム（ここに書いたら事故）。
REAL_HOME = Path.home() / ".ailine"

REGISTER = REPO / "tests" / "refusal_register.json"

#: 歩く時に使う冊（フォルダ経路なので 2 枚置く）。★ 固定する ── 冊が変わると引き金が動く。
BOOKS = {"4月.xlsx": [["商品", "金額"], ["ボルト", 120], ["ナット", 80]],
         "5月.xlsx": [["商品", "金額"], ["ボルト", 150], ["ワッシャー", 300]]}

#: 照合（2 冊）の検体。★★ この経路は **LLM を 1 語も呼ばない**（列対応は機械 3 段:
#:   依頼文の名指し → 型 → 曖昧なら exit 3）。だから揺れが無く **A 群で歩ける** ──
#:   「2 冊と実表が要るから A 群では歩けない」という 2026-09-19 の見立ては外れだった
#:   （2026-09-20 に実際に歩いて確かめた）。見立てで `未調査` に置くと、そこで止まる。
MATCH_BOOKS = {
    # ★ キーの候補が 2 つ（品名・備考）で、依頼文がどちらも名指ししていない
    "ambiguous_key": {
        "A": (["品名", "備考", "金額"], [["ボルト", "至急", 120], ["ナット", "", 80]]),
        "B": (["品名", "備考", "金額"], [["ボルト", "", 150], ["ワッシャー", "", 300]]),
    },
    # ★★ 2026-09-20: 「金額が式のまま」の検体は**もう断りを出さない** ── 道具が自分で
    #   LibreOffice に開かせて値を入れるようになったため（②「前提を道具が満たす」）。
    #   盤が「引き金が引けない（検体が古い）」で掴んだ ── 製品が良くなった側の古さ。
    #   ★ 断りが残るのは「使える列が**本当に無い**」回なので、検体をそちらへ移す。
    "no_usable_amount": {
        "A": (["品名", "金額"], [["ボルト", "要確認"], ["ナット", "未定"]]),
        "B": (["品名", "金額"], [["ボルト", 150], ["ワッシャー", 300]]),
    },
}

#: 同じ見出しのシートを 2 枚持つ冊（対象シートが決まらない断りの引き金）。
TWO_SHEETS = {"4月": [["商品", "金額"], ["ボルト", 120], ["ナット", 80]],
              "5月": [["商品", "金額"], ["ボルト", 150], ["ワッシャー", 300]]}


def _make_book(root: Path, sheets: dict | None = None) -> Path:
    """1 冊の冊を作る（sheets を渡せば複数シート）。"""
    root.mkdir(parents=True, exist_ok=True)
    p = root / "在庫.xlsx"
    wb = openpyxl.Workbook()
    if sheets:
        wb.remove(wb.active)
        for name, rows in sheets.items():
            ws = wb.create_sheet(name)
            for r in rows:
                ws.append(r)
    else:
        ws = wb.active
        ws.title = "在庫"
        for r in BOOKS["4月.xlsx"]:
            ws.append(r)
    wb.save(p)
    wb.close()
    return p


def _make_match_books(root: Path, which: str) -> tuple:
    """照合の検体 2 冊を作る（★ 式の列は文字列でなく**式として**書く）。"""
    spec = MATCH_BOOKS[which]
    out = []
    root.mkdir(parents=True, exist_ok=True)
    for side in ("A", "B"):
        headers, rows = spec[side]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "表"
        ws.append(headers)
        for r in rows:
            ws.append(r)
        p = root / f"{side}.xlsx"
        wb.save(p)
        wb.close()
        out.append(p)
    return tuple(out)


def load_register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def _make_folder(root: Path) -> Path:
    d = root / "月次"
    d.mkdir(parents=True, exist_ok=True)
    for name, rows in BOOKS.items():
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in rows:
            ws.append(r)
        wb.save(d / name)
        wb.close()
    return d


def _run(argv: list, plan, second=None) -> tuple:
    """製品を 1 回走らせ、(exit, 画面) を返す。plan を渡せば翻訳をそれに固定する。

    ★ second を渡すと **2 回目の読みだけ別の計画**にする（読みの割れを再現するため）。
    """
    # ★★ 2026-09-23: 切り離されていない状態で走らせない（呼ぶ側が忘れても本物に書けない）。
    #   実害: pytest の外から歩き手を呼んだサブエージェントが、本物の ~/.ailine に 103 行書いた。
    if not is_isolated(ailine, REAL_HOME):
        raise RuntimeError("ailine の保存先が本物の ~/.ailine を向いたまま ── walk_hint / walk_one の中"
                           "（isolated_home）から呼ぶこと")
    buf = io.StringIO()
    # ★ 2026-09-23: 歩き手は確認の問い（y/N）に答えない ── 標準入力を空にする（EOF ＝ 答えない）。
    #   pytest の中では標準入力が塞がっていて、読みに行くと OSError で落ちていた（外では EOF で進む）。
    real_stdin = sys.stdin
    sys.stdin = io.StringIO("")
    real, real_fixed = ailine.translate_task, ailine.translate_task_fixed_op
    # ★ 2026-10-02: 文書の正規化（LibreOffice で開いて保存し直す）は素通しにする ── conftest が関数ごとにやっている
    #   ことと同じ（ここは module の fixture から呼ばれ、その切り替えの外だった）。歩いて確かめたいのは案内であって
    #   正規化ではなく、案内を 200 本歩く間に LibreOffice を毎回起こすと 3 分近くかかった。実機の番人は本物で歩く。
    real_norm = ailine.normalize_book
    # ★ 2026-10-02: ollama も同じ ── conftest の _no_real_ollama は関数ごとの fixture で、module の fixture から呼ばれる
    #   ここには届かない。届かないと素の環境（ollama が居ない）で、語彙外の近い候補を訊く判定器（judge_ops_via_llm）が
    #   実 ollama に当たり、接続拒否が exit 9 で画面を埋めて案内が出なかった（手元は ollama が答えるので緑）。
    #   判定器は「候補なし」・自由生成は「空」に固定し（翻訳を固定するのと同じ作法）、残りの口は塞ぐ ── 漏れは歩きの失敗として手元でも赤になる。
    real_ollama = (ailine.judge_ops_via_llm, ailine.ollama_generate_json, ailine.ollama_generate)
    if not os.environ.get("AILINE_WALK_ON_MACHINE"):
        ailine.normalize_book = lambda book, workdir, timeout=None: book
        def _no_ollama(*a, **k):
            raise AssertionError("歩き手が実 ollama を呼んだ（固定していない口 ── 素の環境には存在しない）")
        ailine.judge_ops_via_llm = lambda task, about=None: []
        ailine.ollama_generate_json = _no_ollama
        #   自由生成（語彙外段）は「何も返さない」に固定 ── 歩きたいのは生成の前に出る通知で、生成物ではない。
        ailine.ollama_generate = lambda *a, **k: ""
    calls = []
    if plan is not None:
        def fake(*a, **k):
            calls.append(1)
            return {"plan": second if (second and len(calls) == 2) else plan}
        ailine.translate_task = fake
        # ★ 2026-10-02: 引数を渡すのは**固定する op が計画の先頭の op と同じ時だけ**。旧版は別の op にも先頭の引数を
        #   そのまま返した ── 製品は「抽出の依頼を 1 セル書換に読み直せるか」を第二段（op 固定の翻訳）の値で確かめるので、
        #   抽出の値（営業）が 1 セル書換の値として返り、**抽出の例が 1 セル書換に化けた**（治具の側の嘘・製品は正しい）。
        first_op = str((plan[0] or {}).get("op") or "") if plan else ""
        ailine.translate_task_fixed_op = lambda model, op, task, meta, **k: {
            "op": op, "args": dict((plan[0] or {}).get("args") or {}) if op == first_op else {}}
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                rc = ailine.main(argv)
            except SystemExit as e:
                rc = e.code
    finally:
        ailine.translate_task, ailine.translate_task_fixed_op = real, real_fixed
        ailine.normalize_book = real_norm
        ailine.judge_ops_via_llm, ailine.ollama_generate_json, ailine.ollama_generate = real_ollama
        sys.stdin = real_stdin
    return rc, buf.getvalue()


#: ★ 製品が「ここに通る書き方を並べる」と宣言している行の頭（`examples.py` 側の文面）。
_EXAMPLE_LINES = ("そのまま打てます:", "はこう頼めます:", "（例:", "（例：")


def _example_in(text: str) -> str | None:
    """断り文が見せている「通る書き方」を取り出す。

    ★★ 2026-09-20 の直し: 旧版は**画面じゅうの「…」を語で当てて**拾っていた。
      だから製品が例として掲げた行が増えると、拾う文が黙って別のものに変わる
      （盤の判定が製品の文面の並びに依存する ── 番人を字面で書いた形）。
    ★ いまは**例として掲げた行**を先に見つけ、その行の先頭の「…」を取る。
      掲げた行が無い回だけ、昔のやり方に落ちる（後方互換）。
    """
    import re
    for line in text.splitlines():
        if any(h in line for h in _EXAMPLE_LINES):
            m = re.search(r"「([^」]{4,60})」", line)
            if m:
                return m.group(1)
    for m in re.finditer(r"[『「]([^』」]{6,60})[』」]", text):
        s = m.group(1)
        if any(w in s for w in ("抜き出", "並べ替え", "追加", "にして", "消して", "入れて")):
            return s
    return None


def _trigger(w: dict, root: Path) -> tuple:
    """引き金を引く ── (exit, 画面, 引数の組み立てに使う土台)。"""
    kind = w.get("kind")
    extra = list(w.get("argv_extra") or [])
    if kind == "folder":
        folder = _make_folder(root)
        argv = ["run", str(folder), w["task"], "--out", str(root / "結果.xlsx")]
        return (*_run(argv, w.get("plan"), w.get("second")), folder)
    if kind == "match":
        # ★ 翻訳を差し替えない（plan=None）── この経路は LLM を呼ばないので、
        #   差し替えると「呼ばれていない」ことを隠してしまう。
        a_path, b_path = _make_match_books(root, w["books"])
        return (*_run(["run", str(a_path), str(b_path), w["task"]], None), (a_path, b_path))
    if kind == "two_sheets":
        book = _make_book(root, TWO_SHEETS)
        return (*_run(["run", str(book), w["task"], "--copy"] + extra,
                      w.get("plan"), w.get("second")), book)
    book = _make_book(root)
    return (*_run(["run", str(book), w["task"], "--copy"] + extra,
                  w.get("plan"), w.get("second")), book)


def _walk_path(w: dict, path: dict, base, root: Path) -> tuple:
    """示された道を歩く ── (exit, 画面)。歩けない種類なら (None, 理由)。"""
    kind, extra = path.get("kind"), list(path.get("argv_extra") or [])
    plan = path.get("plan") or w.get("plan")
    if kind == "single_book":
        one = base / next(iter(BOOKS))
        return _run(["run", str(one), w["task"], "--copy"], plan)
    if kind == "example":
        task2 = path.get("task") or _example_in(_LAST_SCREEN[0])
        if not task2:
            return None, "例を示していると数えられているが、歩ける文が取り出せない"
        if w.get("kind") == "folder":
            return _run(["run", str(base), task2, "--out", str(root / "結果2.xlsx")], plan)
        return _run(["run", str(base), task2, "--copy"] + extra, plan)
    if kind == "retry_with_task":
        # ★ 断りが示した通りに**言い直して**もう一度打つ（照合の道はこれ）。
        a_path, b_path = base
        return _run(["run", str(a_path), str(b_path), path["task"]], None)
    if kind == "walked_on_the_real_machine":
        # ★★ 道そのものは実機が要る（LibreOffice で開いて保存する等）。盤は A 群なので
        #   ここでは歩かず、**歩いている試験を名指しする** ── 「見ていない」と
        #   「別の所で見た」を混ぜない。名指しが腐らないよう番人が突き合わせる。
        return None, f"実機側で歩く: {path['walked_by']}"
    if kind == "forced_op":
        return _run(["run", str(base), w["task"], "--copy", "--op", path["op"]], plan)
    if kind == "sheet":
        return _run(["run", str(base), w["task"], "--copy", "--sheet", path["sheet"]], plan)
    return None, f"歩き方が不明: {kind}"


#: ★ 例を断り文から拾う時に使う（直前の画面）。引数で回すと walk_one の形が崩れるので棚に置く。
_LAST_SCREEN = [""]


def walk_one(key: str, entry: dict, root: Path) -> dict:
    """1 件の断りを出させ、示された道を歩く。戻り値に verdict を入れる（台帳は読まない）。"""
    root.mkdir(parents=True, exist_ok=True)
    with isolated_home(ailine, root / "_ailine_home"), contextlib.chdir(root):
        return _walk_one(key, entry, root)


def _walk_one(key: str, entry: dict, root: Path) -> dict:
    w = entry.get("walk")
    if not w:
        return {"key": key, "verdict": "未記入", "detail": "walk 欄が無い"}
    if w.get("kind") == "match_unwalkable":
        # ★ 突き合わせは 2 冊と実表が要る ── この器では歩けないと**書いて残す**。
        #   「歩いていない」を walked と混ぜない（見ていないものを見たことにしない）。
        return {"key": key, "verdict": "未調査", "detail": w.get("why") or "この器では歩けない"}

    rc, out, base = _trigger(w, root)
    _LAST_SCREEN[0] = out
    if rc == BUSY:
        # ★★ 2026-09-20: 引き金すら引けていない ── 断りが出たのではなく機械が塞がっていた。
        #   rc != 0 で「断りが出た」と数えると、その先の判定が全部この上に乗る。
        return {"key": key, "verdict": "歩けなかった",
                "detail": "機械が塞がっていて引き金が引けない（exit 6）", "screen": out[-160:]}
    if rc == 0:
        return {"key": key, "verdict": "引き金が引けない",
                "detail": "断りが出ずに到達した（検体が古い）", "screen": out[-200:]}

    path = w.get("path") or {}
    if path.get("kind") == "none":
        # ★ 道を示していない ── 理由が粗いのか、本当に道が無いのかは人が仕分ける。
        return {"key": key, "verdict": "vague" if w.get("why") else "no_path",
                "detail": w.get("why") or "示す道が無い（機能・語彙が無い）",
                "screen": out.strip()[-160:]}

    rc2, out2 = _walk_path(w, path, base, root)
    if rc2 is None and path.get("kind") == "walked_on_the_real_machine":
        return {"key": key, "verdict": "実機で歩く", "detail": out2,
                "screen": out.strip()[-160:]}
    if rc2 is None:
        return {"key": key, "verdict": "vague", "detail": out2, "screen": out.strip()[-160:]}
    if rc2 == BUSY:
        return {"key": key, "verdict": "歩けなかった",
                "detail": "道の途中で機械が塞がっていた（exit 6）", "screen": out2.strip()[-160:]}
    if rc2 == 0:
        return {"key": key, "verdict": "walked",
                "detail": f"道を歩いて到達（{path.get('kind')}）"}
    return {"key": key, "verdict": "path_fails",
            "detail": f"道を歩いたが exit {rc2}", "screen": out2.strip()[-200:]}


def survey() -> list:
    reg = load_register()
    rows = []
    # ★ 片付けの失敗で判定を落とさない（導線の盤と同じ ── 2026-09-23 に素の環境で「使用中」で落ちた）。
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
        for key, entry in reg["refusals"].items():
            if str(entry.get("unlock", "")).startswith("無い"):
                rows.append({"key": key, "verdict": "by_design",
                             "detail": str(entry["unlock"])[:60]})
                continue
            if not entry.get("walk"):
                rows.append({"key": key, "verdict": "未記入", "detail": "walk 欄が無い"})
                continue
            rows.append(walk_one(key, entry, Path(td) / key.replace("#", "_")))
    return rows


#: ★ 製品が「機械が塞がっている」と言う終了コード（別の ailine が実行中）。
#:   ★★ 2026-09-20: ロックの**ファイルを見る**形は使えないと実装して確かめた
#:     （conftest が AILINE_HOME を差し替えるのでテストからは見えない）。
#:     ここは**製品自身の終了コード**を読む ── 走らせ方によらず同じ答えになる。
BUSY = 6

ORDER = ["path_fails", "vague", "no_path", "未調査", "歩けなかった", "未記入",
         "引き金が引けない", "実機で歩く", "walked", "by_design"]


def broken_paths(survey_fn=None) -> list:
    """★★ 通らない道を示している断り ── **2 回歩いて再現したものだけ**返す。

    ★ 2026-09-20（前夜に偽の赤を踏んだ）: 盤は道を歩く時に LibreOffice を使うので、
      実機が過負荷だと道が歩けず path_fails と誤判定する（その回だけ赤く、空いた環境では
      0 件だった）。機械ロックを見る形は使えないと実装して確かめた ── conftest が
      AILINE_HOME をテストごとに差し替えるのでテストからロックが見えず、しかもロックを
      握られていても盤は歩けた（ロックは作法の取り決めで soffice を排他していない）。
    ★ だから 1 回の失敗では断定しない。判定はここ 1 箇所に置く（番人が書き写さない）。
    """
    run = survey_fn or survey
    first = [r for r in run() if r["verdict"] == "path_fails"]
    if not first:
        return []
    again = {r["key"] for r in run() if r["verdict"] == "path_fails"}
    return [r for r in first if r["key"] in again]


def render(rows: list) -> str:
    from collections import Counter
    c = Counter(r["verdict"] for r in rows)
    lines = ["導通の盤 ── 断りが示した道を歩く", ""]
    for v in ORDER:
        if c.get(v):
            lines.append(f"  {v:<16} {c[v]:>3}")
    lines.append(f"  {'合計':<16} {len(rows):>3}")
    lines.append("")
    for v in ORDER:
        hits = [r for r in rows if r["verdict"] == v and v != "by_design"]
        if not hits:
            continue
        lines.append(f"■ {v}")
        for r in hits:
            lines.append(f"    {r['key']:<28} {r['detail']}")
            if r.get("screen"):
                lines.append(f"        画面: {r['screen'][:110]}")
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 導線の台帳（typable_hints_register.json）を歩く ── 2026-09-23
# ---------------------------------------------------------------------------
# ★★ なぜ在るか: 「こう打てば進める」と人に言っている所が 40 件在り、`walked` は 38 件が
#   `未調査` のまま（人が手で書く欄だった）。盲検 6 体目の致命 #6（案内どおりに打つと落ちる）は
#   この家系で、歩く道具は**既に在った** ── 足りなかったのは台帳への配線だった。
# ★ 断りの台帳（上）の歩き方は `run` の引き金しか知らない。導線は accounts / csv / verify /
#   doctor / vocab … と広いので、**任意のサブコマンドと引数**で引き金を引き、示した道も歩く。
# ★★ 上より 1 つ厳しい: 上は「exit が 0 でない＝断りが出た」と数える。それだと**別の理由で
#   落ちた回も案内が出たことになる**。ここは `expect`（案内の文面の一部）が**画面に実際に
#   出たか**まで確かめる。出ていなければ「引き金が引けない」（直すのは歩き方の側）。
#
# 歩き方（台帳の各 hint の `walk`）:
#   books   {"名.xlsx": {"シート名": [[行], ...]}}  ── root に作る冊（任意）
#   texts   {"名.csv": "中身"}                       ── root に置くテキスト（本物の CSV 等・UTF-8）
#   marks   {"名.xlsx": {"creator": "ailine split", "description": "…"}} ── 冊に焼く印
#   dirs    ["空のフォルダ", ...]                      ── root に作る空のフォルダ
#   setup   [{"argv": [...], "plan": [...]}, ...]    ── 引き金の前に製品を走らせて本物の出力を作る
#           （★ どれか 1 つでも exit 0 でなければ「引き金が引けない」── 検体づくりの失敗を判定に混ぜない）
#   argv    ["accounts", "{book:名.xlsx}", ...]     ── 引き金。{book:名} は作った冊のパス、{root} は作業場
#   plan    [...]                                     ── 翻訳を固定する（任意・LLM を回さない）
#   expect  "その列でよければ"                         ── ★ 必須。案内が出たことの証拠
#   path    {"argv": [...], "plan": [...], "expect_rc": 0, "expect": "…", "follow": "--column"}
#           ── follow を書くと、引き金の画面の `expect` より後ろに出た `` `--column …` `` を
#              **画面から拾って** argv の後ろに足す（★ 正解を手で書かない ── 案内どおりに打つ）
#           ── 示した道。expect_rc（既定 0）で着いたとみなす。expect を書けばその文面も要る
#   machine true ── その場の機械の状態（ollama・LibreOffice・doctor の点検・demo）に左右される歩き。
#           素の環境（CI）では歩かず「実機で歩く」と数え、実機の番人（-m local）が AILINE_WALK_ON_MACHINE=1 で歩く
#   ── 2026-10-02 に足した口（どれも「歩き方に正解を手で書かない」「機械の状態を名前で決める」ため）──
#   env     {"COLUMNS": "400"}                         ── 歩きの間の環境変数（argparse の -h が折り返さないように）
#   patch   {"doctor_missing": [...] | "doctor_ok": true | "struct_dump_missing": true | "fidelity_lost": true}
#           ── 引き金の間だけ、製品の外の状態を名前で決め打ちにする（path.patch は道の間だけ）。_patched を見る
#   hold_lock true ── 引き金の間だけ、別のプロセスが実行ロックを持っている形にする（道は離してから歩く）
#   second  [...]                                      ── 2 回目の読みだけ別の計画（読みの割れの検体）
#   setup の {"edit_cell": {...}} / {"copy": [元, 先]}  ── 製品が作った冊を人が直した形にする／案内が名指しする名前を付ける
#   path    の follow_* ── 画面から拾って打つ（正解を歩き方に手で書かない）:
#           follow_command    "ailine undo \"" 等（リストなら全部）── 打てと言われた 1 行そのもの（Windows の打ち方で割る）
#           follow_task       {"at": argv の位置, "after": "例:", "nth": 1, "next": 0} ── 「（例:「…」）」の例を依頼文にして打つ（入れ子の括弧を読む）
#           follow_option     旗の番号 / その旗の説明の文 ── 「以下のいずれかを指定して」の旗
#           follow_choice     何番目か ── 画面の『候補: <op>』を --op に渡す
#   path    の remove（先に消すパス）/ then（案内が「先に A、それから B」と言う時の B。リストなら順に）/ second / patch
#   {sha256:名} ── 作業場にある冊の sha256（--base-sha に渡す値）
#   walk.not_guidance true（kind が by_design の時）── 走査が拾ったが、次に打つものを指していなかった（理由は why）
#   kind    "by_design"（意図した行き止まり・why 必須）/
#           "walked_on_the_real_machine"（walked_by に歩いている試験名）── 歩かない種類

HINTS_REGISTER = REPO / "tests" / "typable_hints_register.json"


def load_hints() -> list:
    """歩く対象 ── hints（バッククォートの打てるもの）と guidance（それ以外の案内・2026-10-01）の両方。

    ★ guidance を別の歩き手にしない ── 同じ盤で歩けば、walked 欄を縛る番人 3 本がそのまま両方に効く。
    """
    reg = json.loads(HINTS_REGISTER.read_bytes().decode("utf-8"))
    return list(reg["hints"]) + list(reg.get("guidance") or [])


def hint_key(h: dict) -> str:
    """★ 2026-09-23: 鍵は『ファイル＋文面』（行番号は上に 1 行足すだけで全部ずれた ── 同じ日に 2 回）。
    文面が変われば歩き直すべき時なので、鳴ってほしい時にだけ鳴る。"""
    return f"{h['file']}:{h['text']}" + (f"#{h['nth']}" if h.get("nth", 1) != 1 else "")


def _prepare(root: Path, w: dict) -> str | None:
    """冊・テキスト・印・空のフォルダを置き、setup を走らせる。失敗したら理由を返す。"""
    _make_books(root, w.get("books") or {})
    for name, text in (w.get("texts") or {}).items():
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_bytes(str(text).encode("utf-8"))
    for name, props in (w.get("marks") or {}).items():
        wb = openpyxl.load_workbook(root / name)
        for k, v in props.items():
            setattr(wb.properties, k, v)
        wb.save(root / name)
        wb.close()
    for d in w.get("dirs") or []:
        (root / d).mkdir(parents=True, exist_ok=True)
    for i, step in enumerate(w.get("setup") or []):
        if step.get("edit_cell"):
            # ★ 2026-10-02: 製品が作った冊を、人が手で直した形にする（「そのあと変更されています」の検体）
            e = step["edit_cell"]
            wb = openpyxl.load_workbook(root / e["book"])
            wb[e["sheet"]][e["cell"]] = e["value"]
            wb.save(root / e["book"])
            wb.close()
            continue
        if step.get("copy"):
            # ★ 2026-10-02: 製品が作った出力に、案内が名指しする名前を付ける（「照合.xlsx」等。名前は検体の都合）
            src, dst = _subst(step["copy"], root)
            shutil.copy2(src, dst)
            continue
        rc, out = _run(_subst(step["argv"], root), step.get("plan"))
        if rc != 0:
            return f"setup {i + 1} が exit {rc}（検体を作れていない）: {out.strip()[-120:]}"
    return None


def _make_books(root: Path, books: dict) -> None:
    root.mkdir(parents=True, exist_ok=True)
    for name, sheets in (books or {}).items():
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sheet, rows in sheets.items():
            ws = wb.create_sheet(sheet)
            for r in rows:
                ws.append(r)
        # ★ 2026-10-02: 「配る/人の資料.xlsx」のようにフォルダの中へ置ける（出力先の関所の検体）
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        wb.save(root / name)
        wb.close()


def _subst(args: list, root: Path) -> list:
    out = []
    for a in args:
        a = str(a).replace("{root}", str(root))
        while "{book:" in a:
            i = a.index("{book:")
            j = a.index("}", i)
            a = a[:i] + str(root / a[i + 6:j]) + a[j + 1:]
        while "{sha256:" in a:
            # ★ 2026-10-02: その時の冊の指紋（`--base-sha` に渡す値 ── 歩き方に値を手で書かない）
            i = a.index("{sha256:")
            j = a.index("}", i)
            import hashlib
            a = a[:i] + hashlib.sha256((root / a[i + 8:j]).read_bytes()).hexdigest() + a[j + 1:]
        out.append(a)
    return out


def walk_hint(h: dict, root: Path) -> dict:
    """1 件の導線を出させ、示した道を歩く。★ 台帳の walked は読まない（歩いた結果から決める）。
    ★ 引き金から示した道まで同じ切り離しの中で走らせる（undo は前の実行の控えを使う）。"""
    # ★ 今いるフォルダも作業場へ移す ── 2026-09-23 に歩きの中の `ailine demo` が repo の根っこへ
    #   見本の冊を 5 つ書いた（保存先と同じく、呼ぶ側が忘れても汚さない）。
    root.mkdir(parents=True, exist_ok=True)
    env = (h.get("walk") or {}).get("env")
    with isolated_home(ailine, root / "_ailine_home"), contextlib.chdir(root), _env(env):
        return _walk_hint(h, root)


@contextlib.contextmanager
def _patched(spec: dict | None):
    """歩きの間だけ製品の外の状態を決め打ちにする（★ 2026-10-02・名前で宣言できるものだけ）。
      doctor_missing: ["LibreOffice", …]  doctor の点検を、この名前の項目だけ「足りない」にする
                      （★ 機械の状態に左右されない ── demo の「先に足りないものがあります」の検体）
      doctor_ok: true                     doctor の点検を全部「在る」にする（demo の「次にこれを打ってみてください」の検体）
      struct_dump_missing: true           構造の読み取りが取れなかった形にする（LibreOffice の一時不調の検体）
      fidelity_lost: true                 往復の忠実度ゲートが「失われる」と言う形にする（実機の LibreOffice が要らない）"""
    spec = spec or {}
    real = (ailine.doctor_checks, ailine._struct_dump_info_missing, ailine.check_round_trip_fidelity)
    missing = list(spec.get("doctor_missing") or [])
    if missing:
        ailine.doctor_checks = lambda *a, **k: [(n, False, "足りない（検体）") for n in missing]
    elif spec.get("doctor_ok"):
        ailine.doctor_checks = lambda *a, **k: [("（検体）", True, "")]
    if spec.get("struct_dump_missing"):
        ailine._struct_dump_info_missing = lambda *a, **k: True
    if spec.get("fidelity_lost"):
        ailine.check_round_trip_fidelity = lambda *a, **k: {"lost": True, "items": [{"label": "図形", "count": 1}]}
    try:
        yield
    finally:
        ailine.doctor_checks, ailine._struct_dump_info_missing, ailine.check_round_trip_fidelity = real


@contextlib.contextmanager
def _env(extra: dict | None):
    """歩きの間だけ環境変数を足す（★ 2026-10-02: argparse の -h は端末の幅で折り返す ── COLUMNS を固定して
    文面の途中で割れないようにする）。終わったら元に戻す。"""
    saved = {k: os.environ.get(k) for k in (extra or {})}
    os.environ.update({k: str(v) for k, v in (extra or {}).items()})
    try:
        yield
    finally:
        for k, v in saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


@contextlib.contextmanager
def _held_run_lock(enabled):
    """★ 2026-10-02: 「別の ailine が実行中です」の検体 ── 別プロセスが持っている形を、**同じプロセスの別の fd** で
    OS の排他ロックを掛けて作る（製品は鍵を別の fd で開き直して掛けに行くので、掛けられず断る）。
    引き金を引く間だけ持ち、道を歩く前に必ず離す。"""
    if not enabled:
        yield
        return
    Path(ailine.RUN_LOCK_FILE).parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(str(ailine.RUN_LOCK_FILE), os.O_CREAT | os.O_RDWR)
    try:
        if os.fstat(fd).st_size == 0:
            os.write(fd, b" ")
        if not ailine._try_os_lock(fd):
            raise RuntimeError("検体を作れない: 実行ロックを先に掛けられなかった")
        yield
    finally:
        ailine._release_os_lock(fd)
        os.close(fd)


def _walk_hint(h: dict, root: Path) -> dict:
    key = hint_key(h)
    w = h.get("walk")
    if not w:
        return {"key": key, "verdict": "未記入", "detail": "walk 欄が無い"}
    if w.get("kind") == "by_design":
        return {"key": key, "verdict": "by_design", "detail": w.get("why") or "（why が無い）"}
    if w.get("kind") == "walked_on_the_real_machine":
        return {"key": key, "verdict": "実機で歩く", "detail": w.get("walked_by", "")}
    if not w.get("expect"):
        return {"key": key, "verdict": "引き金が引けない",
                "detail": "expect が無い ── 案内が出たことを確かめられない（歩き方の側を直す）"}
    if w.get("machine") and not os.environ.get("AILINE_WALK_ON_MACHINE"):
        # ★ 2026-09-23: 素の環境（pre-push の CI 相当）で demo・doctor・翻訳を固定しない run の 4 件が落ちた ──
        #   歩き方が悪いのでなく、機械の状態で答えが変わる歩き。判定に混ぜず、実機の番人で歩く。
        return {"key": key, "verdict": "実機で歩く",
                "detail": "機械の状態に左右される歩き ── test_typable_hints_walk_on_the_machine が歩く"}
    why = _prepare(root, w)
    if why:
        return {"key": key, "verdict": "引き金が引けない", "detail": why}
    with _held_run_lock(w.get("hold_lock")), _patched(w.get("patch")):     # ★ 決め打ちは引き金の間だけ（道は本物の状態で歩く）
        rc, out = _run(_subst(w["argv"], root), w.get("plan"), w.get("second"))
    if rc == BUSY and not w.get("hold_lock"):
        return {"key": key, "verdict": "歩けなかった", "detail": "機械が塞がっていた（exit 6）"}
    if w["expect"] not in out:
        return {"key": key, "verdict": "引き金が引けない",
                "detail": f"案内『{w['expect']}』が画面に出なかった（exit {rc}）",
                "screen": out.strip()[-200:]}
    path = w.get("path") or {}
    for victim in _subst(path.get("remove") or [], root):
        # ★ 2026-10-02: 「そのファイルを別の場所へ移すか削除してから、もう一度実行して」型 ── 言われたとおり先に消す
        Path(victim).unlink()
    if path.get("follow_command"):
        # ★★ 2026-10-02: 「例: ailine verify 縦積み.xlsx 受領フォルダ」型 ── 打てと言われた**1 行そのもの**を
        #   画面から拾って打つ（歩き方に手で書くと、案内が何と言っても通ってしまう）。
        #   冊は作業場に置いてあること（歩きは作業場で走る）。
        #   リストなら 1 つの案内が打てと言った**全部**を打つ（最後の 1 つは下の共通の歩きで打つ）。
        starts = path["follow_command"] if isinstance(path["follow_command"], list) else [path["follow_command"]]
        cmds = []
        for start in starts:
            cmd = _typed_command(out, start)
            if not cmd:
                return {"key": key, "verdict": "vague",
                        "detail": f"画面に『{start}』で始まる打てる 1 行が出ていない"}
            cmds.append(cmd)
        for cmd in cmds[:-1]:
            rc_i, out_i = _run(cmd, path.get("plan", w.get("plan")), path.get("second"))
            if rc_i != path.get("expect_rc", 0):
                return {"key": key, "verdict": "path_fails",
                        "detail": f"案内どおり `ailine {' '.join(cmd)}` と打ったが exit {rc_i}",
                        "screen": out_i.strip()[-200:]}
        path = {**path, "argv": cmds[-1]}
    if not path.get("argv"):
        return {"key": key, "verdict": "vague", "detail": "path.argv が無い ── 示した道を書いていない"}
    argv2 = _subst(path["argv"], root)
    if path.get("follow"):
        # ★★ 2026-09-23: 示した道を手で書くと、画面が何と言っても正解を打ってしまう（変異が素通りした）。
        #   案内の後ろのバッククォートから拾う ── 案内が 1 冊ぶんしか言わなければ、1 冊ぶんしか打たない。
        tail = out.split(w["expect"], 1)[1]
        flag = path["follow"]
        picked = re.findall(r"`(" + re.escape(flag) + r" [^`]+)`", tail)
        if not picked:
            return {"key": key, "verdict": "vague",
                    "detail": f"案内の後ろに `{flag} …` が出ていない ── 拾う道が無い"}
        for p in picked:
            argv2 += shlex.split(p)
    if path.get("follow_quoted") is not None:
        # ★★ 2026-10-01: 「依頼文に『A』『B』を書き足して」型の案内を、**画面から拾って**書き足す
        #   （follow と同じ理由 ── 歩き方に正解を手で書くと、案内が何と言っても通ってしまう）。
        line = next((ln for ln in out.splitlines() if w["expect"] in ln), "")
        quoted = re.findall(r"『([^』]+)』", line)
        if not quoted:
            return {"key": key, "verdict": "vague",
                    "detail": "案内の行に『…』が出ていない ── 書き足す言い方が無い"}
        i = path["follow_quoted"]
        argv2[i] = argv2[i] + "、" + "、".join(quoted)
    if path.get("follow_option") is not None:
        # ★★ 2026-10-02: 「以下のいずれかを指定して再実行してください」型 ── 画面に並んだ旗の n 番目を、そのまま足して打つ
        #   （どの旗を勧めるかを歩き方に手で書かない）。follow_option = 何番目か（1 起点）。
        #   follow_option が文字列なら「その文を説明にしている旗」（案内の文が旗の説明の側にある時）。
        fo = path["follow_option"]
        if isinstance(fo, str):
            flags_seen = [m.group(1) for m in (re.match(r"\s{2,}(--[a-z][a-z-]*)(?=\s)", ln)
                                               for ln in out.splitlines() if fo in ln) if m]
            fo = 1
        else:
            flags_seen = re.findall(r"^\s{2,}(--[a-z][a-z-]*)(?=\s)", out.split(w["expect"], 1)[1], flags=re.M)
        if len(flags_seen) < fo:
            return {"key": key, "verdict": "vague", "detail": "案内に旗の選択肢が並んでいない"}
        argv2 += [flags_seen[fo - 1]]
    if path.get("follow_choice") is not None:
        # ★★ 2026-10-02: 「片方を選ぶなら、その候補を --op で固定して」型 ── 画面の『候補: <op>』を拾って --op に渡す
        #   （正解の op を歩き方に手で書かない）。follow_choice = 何番目の候補か（1 起点）。
        ops_seen = re.findall(r"^候補: ([A-Z_]+)\t", out, flags=re.M)
        if len(ops_seen) < path["follow_choice"]:
            return {"key": key, "verdict": "vague", "detail": "画面に『候補: <op>』が出ていない ── 固定する候補が無い"}
        argv2 += ["--op", ops_seen[path["follow_choice"] - 1]]
    if path.get("follow_task") is not None:
        # ★★ 2026-10-02: 「（例:「…」）」型の案内を、**画面の例そのもの**を依頼文にして打つ（follow と同じ理由）。
        #   follow_task = {"at": argv の位置, "after": この語より後ろの最初の「…」（省略可）, "nth": 何番目か（省略可）}
        ft = path["follow_task"]
        screen = out.splitlines()
        at_line = next((i for i, ln in enumerate(screen) if w["expect"].splitlines()[0] in ln), None)
        # ★ next = 案内の行の何行あとに例が出るか（省略時は同じ行）。★ 入れ子の『…』は「…」の中身として読む
        line = screen[at_line + ft.get("next", 0)] if at_line is not None and at_line + ft.get("next", 0) < len(screen) else ""
        seg = line.split(ft["after"], 1)[1] if ft.get("after") and ft["after"] in line else line
        said = _quoted_examples(seg)
        if len(said) < ft.get("nth", 1):
            return {"key": key, "verdict": "vague", "detail": "案内の行に「…」の例が出ていない ── 打つ例が無い"}
        argv2[ft["at"]] = said[ft.get("nth", 1) - 1]
    with _patched(path.get("patch")):        # ★ 道の側の決め打ちは path.patch に別に書く（引き金のとは混ぜない）
        rc2, out2 = _run(argv2, path.get("plan", w.get("plan")), path.get("second"))
    if rc2 == BUSY:
        return {"key": key, "verdict": "歩けなかった", "detail": "道の途中で機械が塞がっていた（exit 6）"}
    want = path.get("expect_rc", 0)
    if rc2 == want and (not path.get("expect") or path["expect"] in out2):
        thens = path.get("then") or []
        for then in (thens if isinstance(thens, list) else [thens]):
            # ★ 2026-10-02: 案内が「先に A をして、それから B」と言う道（例: 用語集に登録してから、もう一度）── A を打った
            #   後に B まで歩いて初めて着いたと数える。then = {"argv", "plan"（省略時は上と同じ）, "expect_rc", "expect"}（並べれば順に）
            rc3, out3 = _run(_subst(then["argv"], root), then.get("plan", path.get("plan", w.get("plan"))), then.get("second"))
            want3 = then.get("expect_rc", 0)
            if rc3 != want3 or (then.get("expect") and then["expect"] not in out3):
                return {"key": key, "verdict": "path_fails",
                        "detail": f"案内の前半は通ったが、その後の実行が exit {rc3}（期待 {want3}）",
                        "screen": out3.strip()[-200:]}
        return {"key": key, "verdict": "walked", "detail": f"道を歩いて到達（exit {rc2}）"}
    return {"key": key, "verdict": "path_fails",
            "detail": f"道を歩いたが exit {rc2}（期待 {want}）", "screen": out2.strip()[-200:]}


def _quoted_examples(seg: str) -> list:
    """行の中の最も外側の「…」『…』の中身を、出てきた順に返す（★ 入れ子を読む: 「備考の列を全部「確認済」に書き換えて」
    は 1 つの例 ── 最初の閉じ括弧で切らない）。"""
    pairs = {"「": "」", "『": "』"}
    out, depth, start, closer = [], 0, 0, ""
    for i, ch in enumerate(seg):
        if depth == 0 and ch in pairs:
            depth, start, closer = 1, i + 1, pairs[ch]
        elif depth and ch == closer:
            depth -= 1
            if depth == 0:
                out.append(seg[start:i])
        elif depth and ch == {v: k for k, v in pairs.items()}[closer]:
            depth += 1
    return out


def _typed_command(out: str, start: str) -> list | None:
    """画面の中の『ailine …』で始まる 1 行（start を含むもの）を、打てる引数の列にして返す。
    ★ 行の終わりか、閉じ括弧・読点（）)」』、。）までを 1 つのコマンドと数える（':' は数えない ── パスのドライブ名）。
    ★ start は文脈つきでもよい（「世代の一覧は ailine undo」）── 打つのは、その後ろの最初の `ailine` から。"""
    for ln in out.splitlines():
        i = ln.find(start)
        if i < 0:
            continue
        j = ln.find("ailine", i)
        seg = re.split(r"[）)」』、。]", ln[j if j >= 0 else i:], 1)[0].strip()
        try:
            # ★ Windows の打ち方で割る（posix=False）── パスの \\ を畳まない。引用符は外す。
            toks = [t[1:-1] if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'" else t
                    for t in shlex.split(seg, posix=False)]
        except ValueError:
            return None
        return toks[1:] if toks and toks[0] == "ailine" else toks
    return None


def survey_hints() -> list:
    rows = []
    # ★ 片付けの失敗で判定を落とさない ── 2026-09-23 に export-csv の歩きの後で冊が「使用中」のまま
    #   残った（製品が冊を開いたまま離していない疑い・別件として残す）。
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
        for i, h in enumerate(load_hints()):
            rows.append(walk_hint(h, Path(td) / f"h{i:02d}"))
    return rows


def broken_hint_paths(survey_fn=None) -> list:
    """★ 通らない道を示している導線 ── 断りの盤と同じく **2 回歩いて再現したものだけ**。"""
    return broken_paths(survey_fn or survey_hints)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="断りが示した道を歩く（導通試験）")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--hints", action="store_true",
                    help="導線の台帳（typable_hints_register.json）を歩く")
    a = ap.parse_args(argv)
    rows = survey_hints() if a.hints else survey()
    print(json.dumps(rows, ensure_ascii=False, indent=1) if a.json else render(rows))
    # ★ path_fails が在る回だけ落とす ── vague/no_path は在庫であって赤ではない。
    #   ★ 判定は broken_paths 1 箇所（再現したものだけ）。ここで書き写さない。
    return 1 if broken_paths(lambda: rows) else 0


if __name__ == "__main__":
    raise SystemExit(main())
