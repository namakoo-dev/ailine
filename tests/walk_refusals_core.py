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
import re
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
    calls = []
    if plan is not None:
        def fake(*a, **k):
            calls.append(1)
            return {"plan": second if (second and len(calls) == 2) else plan}
        ailine.translate_task = fake
        ailine.translate_task_fixed_op = lambda model, op, task, meta, **k: {
            "op": op, "args": dict((plan[0] or {}).get("args") or {})}
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                rc = ailine.main(argv)
            except SystemExit as e:
                rc = e.code
    finally:
        ailine.translate_task, ailine.translate_task_fixed_op = real, real_fixed
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
    with tempfile.TemporaryDirectory() as td:
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
#   kind    "by_design"（意図した行き止まり・why 必須）/
#           "walked_on_the_real_machine"（walked_by に歩いている試験名）── 歩かない種類

HINTS_REGISTER = REPO / "tests" / "typable_hints_register.json"


def load_hints() -> list:
    return json.loads(HINTS_REGISTER.read_bytes().decode("utf-8"))["hints"]


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
        out.append(a)
    return out


def walk_hint(h: dict, root: Path) -> dict:
    """1 件の導線を出させ、示した道を歩く。★ 台帳の walked は読まない（歩いた結果から決める）。
    ★ 引き金から示した道まで同じ切り離しの中で走らせる（undo は前の実行の控えを使う）。"""
    # ★ 今いるフォルダも作業場へ移す ── 2026-09-23 に歩きの中の `ailine demo` が repo の根っこへ
    #   見本の冊を 5 つ書いた（保存先と同じく、呼ぶ側が忘れても汚さない）。
    root.mkdir(parents=True, exist_ok=True)
    with isolated_home(ailine, root / "_ailine_home"), contextlib.chdir(root):
        return _walk_hint(h, root)


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
    why = _prepare(root, w)
    if why:
        return {"key": key, "verdict": "引き金が引けない", "detail": why}
    rc, out = _run(_subst(w["argv"], root), w.get("plan"))
    if rc == BUSY:
        return {"key": key, "verdict": "歩けなかった", "detail": "機械が塞がっていた（exit 6）"}
    if w["expect"] not in out:
        return {"key": key, "verdict": "引き金が引けない",
                "detail": f"案内『{w['expect']}』が画面に出なかった（exit {rc}）",
                "screen": out.strip()[-200:]}
    path = w.get("path") or {}
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
            argv2 += p.split(" ", 1)
    rc2, out2 = _run(argv2, path.get("plan", w.get("plan")))
    if rc2 == BUSY:
        return {"key": key, "verdict": "歩けなかった", "detail": "道の途中で機械が塞がっていた（exit 6）"}
    want = path.get("expect_rc", 0)
    if rc2 == want and (not path.get("expect") or path["expect"] in out2):
        return {"key": key, "verdict": "walked", "detail": f"道を歩いて到達（exit {rc2}）"}
    return {"key": key, "verdict": "path_fails",
            "detail": f"道を歩いたが exit {rc2}（期待 {want}）", "screen": out2.strip()[-200:]}


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
