"""適用の前後を見る助言は、入口が **1 本**であること（2026-09-06）。

★★ なぜ（この日の朝、足す前に測って分かった）: 「値の書き込みが式を潰しても誰も
  言わない」を直そうとして呼び出し側を見たら、こう書いてあった ──

      compose_dsl_step_advisories(...) + formula_error_advisory(...) + broken_identity_advisory(...)

  ★ **同じ並びが 4 箇所**に写経されていた（単発・確認つき・上書き・複合計画の段）。
    3 本目をそのまま足せば **4 箇所 × 3 本**になる ── 前日 1 日かけて潰した
    「二重化した経路は片配線が既定で起きる」を、自分で 1 世代増やすところだった。

★ だから順番を守った: **畳む（挙動不変）→ 番人 → 足す**。
  この番人が守るのは「畳んだ形が戻らないこと」。次に前後を見る助言を足したい人が
  呼び出し側に `+ 新しい助言(...)` と書き足せば、ここが赤くなる。

★ 分母の作り方（入口側から数える）: 「適用の前後を受け取る助言」を名簿で持ち、
  呼び出し側にそれが**直接**現れていないことを見る。名前を列挙するのではなく、
  **合流点を 1 つに保つ**ことを縛っている。
"""
import inspect
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core import formula_health  # noqa: E402

SRC = inspect.getsource(ailine)
NL = chr(10)

def _advisories_behind_the_door() -> tuple:
    """扉（`before_after_advisories`）が**実際に呼んでいる**助言の名前を機械で引く。

    ★★ 2026-09-06（自作 review が拾った）: ここは手書きの名簿だった
      （「★ 足したらここにも 1 行」）。案の定 **3 本目（formula_loss_advisory）が漏れ**、
      その助言だけ「呼び出し側が直接呼んでも捕まらない」状態になっていた。
      ★ この repo の原則は「索引は手書きしない」── 名簿を実装から引く形に替える。
    ★★ 助言は **2 通りの住み方**をする。片方だけ引くと名簿が痩せる:
      ① `formula_health` の中で直接呼ばれるもの（formula_error / formula_loss）
      ② `ailine.py` 側に住み、扉へ **注入**されるもの（`identity_advisory=...`）
      ★ 実際 2026-09-06 に①だけを引く版を書いて、手書き名簿に在った
        `broken_identity_advisory` を**落とした**（しかも試験は緑のままだった）──
        機械化したら被覆が減る、という形。だから下の**下限**を必ず併せて置く。
    """
    import re
    src = inspect.getsource(formula_health.before_after_advisories)
    direct = {n for n in re.findall(r"([A-Za-z_]+_advisory)\s*\(", src)
              if hasattr(formula_health, n)}
    injected = {n for n in re.findall(r"identity_advisory=([A-Za-z_]+)", SRC)
                if n != "None" and hasattr(ailine, n)}
    return tuple(sorted(direct | injected))


#: 適用の前後を受け取る助言（★ 実装から機械で引く ── 手書きしない）
BEFORE_AFTER_ADVISORIES = _advisories_behind_the_door()

#: 合流点（呼び出し側が通ってよい唯一の入口）
DOOR = "before_after_advisories"


def test_the_door_exists_and_is_used():
    assert hasattr(formula_health, DOOR), f"{DOOR} が無い"
    assert SRC.count(DOOR + "(") >= 4, (
        f"{DOOR} を通る呼び出しが {SRC.count(DOOR + '(')} 箇所しかない"
        "（単発・確認つき・上書き・複合計画の 4 経路が在るはず）")


def test_no_caller_assembles_them_by_hand():
    """★ 本命: 呼び出し側が助言を**手で並べて**いないこと。"""
    door_body = SRC.split("def " + DOOR)[0]   # 定義そのものは対象外（下で別に見る）
    del door_body
    offenders = []
    for name in BEFORE_AFTER_ADVISORIES:
        for m in re.finditer(re.escape(name) + r"\(", SRC):
            line_no = SRC.count(NL, 0, m.start()) + 1
            line = SRC.splitlines()[line_no - 1]
            if line.lstrip().startswith(("from ", "import ", "def ")):
                continue          # import と定義は入口ではない
            offenders.append((line_no, line.strip()[:70]))
    assert not offenders, (
        f"前後を見る助言を直接呼んでいる箇所がある: {offenders} ── "
        f"{DOOR} の中に足すこと（呼び出し側に持たせない）")


def test_the_door_actually_calls_them_all():
    """★ 逆向き: 名簿に在るものが、入口の中で本当に呼ばれていること。

    ★ 2026-09-06: 初版は for の中でしか assert しておらず、**名簿が空なら 1 回も
      回らずに緑**だった（`tests/test_guard_ledger.py` が捕まえた ── repo が
      「番人の書き方」自体を縛っている）。分母を先に確かめる。
    """
    assert BEFORE_AFTER_ADVISORIES, "名簿が空（この試験は 1 度も回らない）"
    # ★ **下限**（2026-09-06 に引き方が狭まって 3 → 2 に痩せたのを、試験が見逃した）。
    #   減ったら「助言を消した」か「引き方が狭まった」── どちらも人が見るべき事件。
    assert len(BEFORE_AFTER_ADVISORIES) >= 3, (
        f"名簿が痩せた: {BEFORE_AFTER_ADVISORIES}（引き方が狭まっていないか）")
    body = inspect.getsource(getattr(formula_health, DOOR))
    missing = [n for n in BEFORE_AFTER_ADVISORIES
               if n not in body and "identity_advisory" not in body]
    assert not missing, f"入口を通っていない助言: {missing}"


def test_the_door_takes_the_before_and_the_after():
    """★ 「前後を見る」の中身 ── 適用前と適用後の両方を受け取ること。"""
    params = list(inspect.signature(getattr(formula_health, DOOR)).parameters)
    assert params[:2] == ["before_path", "after_path"], params


def test_the_door_actually_passes_both_through(tmp_path):
    """★ **通っていること**と**効いていること**は別（変異試験がすり抜けて分かった）。

    入口の中身を空にしても、字面を見る番人は全部緑だった ── 実際に前後 2 冊を渡して
    両方の助言が出ることを確かめる。
    ★ 検体は「式が #DIV/0! に壊れた」形（formula_error_advisory の担当）と、
      呼ばれたことを記録する偽の identity_advisory の 2 本。
    """
    import openpyxl
    before = tmp_path / "b.xlsx"
    after = tmp_path / "a.xlsx"
    for path, bad in ((before, False), (after, True)):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["A", "B"])
        ws.append([1, "#DIV/0!" if bad else 2])
        wb.save(path)
    seen = []

    def fake_identity(b, a, resolved):
        seen.append((b, a, resolved))
        return ["⚠ 等式が崩れました（偽）"]

    got = formula_health.before_after_advisories(
        before, after, {"col": "B"}, cell_ref=lambda r, c: f"R{r}C{c}",
        identity_advisory=fake_identity)
    assert any("エラー値のセルが増え" in g for g in got), got
    assert any("等式が崩れました" in g for g in got), got
    assert seen and seen[0][2] == {"col": "B"}, seen

    # ★ 宣言が無い経路（自由生成）: 等式は呼ばれず、エラー値の方だけ出る
    only = formula_health.before_after_advisories(
        before, after, None, cell_ref=lambda r, c: f"R{r}C{c}")
    assert any("エラー値のセルが増え" in g for g in only), only
    assert not any("等式" in g for g in only), only


def test_the_door_does_not_import_ailine():
    """★ ailine_core の作法（移植可能性）── 本体の関数は引数で受ける。"""
    text = (REPO / "src" / "ailine_core" / "formula_health.py").read_text(encoding="utf-8")
    assert "import ailine" not in text, "ailine_core が本体を import している"
    body = inspect.getsource(getattr(formula_health, DOOR))
    assert "cell_ref" in body and "identity_advisory" in body, (
        "本体側の関数を引数で受けていない（読み込み時に固定すると monkeypatch が効かない）")
