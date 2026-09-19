# -*- coding: utf-8 -*-
"""同じ依頼を 2 回読んで読み方が分かれたら、実行しない（2026-09-19）。

★★ 出所（Namakoo「揺れが絡んできたのは厄介だな」→「揺れだと判定できるかどうかも重要だ」
  →「応答速度より信頼性を取りたい」→「1で設計して実装してほしい」）:

  HEAD を固定して同じ依頼を 40 回翻訳した実測:
      みかんの右に東棟    列移動 75% ／ セル分割 25%
      担当者に分類して    集計 87.5% ／ ピボット 12.5%
  ★ まったく違う操作が、同じ依頼に 4 回に 1 回返る。しかもその割合は走行の間で動く
    （別の走行では セル分割 9/10 ── p=0.25 なら確率 0.003%）。
  ★ 開発側で検体を足しても出荷後の挙動は保証できない ── **走行時に自分で気づく**。
    連続 2 回は独立（食い違い 45%・独立なら 35.9%）なので、2 回目を引けば揺れる依頼で
    22〜38% 鳴る。易しい依頼（battery 20 種 × 10 回）は揺れ 0 ── そこでは鳴らない。

★★ 2026-09-19（実機 4 本を落として直した）: 受け皿は**読み直しの後**に置く。
  初版は生の翻訳を見て断っており、**製品が正しく扱える依頼を断っていた**:

      「原価の右に備考の列を追加して」→ 生の翻訳は **85% が セル分割**（20 回で実測）
       だが task_asks_to_add_a_column が依頼文だけで軸を決め、列追加へ直している。

  ★ 機械が自分で決めた回は、モデルが何を返そうと結果は同じ ── 断る理由が無い。
  ★ 「機械が決めたか」は読み直し層に言わせる（第 3 の戻り値）── 判定を書き写さない。

★ 何を守るか: **機械が自分で決められず**、かつ読み方が分かれた回に、
  黙ってどちらかで実行しないこと。
★ 何を守らないか: 揺れそのもの（これは受け皿であって治療ではない）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))
import ailine  # noqa: E402

MOVE = {"plan": [{"op": "MOVE_COLUMN", "args": {"col": "分類"}}]}
SPLIT = {"plan": [{"op": "SPLIT_CELL", "args": {"col": "分類", "sep": "の"}}]}
META = {"sheets": ["Sheet"], "headers": {"Sheet": ["分類", "金額", "原価"]}}


def _book(tmp_path: Path) -> Path:
    p = tmp_path / "在庫.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.append(["分類", "金額", "原価"])
    ws.append(["みかん", 100, 60])
    ws.append(["ぶどう", 200, 120])
    wb.save(p)
    wb.close()
    return p


def _fake_translate(answers: list, calls: list):
    """呼ばれるたびに answers を順に返す偽の翻訳（何回呼ばれたかを calls に残す）。"""
    def fake(model, task, book_meta, temperature=0.1):
        calls.append(task)
        return answers[min(len(calls) - 1, len(answers) - 1)]
    return fake


def _fake_preview(previews: list):
    """下書き当ての偽物（単体では LibreOffice を回さない）。

    ★ 差分は**計画の op の並び**で決める ── 同じ op なら同じ差分、違えば違う差分。
      エンジンの外の op（CLARIFY 等）は本物と同じく None（測れない）。
    """
    def fake(a, book, source_book, book_meta, plan):
        ops = tuple(str((s or {}).get("op") or "") for s in plan)
        previews.append(ops)
        if any(o not in ailine.CODEGEN_BY_OP for o in ops):
            return None
        # ★ 現実に寄せる: 同じ op を続けて 2 回当てても冊は 1 回分しか変わらない
        #   （列を末尾へ 2 回移す → 2 回目は空振り）。本物の証明は実機の
        #   tests/test_two_readings_are_compared_on_the_book.py が LibreOffice で行う。
        key = tuple(dict.fromkeys(ops))
        return ("＊変更: " + "+".join(key),)
    return fake


def _run(monkeypatch, tmp_path, argv_tail, answers, calls, previews=None):
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.delenv("AILINE_SINGLE_READ", raising=False)
    monkeypatch.setattr(ailine, "translate_task", _fake_translate(answers, calls))
    monkeypatch.setattr(ailine, "preview_plan_changes",
                        _fake_preview(previews if previews is not None else []))
    book = _book(tmp_path)
    before = book.read_bytes()
    # ★ 依頼は 2 つの条件を満たすものを実測で選んだ:
    #   ① 読み直し層が claim しない（「みかんの右に東棟」は 1 セル書換へ直されてしまう）
    #   ② 位置が機械に解ける（「右へ」は解けず、一致した回まで別の理由で断られる）
    rc = ailine.main(["run", str(book), "分類の列を末尾に移して", "--dry", *argv_tail])
    return rc, book.read_bytes() == before


# --- 事故そのもの --------------------------------------------------------------------

def test_a_split_reading_stops_before_acting(monkeypatch, tmp_path, capsys):
    """★★ 1 回目が列移動・2 回目がセル分割 → 実行せず、2 通りを見せて選ばせる。"""
    calls = []
    rc, unchanged = _run(monkeypatch, tmp_path, [], [MOVE, SPLIT], calls)
    out = capsys.readouterr().out
    assert rc == 3, out
    assert unchanged, "★ 読み方が分かれたのに冊が変わっている"
    assert len(calls) == 2, f"翻訳の回数が 2 でない: {len(calls)}"
    assert "読み方が分かれました" in out, out
    assert "読み方 1: 列移動" in out and "読み方 2: セル分割" in out, out
    # ★ 行き止まりにしない ── 選べる形（候補: 行 → --op で固定できる）
    assert f"{ailine.CHOICE_PREFIX}MOVE_COLUMN" in out, out
    assert f"{ailine.CHOICE_PREFIX}SPLIT_CELL" in out, out
    assert "--op" in out, "★ どうすれば通るかを言っていない（断りの 5 条）"


def test_an_agreed_reading_proceeds(monkeypatch, tmp_path, capsys):
    """★ 2 回とも同じ読みなら、従来どおり進む（2 回目を引いたこと以外は何も変わらない）。"""
    calls = []
    rc, _ = _run(monkeypatch, tmp_path, [], [MOVE, MOVE], calls)
    out = capsys.readouterr().out
    assert rc != 3, out
    assert len(calls) == 2, f"翻訳の回数が 2 でない: {len(calls)}"
    assert "読み方が分かれました" not in out


# --- 2 回目を引かない回（守る行動が無い／人が選んでいる）-----------------------------------

def test_a_non_acting_first_reading_is_not_rechecked(monkeypatch, tmp_path, capsys):
    """★ 1 回目が聞き返しなら 2 回目を引かない（行動しない回に 3 秒足しても守れない）。"""
    calls = []
    clarify = {"plan": [{"op": "CLARIFY", "question": "どの列ですか"}]}
    _run(monkeypatch, tmp_path, [], [clarify, MOVE], calls)
    assert len(calls) == 1, f"聞き返しなのに 2 回目を引いた: {len(calls)}"


def test_a_forced_op_is_not_rechecked(monkeypatch, tmp_path, capsys):
    """★ `--op` で固定した回は、人の選択に機械が別案を当てない。"""
    calls = []
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": "MOVE_COLUMN", "args": {"col": "分類"}})
    _run(monkeypatch, tmp_path, ["--op", "MOVE_COLUMN"], [MOVE, SPLIT], calls)
    assert len(calls) == 0, f"--op 固定なのに素の翻訳を引いた: {len(calls)}"


def test_the_opt_out_is_honored(monkeypatch, tmp_path, capsys):
    """★ AILINE_SINGLE_READ=1 なら 1 回だけ読む（速度を取りたい人の逃げ道）。"""
    calls = []
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setenv("AILINE_SINGLE_READ", "1")
    monkeypatch.setattr(ailine, "translate_task", _fake_translate([MOVE, SPLIT], calls))
    ailine.main(["run", str(_book(tmp_path)), "分類の列を末尾に移して", "--dry"])
    assert len(calls) == 1


# --- 何を「分かれた」と数えるか（恒真殺し・拾いすぎ殺し）-------------------------------------

def test_clarify_wording_differences_are_not_a_split(monkeypatch):
    """★★ 拾いすぎ殺し: 引数（聞き返しの文言）が違うだけでは分かれたと言わない。

    ★ 実測: CLARIFY の質問文は同じ依頼でも 4/10 で言い回しが変わる。引数まで比べると
      「行挿入 → 聞き返す」の 2 段計画が毎回鳴る（オオカミ少年）。
    """
    first = {"plan": [{"op": "INSERT_ROWS", "args": {"at": 3}},
                      {"op": "CLARIFY", "question": "梨の売上はいくらですか"}]}
    second = {"plan": [{"op": "INSERT_ROWS", "args": {"at": 3}},
                       {"op": "CLARIFY", "question": "『梨』の行に入れる値を教えてください"}]}
    monkeypatch.delenv("AILINE_SINGLE_READ", raising=False)
    monkeypatch.setattr(ailine, "translate_task", lambda *a, **k: second)
    assert ailine.recheck_translation("m", "t", META, first) is None


def test_a_different_plan_length_is_a_split(monkeypatch):
    """★ 段数の違いは分かれた読み（実測: 抽出 1 段 ／ 抽出 2 段 が 4:6 で揺れた）。"""
    one = {"plan": [{"op": "EXTRACT", "args": {"col": "分類", "cmp": "eq", "value": "x"}}]}
    two = {"plan": [one["plan"][0], one["plan"][0]]}
    monkeypatch.delenv("AILINE_SINGLE_READ", raising=False)
    monkeypatch.setattr(ailine, "translate_task", lambda *a, **k: two)
    assert ailine.recheck_translation("m", "t", META, one) is two


def test_same_head_op_offers_no_false_choice(capsys):
    """★ 先頭の op が同じで段数だけ違う時、`--op` では選べない ── 偽の候補を出さない。"""
    one = {"plan": [{"op": "EXTRACT", "args": {}}]}
    two = {"plan": [{"op": "EXTRACT", "args": {}}, {"op": "EXTRACT", "args": {}}]}
    assert ailine._refuse_split_reading(one, two) == 3
    out = capsys.readouterr().out
    assert ailine.CHOICE_PREFIX not in out, f"選べない候補を出している:\n{out}"
    assert "言い直して" in out


# --- 案 D: 宣言でなく実体（冊の差分）で比べる ---------------------------------------------------

def test_readings_that_change_the_book_the_same_way_are_one_reading(monkeypatch, tmp_path, capsys):
    """★★ 案 D の要（実機を 2 度落とした形）: 列移動→列移動 と 列移動 は冊に起きることが同じ。

    ★ 宣言（op の並び）で比べると断り、効果の語彙で比べると粗すぎて素通りする ──
      両方を下書きに当てて差分で比べれば、書き方が違うだけの同じ読みは**断らずに通る**。
    """
    calls, previews = [], []
    # ★ push3 で実機を落とした形そのもの:「列移動 → 列移動」vs「列移動」。
    #   読み直し層はこの依頼を claim せず（実測）、畳みで 2 段目が消えて同じ計画になる。
    two = {"plan": [MOVE["plan"][0], dict(MOVE["plan"][0])]}
    rc, _ = _run(monkeypatch, tmp_path, [], [two, MOVE], calls, previews)
    out = capsys.readouterr().out
    assert "読み方が分かれました" not in out, out
    assert rc != 3, out
    assert len(previews) == 2, "両方を下書きに当てていない: " + str(previews)


def test_readings_that_change_the_book_differently_stop_with_the_effects(monkeypatch, tmp_path, capsys):
    """★ 差分が違えば**書く前に**止まり、候補に『冊に起きること』を添える（op 名だけにしない）。"""
    calls, previews = [], []
    rc, unchanged = _run(monkeypatch, tmp_path, [], [MOVE, SPLIT], calls, previews)
    out = capsys.readouterr().out
    assert rc == 3 and unchanged, out
    assert len(previews) == 2
    assert "冊に起きること" in out, "差分の言葉が無い（op 名だけで選ばせている）: " + out


def test_an_unmeasurable_reading_falls_to_the_refusing_side(monkeypatch, tmp_path, capsys):
    """★★ 測れなかった回（None）は**断る側**へ倒す ── 黙って通す方向には倒さない。"""
    calls = []
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.delenv("AILINE_SINGLE_READ", raising=False)
    monkeypatch.setattr(ailine, "translate_task", _fake_translate([MOVE, SPLIT], calls))
    monkeypatch.setattr(ailine, "preview_plan_changes", lambda *a, **k: None)
    rc = ailine.main(["run", str(_book(tmp_path)), "分類の列を末尾に移して", "--dry"])
    out = capsys.readouterr().out
    assert rc == 3 and "読み方が分かれました" in out, out


def test_asking_versus_acting_is_never_reconciled_by_a_preview(monkeypatch, tmp_path, capsys):
    """★★ 「片方は聞き返し・片方は実行」は下書きでは比べられない ── 必ず止まる。

    ★ 実測（40 回振り）で 8 割は確認を求め 2 割は黙って実行する依頼が在った。
      黙って実行する側を引いた回にだけ書いてしまうのが、いちばん怖い形。
    """
    calls, previews = [], []
    # ★ 読み 1 は読み直し層が claim しない依頼で**実行に届く**もの（MOVE）にする ──
    #   INSERT_ROWS だと読み直し層が「入れる値が無い」と自前で断り、割れに届かない。
    acts = MOVE
    asks = {"plan": [MOVE["plan"][0], {"op": "CLARIFY", "question": "どの列の右ですか"}]}
    rc, unchanged = _run(monkeypatch, tmp_path, [], [acts, asks], calls, previews)
    out = capsys.readouterr().out
    assert rc == 3 and unchanged, out
    assert "読み方が分かれました" in out


def test_the_preview_never_touches_the_original_or_the_out(tmp_path):
    """★ 下書き当ては原本にも .out にも触らない ── 工房の中だけで済ませ、消して帰る。"""
    from _product_source import window_around
    body = window_around("def preview_plan_changes(", after=3600)
    assert "run_output_path(" not in body, "★ 下書き当てが .out の場所を使っている"
    assert "scratch.unlink()" in body, "★ 下書きを消していない"
    assert "redirect_stdout" in body, "★ 下書き当てが画面に出す（判定の材料が実行に見える）"


def test_a_second_reading_that_would_stop_is_a_split(monkeypatch, tmp_path, capsys):
    """★★ 変異試験が名指しした穴: 2 つ目の読みが**読み直し層で止まる**回は「片方は止まり、
    片方は実行する」── 黙って実行してはいけない形なので分かれた側へ倒す。

    ★ 読み 2 の INSERT_ROWS は、この依頼文だと読み直し層が「入れる値を依頼文から
      決められません」と自前で止める（実測）。下書きに当てる前に決まる。
    """
    calls, previews = [], []
    stops = {"plan": [{"op": "INSERT_ROWS", "args": {"at": 3}}]}
    rc, unchanged = _run(monkeypatch, tmp_path, [], [MOVE, stops], calls, previews)
    out = capsys.readouterr().out
    assert rc == 3 and unchanged, out
    assert "読み方が分かれました" in out, out
    assert previews == [], "★ 止まる読みを下書きに当てている（比べる前に分かる）: " + str(previews)


# --- 配線 ----------------------------------------------------------------------------------

def test_the_machine_deciding_for_itself_beats_the_split(monkeypatch, tmp_path, capsys):
    """★★ 実機 4 本を落とした形: 読み直し層が計画を決め直した回は、割れても断らない。

    ★ 「原価の右に備考の列を追加して」は生の翻訳が 85% セル分割（実測）だが、
      task_asks_to_add_a_column が依頼文だけで軸を決めて列追加へ直す ── そこでは
      モデルが何を返そうと結果は同じで、断ると**動く機能をコイン投げの断りに変える**。
    """
    calls = []
    split = {"plan": [{"op": "SPLIT_CELL", "args": {"col": "原価", "sep": "-"}}]}
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": "ADD_COLUMN", "args": {"name": "備考"}})
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.delenv("AILINE_SINGLE_READ", raising=False)
    monkeypatch.setattr(ailine, "translate_task",
                        _fake_translate([split, {"plan": [{"op": "ADD_COLUMN",
                                                            "args": {"name": "備考"}}]}], calls))
    book = _book(tmp_path)
    rc = ailine.main(["run", str(book), "原価の右に備考の列を追加して", "--dry"])
    out = capsys.readouterr().out
    assert len(calls) == 2, f"2 回読んでいない: {len(calls)}"
    assert "読み方が分かれました" not in out, (
        "★ 機械が列追加へ決め直したのに断っている（動く機能を潰す）: " + out)
    assert rc != 3, out


def test_the_reread_layer_reports_whether_it_decided():
    """★ 第 3 の戻り値が在ること ── ここが消えると受け皿が判定を書き写す側へ戻る。"""
    from _product_source import count_in_product
    # ★ 案 D で 2 つ目の読みも同じ層を通すようになった: 定義 1 + 呼び出し 2。
    assert count_in_product("_reread_the_plan(") == 3, "★ 定義 1 + 呼び出し 2 でない"
    assert count_in_product("_machine_decided") == 2, (
        "★ 読み直し層の『自分で決めたか』が受け皿へ渡っていない")


def test_both_translation_call_sites_are_wired():
    """★★ 片配線を作らない ── 翻訳の呼び口は 2 つ（1 冊の run と N 冊の run）。両方が呼ぶ。

    ★ 場所で決め打ちしない（tests/test_guard_ledger.py が鳴る）── 製品の出所を配線経由で読む。
    """
    from _product_source import count_in_product
    assert count_in_product("recheck_translation(") == 3, (
        "★ 定義 1 + 呼び出し 2 になっていない（片方の呼び口が外れたか、判定を書き写した）")
    assert count_in_product("_refuse_split_reading(") == 3, (
        "★ 断りが 2 つの呼び口の両方から呼ばれていない")
