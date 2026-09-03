# GUI の番人（2026-08-26）── 画面は**判定を作らない**。
#
# ★ なぜ機械で縛るか: 「GUI は薄い殻にする」は口約束では守れない。
#   この repo は 2026-08 の盲検で「検算の分母が、疑うべき対象と同じ所から作られる」形の
#   欠陥を 4 回踏んだ。画面が advisories を数えたり postcondition から印を導いたりすれば、
#   それは 2 つ目の実装で、同じ欠陥をこちらで新造することになる。
#
# 契約:
#   ① 画面が読む判定は `verdict` ただ 1 つ（postcondition/advisories から導かない）
#   ② 判定の語彙は ailine 本体が出しうる値と一致する（画面が勝手な語を持たない）
#   ③ 殻は**この repo の本体**を叩く（site-packages の古い版を叩かない・実測で踏んだ）
#   ④ 新しい依存を足さない（標準ライブラリだけ）
#   ⑤ 画面を止めるモーダル（alert/confirm/prompt）を使わない

import ast
import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
GUI = REPO / "gui"
HTML = (GUI / "index.html").read_text(encoding="utf-8")
SERVER = (GUI / "server.py").read_text(encoding="utf-8")
sys.path.insert(0, str(REPO / "src"))


def _run_handler() -> str:
    """`/api/run` の処理ブロックだけを取り出す。

    ★ 実測で 2 度踏んだ: 固定の文字数で窓を切っていたので、コードが伸びるたびに
      番人が**中身を見なくなって**赤くなった（あるいは黙って通した）。
      ★ 窓は**構造で**切る ── 次の `elif u.path ==` までを 1 つの塊として読む。
    """
    i = SERVER.index('if u.path == "/api/run":')
    j = SERVER.index('elif u.path ==', i)
    return _code_only(SERVER[i:j])


def _code_only(py: str) -> str:
    """Python の行コメントを落とす。

    ★ 2 度踏んだ: 番人が**自分の説明文**に当たって赤くなる（JS 側でも同じことをした）。
      契約はコードについてのものなので、見る対象を正す ── 緩めてはいない。
      ★ 文字列リテラル中の `#` は落とさない（この番人が見る範囲には無いことを前提にする）。
    """
    return chr(10).join(ln.split("#")[0] for ln in py.split(chr(10)))


def _script(html: str, code_only: bool = False) -> str:
    m = re.search(r"<script>(.*?)</script>", html, re.S)
    assert m, "script が無い"
    js = m.group(1)
    if not code_only:
        return js
    # ★ 番人が**自分の説明文**に引っかかった（コメントで `advisories` や `alert(` に触れて
    #   いる）。契約はコードについてのものなので、見る対象を正す ── 緩めてはいない。
    #   ★ 行コメントだけを落とす（この画面にブロックコメントは無い ── 下で機械が確かめる）。
    assert "/*" not in js, "ブロックコメントが増えた（この番人の前提が崩れる）"
    return chr(10).join(re.sub(r"//.*$", "", ln) for ln in js.split(chr(10)))


# --- ① 判定を導かない -----------------------------------------------------------------

def test_the_page_never_derives_a_verdict():
    js = _script(HTML, code_only=True)
    for forbidden in ("postcondition", "advisories", "claims"):
        assert forbidden not in js, (
            f"画面が `{forbidden}` を読んでいる ── 判定を自分で導きかけている。"
            "映してよいのは ailine が返した verdict だけ")
    assert "j.verdict" in js, "verdict を読んでいない"


def test_the_server_does_not_touch_the_verdict():
    tree = ast.parse(SERVER)
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            assert node.value not in ("verified", "warned", "unverified", "unobservable"), (
                "サーバが判定の語を組み立てている（本体の値をそのまま運ぶだけにする）")


# --- ② 語彙が本体と一致する -----------------------------------------------------------

def test_the_page_knows_exactly_the_verdicts_the_tool_can_emit():
    """★ 恒真殺し: 本体が出す語を**本体側から**取り、画面の表と突き合わせる。

    本体が新しい判定を足したのに画面が知らなければ、無印で表示されてしまう。
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index('result["verdict"] = (')
    block = src[i:i + 400]
    emitted = set(re.findall(r'"(verified|warned|unverified|unobservable)"', block))
    assert emitted == {"verified", "warned", "unverified", "unobservable"}, emitted
    emitted.add("not_applied")     # 適用まで行かなかった run（_finish_run の既定）
    js = _script(HTML)
    known = set(re.findall(r"^\s{2}(\w+):\s*\{mark:", js, re.M))
    assert emitted <= known, f"画面が知らない判定がある: {sorted(emitted - known)}"


# --- ③ この repo の本体を叩く ---------------------------------------------------------

def test_the_shell_runs_this_repo_not_the_installed_copy():
    """★ 実測（2026-08-26）: これが無いと子プロセスは site-packages の**古い版**を
       import する。盲検 2 回目で検分者が古いタグを測ったのと同じ形の事故。"""
    assert 'env["PYTHONPATH"]' in SERVER, "本体の在り処を子プロセスに渡していない"
    assert 'REPO / "src"' in SERVER


# --- ④⑤ 依存とモーダル ----------------------------------------------------------------

def test_the_shell_adds_no_dependency():
    """④ 画面のために**新しい依存を増やさない**。

    ★ 2026-08-26 に契約を正した: 初版は「標準ライブラリだけ」と書いていたが、
      表を見せるために openpyxl（**既に製品が宣言している依存**）と ailine_core
      （この repo 自身）を使った時点で赤くなった。守りたいのは「新しい依存が増えないこと」
      なので、比べる相手を**pyproject の宣言**にする ── 手で書いた白名簿にしない
      （宣言が増えたらこの試験も自動で追随する）。
    """
    declared = set(re.findall(r'"([A-Za-z0-9_.-]+)\s*[><=!]',
                               (REPO / "pyproject.toml").read_text(encoding="utf-8")))
    declared = {d.split("[")[0].replace("-", "_").lower() for d in declared}
    own = {p.name for p in (REPO / "src").iterdir() if p.is_dir()}
    tree = ast.parse(SERVER)
    mods = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            mods.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            mods.add(node.module.split(".")[0])
    extra = mods - set(sys.stdlib_module_names) - declared - own
    assert not extra, (
        f"製品が宣言していない依存を画面が持ち込んだ: {sorted(extra)}"
        f"（宣言済み: {sorted(declared)} ／ 自前: {sorted(own)}）")


def test_no_blocking_modal():
    """★ 実測: alert() がレンダラを止め、画面が固まった（デモ中なら致命的）。"""
    js = _script(HTML, code_only=True)
    for bad in ("alert(", "confirm(", "prompt("):
        assert bad not in js, f"画面を止めるモーダルを使っている: {bad}"


# --- ⑥ 下書きは積み上がる（原本から作り直さない）----------------------------------------

def test_a_draft_run_saves_on_top_and_undo_steps_back():
    """★★ 2026-08-27 に**3 度目の設計変更**（Namakoo が正した）。

    経緯を全部残す（どの困りごとも本物だったから）:
      08-26 夜 「梨を追加」→「梨の売上を2000に」で **1 つ目が消えた** → 積み上げにした
      08-27 昼 同じ依頼を試すたびに **梨が増えた** → 原本からに変えた
      08-27 昼 すると続きの依頼でまた **梨が消えた**（08-26 夜の再演・俺が作り直した）
    ★ 正しい形: **逐次保存（積み上げ）が既定**。増えた分は **undo で 1 つ戻す**
      ── 直し方は「やり直さない」ではなく「**戻せる**」。
      やり直したい時だけ「原本からやり直す」。
    """
    block = _run_handler()
    assert 'not bool(req.get("restart"))' in block, "既定が積み上げでない"
    assert "shutil.copy2(book, draft)" in block, "やり直す経路が無い"
    js = _script(HTML, code_only=True)
    assert "runrestart" in js and "restart:true" in js, "画面に『原本からやり直す』が無い"


def test_undo_targets_the_thing_being_worked_on():
    """★ 「もとに戻す」が**原本**を対象にしていた。原本にはバックアップが無い
       （触っていないので当然）ので、**何も戻せなかった**。
    ★ 戻す相手は「いま触っている物」── 下書きがあれば下書き。
    """
    i = SERVER.index('elif u.path == "/api/undo":')
    block = _code_only(SERVER[i:i + 700])
    assert "_DRAFTS.get(book)" in block, "下書きを戻せない"
    assert "戻す相手" in SERVER, "どれを戻したか言っていない"


def test_the_page_says_which_file_and_sheet_it_is_touching():
    """★ 2026-08-27（Namakoo「今どのシートに対して操作を行っているのか分かりにくい」）:
       ファイル名もシート名も、表の下の小さい注記にしか出ていなかった。
    ★ 触っている物は**一番上に、常に**出す。
    """
    assert 'id="working"' in HTML
    js = _script(HTML, code_only=True)
    assert "setWorking(" in js and "いま触っている" in js
    assert "シート" in js, "シート名を出していない"


def test_the_result_pane_shows_the_file_actually_worked_on():
    """★ 実測: 断られた回に `j.out` が原本を指し、**下書きの表が原本に差し替わって**
       「梨の行が消えた」ように見えた（ファイルは無事だった）。
    ★ 出すのは**いま触っている物**（サーバが返した target）── j.out は補助でしかない。
    """
    js = _script(HTML, code_only=True)
    assert "window._target || j.out" in js, "j.out を先に信じている"


def test_the_draft_name_never_stacks():
    """★ `.out.out.out…` を作らない（実測で行き止まりになった形）。"""
    import importlib.util
    spec = importlib.util.spec_from_file_location("gui_server", REPO / "gui" / "server.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    once = mod._draft_path(Path("C:/x/売上.xlsx"))
    twice = mod._draft_path(once)
    assert once == twice, f"下書きの下書きを作った: {once} → {twice}"
    assert once.name == "売上（下書き）.xlsx", once.name


def test_the_page_says_where_it_is_writing():
    """★ 黙って別のファイルを触らない ── 今どこに書いているかを必ず見せる。"""
    js = _script(HTML, code_only=True)
    assert "draftnote" in js, "書き込み先を画面に出していない"
    assert "積み上げています" in HTML


def test_dropping_a_draft_does_not_delete_it():
    """『下書きを捨てる』は**手放すだけ**（ファイルは消さない）。

    ★ 消してしまうと、人が育てた成果物が画面のボタン 1 つで消える ──
      この repo が 2026-08-26 に直したばかりの事故（--copy の成果物の無言削除）と同じ形。
    """
    i = SERVER.index('/api/draft_reset')
    block = SERVER[i:i + 700]
    assert "unlink" not in block and "remove" not in block, "下書きを消している"
    assert "残っています" in block


# --- ⑦ 行き止まりを画面が持たない ------------------------------------------------------

def test_the_page_offers_a_way_out_of_a_blocked_output():
    """★ 2026-08-26（Namakoo が 2 度実測）: 前の run が残した作業結果が出力先を塞ぎ、
       **人が手でファイルを消すまで先へ進めない**状態になった。

    ★ 製品は断ったままにする（人が置いたファイルを黙って消さない設計は正しい）。
      画面は、断り文が言っている「別の場所へ移すか削除してから」を**人が 1 手で
      できる形**にする ── 決めるのは人、実行だけを楽にする。
    """
    js = _script(HTML, code_only=True)
    assert "clear_leftover" in js, "塞がれた時の抜け道が画面に無い（行き止まり）"
    assert "出力先に書けません" in js, "塞がれたことを画面が見ていない"


def test_clearing_a_leftover_renames_instead_of_deleting():
    """★ 片づけは**改名**であって削除ではない（取り返しを残す）。

    消してしまうと、この repo が 2026-08-26 に直したばかりの事故
    （--copy の成果物の無言削除）を、画面のボタンで再現することになる。
    """
    i = SERVER.index("/api/clear_leftover")
    block = _code_only(SERVER[i:i + 1400])
    assert "rename(" in block, "改名していない"
    assert "target.unlink()" not in block, "対象を消している"
    assert "（捨てた）" in SERVER


# --- ⑧ エクスプローラと複数ファイル ------------------------------------------------------

def test_the_shell_opens_a_real_dialog_because_the_browser_cannot():
    """★ ブラウザの file input は**完全なパスを返さない**（そういう作りになっている）。
       この道具は原本の場所を知らないと何もできないので、パスが要る。
       localhost に閉じた作りだから、サーバ側で OS のダイアログを開く。
    ★ tkinter は標準ライブラリ ── 依存の番人（上）が「増えていないこと」を確かめる。
    """
    assert "_native_dialog" in SERVER
    block = _code_only(SERVER[SERVER.index("def _native_dialog"):][:1400])
    for need in ("askopenfilename", "askdirectory", "asksaveasfilename"):
        assert need in block, f"{need} が無い（開く/フォルダ/保存の 3 つが要る）"
    assert "-topmost" in block, "前面に出していない（ブラウザの裏に出ると固まって見える）"


def test_saving_is_a_copy_not_a_move():
    """★ 「名前を付けて保存」は**複製**（元の物を消さない・下書きを育てたまま出せる）。"""
    i = SERVER.index("/api/save_as")
    block = _code_only(SERVER[i:i + 900])
    assert "shutil.copy2" in block, "複製していない"
    assert "rename(" not in block and "unlink()" not in block, "元の物を動かしている"


def test_the_multifile_paths_go_through_the_tool_not_the_page():
    """★ 複数ファイルでも画面は判定を作らない ── 本体のコマンドを叩くだけ。

    ★ 需要はここに寄っている（実測の需要地図: 上位 5 件中 4 件が複数ファイル）ので、
      画面から触れる形にした。判定の線は 1 冊の時とまったく同じ。
    """
    i = SERVER.index('elif u.path == "/api/folder":')
    block = _code_only(SERVER[i:i + 1800])
    for cmd in ('"scan"', '"stack"', '"verify"'):
        assert cmd in block, f"{cmd} を叩いていない"
    assert "verdict" not in block, "画面側の経路で判定を作りかけている"


def test_the_page_never_touches_an_element_that_does_not_exist():
    """★ 2026-08-26 に実測で踏んだ（Namakoo「実行できない」）。

    左側を作り直したとき `読み込む` ボタンを消したのに、JS はその id を触り続けた。
    `$("#reload").disabled` が null で落ち、**以降の配線が全部死んだ** ──
    画面は普通に出るのに、どのボタンも効かない。一番わかりにくい壊れ方。
    ★ 目で見ても分からないので、機械で照合する（id は増えるので、白名簿にしない）。
    """
    ids = set(re.findall(r'id="([\w-]+)"', HTML))
    js = _script(HTML, code_only=True)
    used = set(re.findall(r'\$\("#([\w-]+)"\)', js))
    missing = sorted(used - ids)
    assert not missing, f"JS が触るのに HTML に無い id: {missing}（null で以降が全部死ぬ）"


# --- ⑨ 「何が頼めるか」と「登録」が画面から届く ------------------------------------------

def test_the_page_shows_what_can_be_asked_from_the_tool_itself():
    """★ 2026-08-27（Namakoo）:「ailine ops 一覧も見方が分からない」。

    ★ 一覧は**本体から取る**（画面が別に持たない）── 持つと、op が増えた時に
      画面だけ古くなる。今日 3 op 足したばかりで、その形は避けたい。
    """
    i = SERVER.index('if u.path == "/api/ops":')
    block = _code_only(SERVER[i:i + 400])
    assert '_ailine(["ops"])' in block, "本体の ops を呼んでいない（画面が一覧を持っている）"
    js = _script(HTML, code_only=True)
    assert "/api/ops" in js and "opslist" in js


def test_a_refusal_points_at_the_place_to_register():
    """★「できない操作も登録まで誘導されないし、登録画面も実装されていない」。

    CLI には alias add が在るのに、画面からは一生辿り着けなかった
    ── **在るのに届かない機能は、無いのと同じ**。
    """
    js = _script(HTML, code_only=True)
    assert "/api/alias_add" in js, "登録の口が画面に無い"
    assert "頼める操作の一覧" in js, "断られたことを画面が見ていない"
    assert "aliasphrase" in HTML and "aliasop" in HTML, "登録の入力欄が無い"
    i = SERVER.index('/api/alias_add')
    assert '"alias", "add"' in SERVER[i:i + 900], "本体の alias add を呼んでいない"


# --- ⑩ 前後は同じファイルで撮る / 見るために開いたら閉じる --------------------------------

def test_before_and_after_come_from_the_same_file():
    """★ 2026-08-27（Namakoo「できていない」）: 「操作する前」に**原本**を、
       「操作したあと」に**下書き**を出していた ── 別のファイル同士を並べていたので、
       差分も色も意味を成さず、**正しく動いた操作が「できていない」に見えた**。

    ★ 前は**実行の直前にサーバが読む**（画面が後から取りに行くと、もう変わっている）。
    """
    block = _run_handler()
    assert block.count("_read_sheet(") >= 2, "実行の直前に『前』を読んでいない"
    assert '"before"' in block and '"target"' in block, "前と対象を返していない"
    js = _script(HTML, code_only=True)
    assert "res.before" in js, "画面が返された『前』を使っていない"


def test_every_workbook_opened_for_looking_is_closed():
    """★★ 2026-08-27（俺が連れてきたバグ）: read_only のブックは **close しないと
       ファイルハンドルが残る**。実行の直前に「前」を撮ったら、その掴んだままの
       ハンドルで下書きを置換できなくなり、**毎回 exit 9 で失敗**していた
       ── 画面には「できていない」としか見えない。

    ★ 見るために開いたものは、見終わったら必ず閉じる。機械で数える。
    """
    code = _code_only(SERVER)
    opens = code.count("load_workbook(")
    closes = code.count(".close()")
    assert closes >= opens, (
        f"開いた回数 {opens} に対して閉じた回数 {closes} ── "
        "掴んだままのハンドルは、置換を静かに失敗させる")


# --- ⑪ くらべる相手を人が選ぶ ------------------------------------------------------------

def test_the_comparison_basis_is_chosen_by_the_person():
    """★ 2026-08-27（Namakoo「操作前の表も書き換えられてしまう」）。

    下書きを積み上げる作りにしたので「操作する前」が毎回変わる。人が期待するのは
    **原本**（変わらないもの）だった ── 言葉の意味が俺の中で混ざっていた。

    ★ 外部の作法に合わせた:
      ・Google Sheets の版履歴は「どの版とくらべるか」を人に選ばせる
      ・コード編集系（VS Code / Cursor）は編集をその場で確定させず Keep / Undo を出す
    ★ そして**色は必ず、いま左に出ている物との差**にする ── 表示と基準がずれると
      色が嘘になる（この repo が繰り返し潰してきた「分母が別の所から来る」形）。
    """
    js = _script(HTML, code_only=True)
    assert "basis" in js and "redrawBasis" in js, "くらべる相手を選べない"
    assert 'id="basis"' in HTML
    for opt in ("原本", "直前"):
        assert opt in HTML, f"選択肢『{opt}』が無い"
    # ★ 恒真殺し: 基準を変えたら**色も塗り直す**（表示だけ変えて色が古いままにしない）
    i = js.index("function redrawBasis")
    block = js[i:i + 700]
    assert "drawTable($(\"#after\")" in block, "基準を変えても色を塗り直していない"


def test_continuing_is_disclosed():
    """★ 積み上げたことと、**1 つ戻せる**ことを同時に言う。

    ★ この文言は 1 日で 3 回変えた（積む → やり直す → 積む+戻せる）。
      文言が設計の言い換えになっているので、設計が動くたびにここも動く ──
      逆に言えば、**ここが古いままなら画面が嘘をついている**。
    """
    block = _run_handler()
    assert "続けて保存しました" in block, "積み上げたことを言っていない"
    assert "もとに戻す" in block, "戻せることを言っていない"
    assert "_cont" in block


def test_applying_to_the_original_copies_the_draft_instead_of_re_asking():
    """★ 2026-08-27（Namakoo「基本的に操作は下書きに対して行いたい」）。

    「原本に反映」は**原本に対して LLM を走らせ直して**いた ── 同じ依頼をもう一度
    翻訳するので、下書きで確かめた結果と**別の物ができうる**
    （LLM は揺れる。今日それで何度も踏んだ）。
    ★ 反映は**下書きを清書する**こと。翻訳も適用もやり直さない ──
      **人が目で確かめた物が、そのまま原本になる**。
    ★ これは「✓ の意味」と同じ線の話: 確かめた物と渡す物が違ってはいけない。
    """
    block = _run_handler()
    assert "shutil.copy2(_draft, book)" in block, "下書きを清書していない（翻訳し直している）"
    assert "_DRAFTS.get(book)" in block
    js = _script(HTML, code_only=True)
    assert 'task: ""' in js, "画面が反映で依頼文を送っている（頼み直しになる）"


def test_confirmations_can_be_answered_from_the_page():
    """★ 2026-08-27（Namakoo「y/N の入力ができない」）。

    子プロセスに端末が無いので、道具が `[y/N]` を聞く場面で**行き止まり**になっていた
    （上書きに限らず、確認を求める全部）。
    ★ 関所は 1 ミリも緩めない ── **人の答えを運ぶ道**を作る。
      押した時だけ `--overwrite` を付けて**もう一度**走らせる（黙って先に進めない）。
    """
    js = _script(HTML, code_only=True)
    assert "confirmrow" in js and "confirmyes" in js, "画面から答えられない"
    assert "overwrite: true" in js, "承知のうえで、を渡していない"
    assert 'id="confirmyes"' in HTML and 'id="confirmno"' in HTML
    block = _run_handler()
    assert '"--overwrite"' in block, "サーバが承知の合図を本体へ渡していない"
    assert "req.get(\"answer\")" in block, "答えを子プロセスへ運んでいない"


def test_the_gate_is_not_bypassed_by_default():
    """★ 恒真殺し: 既定では**絶対に** --overwrite を付けない（関所を素通りさせない）。"""
    block = _run_handler()
    assert 'if req.get("overwrite") else []' in block, (
        "無条件に上書きを許している疑い ── 関所は人が押した時だけ開く")


# --- 画面のスクリプトが**そもそも動くか**（2026-08-27・実測で 30 分溶かした）------------

def _gui_script() -> str:
    t = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    return t[t.index("<script>") + len("<script>"):t.index("</script>")]


def test_the_page_script_is_balanced():
    """★★ 実測（Namakoo「操作ができない」が 2 回）: 挿入の仕方を間違えて
       `$("#aliasadd").onclick = async () => {` を**二重に**書き込み、`{` が 1 つ余った。
       すると **script 全体が構文エラーで、関数が 1 つも定義されない**。
       画面は普通に描画されるので、**壊れているように見えない**のが最悪だった。
       ★ 同じ形（アンカーの二重書き込み）は今日 3 回踏んだ ── 人の目では見つからない。
       ★ 括弧の釣り合いだけなら外部の道具が要らない（CI に node は無い）。
    """
    import re
    src = _gui_script()
    # 文字列とコメントを潰してから数える（雑でよい ── 目的は「釣り合い」だけ）
    bs = chr(92)
    code = re.sub('"(?:[^"' + bs + bs + ']|' + bs + bs + '.)*"', '""', src)
    code = re.sub("'(?:[^'" + bs + bs + ']|' + bs + bs + ".)*'", "''", code)
    code = re.sub("`(?:[^`" + bs + bs + ']|' + bs + bs + ".)*`", "``", code)
    code = re.sub("//[^" + chr(10) + "]*", "", code)
    for open_c, close_c in (("{", "}"), ("(", ")"), ("[", "]")):
        assert code.count(open_c) == code.count(close_c), (
            f"画面のスクリプトで {open_c}{close_c} が釣り合っていない"
            f"（{code.count(open_c)} 対 {code.count(close_c)}）── "
            "構文エラーだと関数が 1 つも定義されず、画面は動かないのに壊れて見えない")


def test_no_line_repeats_the_same_handler_twice():
    """★ 二重書き込みそのものを名指しで捕まえる（釣り合いが偶然合う書き間違いもある）。"""
    for i, ln in enumerate(_gui_script().split("\n"), 1):
        for m in ("onclick = async", "onchange = async", "function "):
            assert ln.count(m) <= 1, f"{i} 行目に『{m}』が 2 回ある（挿入の二重書き込み）: {ln[:90]}"


def test_the_op_picker_does_not_hardcode_the_list():
    """★ 2026-08-27（Namakoo「登録はドロップダウンで」）: op 名を手で打たせていた
       （SORT と打てる人はまず居ない）。
       ★ ただし**画面が一覧を持たない**線は守る ── 選択肢は本体の `ailine ops --json`
         から取る。ここに op 名を書き並べた瞬間、増えた op が黙って落ちる。"""
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    assert 'id="aliasop"' in html and "<select id=\"aliasop\"" in html, "まだ入力欄のまま"
    assert "/api/oplist" in html, "一覧を本体から取っていない"
    js = _gui_script()
    for op in ("SORT", "COMPUTE_COLUMN", "APPEND_TOTAL", "EXTRACT"):
        assert op not in js, f"画面に op 名 {op} が焼き込まれている（一覧を持ってしまっている）"


def test_the_server_asks_the_product_for_the_op_list():
    src = (REPO / "gui" / "server.py").read_text(encoding="utf-8")
    assert '_ailine(["ops", "--json"])' in src, "一覧を本体から取っていない"


# --- 名前を積み上げない（2026-08-29・Namakoo が実測）------------------------------------

def _srv():
    import importlib.util
    spec = importlib.util.spec_from_file_location("gui_server", REPO / "gui" / "server.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


@pytest.mark.parametrize("name,want", [
    ("請求.xlsx", "請求（下書き）.xlsx"),
    ("請求（下書き）.xlsx", "請求（下書き）.xlsx"),
    ("請求.out.xlsx", "請求（下書き）.xlsx"),
    ("請求（捨てた）.xlsx", "請求（下書き）.xlsx"),
    # ★★ 実測で画面に出ていた形（札が 6 枚積み上がっていた）
    ("請求（下書き）.out（捨てた）（下書き）.out（捨てた）（下書き）.xlsx", "請求（下書き）.xlsx"),
])
def test_the_draft_name_never_accretes(name, want):
    """★★ 2026-08-29: 画面のファイル名がこうなっていた ──
       `1_請求_2026年8月（下書き）.out（捨てた）（下書き）.out（捨てた）（下書き）.xlsx`
    ★ 旧実装は「末尾が（下書き）なら足さない」だけで、末尾が（捨てた）の回に素通り
      していた ── **1 種類しか見ていない列挙**。札は 3 種類あるので在るだけ全部はがす。"""
    from pathlib import Path as _P
    assert _srv()._draft_path(_P(name)).name == want


def test_the_trash_rename_does_not_accrete():
    from pathlib import Path as _P
    m = _srv()
    assert m._canonical_stem("請求（下書き）.out（捨てた）") == "請求"


def test_discarded_files_are_not_offered_in_the_list():
    """★ 片づけた残骸を一覧に出すと、人が選べてしまい そこから名前が積み上がる。
       ★ 消してはいない ── フォルダには残る（取り返しは残す）。"""
    src = (REPO / "gui" / "server.py").read_text(encoding="utf-8")
    i = src.index("items = sorted((p for p in folder.iterdir()")
    seg = src[i:i + 400]
    assert "TRASH_SUFFIX not in p.stem" in seg, seg[:300]


def test_the_tool_suffixes_live_in_one_place():
    """★ 道具が付ける後ろ札は 1 箇所に持つ（別々の場所で書くと必ずずれる）。"""
    src = (REPO / "gui" / "server.py").read_text(encoding="utf-8")
    assert "_TOOL_SUFFIXES = (DRAFT_SUFFIX, TRASH_SUFFIX, \".out\")" in src
    # 直書きの「（捨てた）」が残っていないこと（定数を使う）
    assert src.count('"（捨てた）"') == 1, "『（捨てた）』を直書きしている箇所がある"


def test_the_sheet_choice_does_not_survive_a_file_change():
    """★★ 2026-08-29（Namakoo が実測）: 別のファイルを選んでも**前のシート名が残って**
       いたので、`？ シート『8月請求』がありません。あるシート: 売上表` で止まった。
    ★ シートの選択は**そのファイルのもの** ── 持ち越さない。"""
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    i = html.index("el.onclick = () => { picked = f;")
    seg = html[i:i + 700]
    assert "window._sheet = null" in seg, seg[:400]
    assert "showBefore()" in seg


def test_the_reading_panel_is_closed_when_the_file_changes():
    """★ 別のファイルを選んだのに、前のファイルの読みが残っていたら嘘になる。"""
    html = (REPO / "gui" / "index.html").read_text(encoding="utf-8")
    i = html.index("el.onclick = () => { picked = f;")
    assert "hideRead()" in html[i:i + 700]
