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

REPO = Path(__file__).resolve().parent.parent
GUI = REPO / "gui"
HTML = (GUI / "index.html").read_text(encoding="utf-8")
SERVER = (GUI / "server.py").read_text(encoding="utf-8")
sys.path.insert(0, str(REPO / "src"))


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

def test_draft_continues_instead_of_restarting_from_the_original():
    """★ 2026-08-26（Namakoo が実測）: 「梨を追加して」→「梨の売上を2000に」と続けると、
       **1 つ目の操作が消えた**。

    根: 下書きが毎回 `run <原本> --copy` で、**毎回原本から作り直して**いた。
    人は「1 つやって、次をやる」と積み上げるのに、道具は積み上げていなかった。

    ★ 同じ日に**契約を 1 度訂正した**（最初の直しが連れてきたバグ・これも実測）:
      初版は `--copy` の成果物に反映し続けたので、作業ファイルが `<名前>.out.out.xlsx`
      と積み重なり、失敗した回の残骸が次の run を関所で塞いで**行き止まり**になった。
      → 下書きは**人が読める名前のファイル 1 つ**にして、そこへ普通に反映する。
    """
    assert "_DRAFTS" in SERVER, "下書きの続きを持っていない（毎回原本から作り直す）"
    i = SERVER.index('if u.path == "/api/run":')
    block = SERVER[i:i + 2000]
    assert "_draft_path(" in block, "下書きの置き場を決めていない"
    assert "shutil.copy2(book, draft)" in block, "1 回目に原本から作る経路が無い"
    assert "--copy" not in _code_only(block), (
        "下書きに --copy を使っている ── 名前が .out.out… と積み重なって行き止まりになる")
    # ★ 原本に反映したら下書きの役目は終わり（古い下書きに積み続けない）
    assert "_DRAFTS.pop(book, None)" in block


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
