"""C5: 『✓ 機械検証済み』相当の文字列リテラルが、ailine_core/claim.py（レンダラ）以外の
   .py ソースに現れたら赤くする番人。

   ★ 背景: ブラインド査定で、複合依頼「小計に数量×単価を入れて、見出しを太字にして」の
   2段目が実は前段の新規列に誤爆したのに『✓ 機械検証済み』が出た実測事故
   （docs/behavior-corpus/nodes/verification-scope-honesty.md）。原因は構造 —
   事後条件は計画が宣言した対象を検証しており機械は正しく動いていたが、検証していた
   対象が「依頼」でなく「計画」だった、という表示側のズレ。W10e で範囲注記を文字列で
   足して応急処置したが、文字列で謝っているだけで型が無かった。『✓ 機械検証済み』の
   文字列は複数箇所に散っており、次に経路が増えたときに同じ穴が開く。この番人は
   『機械検証済み』という文字列リテラルの発生源を機械的に1箇所（ailine_core/claim.py）
   に縛る。

   ★ 判定は AST ベース（grep 相当だが docstring は誤検知源になるため除外する）:
   関数/クラス/モジュールの docstring（先頭の裸の文字列文）は「文書化のための引用」
   として許容し、それ以外の文字列リテラル（print 引数・変数代入・return 値・f-string の
   固定部分など、実際に出力へ流れうるもの）だけを対象にする。コメント(#)は ast が
   そもそもパースしないため自動的に対象外。
   ★ 対象は ailine.py / ailine_core/*.py / bench/*.py（実際に出力に使われうる/出力文言を
   コピーされうる .py ソース）。tests/ 配下（ゴールデン・期待値としての引用が大量にある）
   は対象外 — 番人が守るのは「新しい出力経路」であって「出力を検証するテスト」ではない。
   ★ 純ロジックのみ・LibreOffice/basrun 不要（CI で走る）。"""
import ast
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
RENDERER_MODULE = REPO_ROOT / "src" / "ailine_core" / "claim.py"
MARKER = "機械検証済み"

_TARGET_FILES = (
    [REPO_ROOT / "src" / "ailine" / "__init__.py"]
    + sorted(p for p in (REPO_ROOT / "src" / "ailine_core").glob("*.py") if p != RENDERER_MODULE)
    + sorted((REPO_ROOT / "bench").glob("*.py"))
)


def _docstring_constant_ids(tree: ast.AST) -> set:
    """モジュール/クラス/関数の先頭の裸文字列文（docstring）の Constant ノード id を集める。"""
    ids = set()
    candidates = [tree] + [n for n in ast.walk(tree)
                            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))]
    for node in candidates:
        body = getattr(node, "body", None)
        if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) \
                and isinstance(body[0].value.value, str):
            ids.add(id(body[0].value))
    return ids


def _offending_lines(source: str, filename: str = "<string>") -> list:
    """source 内で、docstring 以外の文字列リテラルに MARKER を含む箇所の行番号を返す。"""
    tree = ast.parse(source, filename=filename)
    doc_ids = _docstring_constant_ids(tree)
    hits = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            if id(node) in doc_ids:
                continue
            if MARKER in node.value:
                hits.append(node.lineno)
    return hits


def test_renderer_module_exists_and_contains_the_marker():
    """除外対象そのものの健全性: claim.py は当然マーカーを含む（ここが正しい発生源）。"""
    assert RENDERER_MODULE.exists()
    assert MARKER in RENDERER_MODULE.read_text(encoding="utf-8")


def test_verified_marker_only_appears_in_claim_renderer():
    """『機械検証済み』という文字列リテラル（docstring を除く）が claim.py 以外の
       ソース（ailine.py / ailine_core/*.py / bench/*.py）に現れていないことを検査する。"""
    offenders = {}
    for path in _TARGET_FILES:
        hits = _offending_lines(path.read_text(encoding="utf-8"), filename=str(path))
        if hits:
            offenders[str(path.relative_to(REPO_ROOT))] = hits
    assert not offenders, (
        f"『{MARKER}』という文字列リテラルが ailine_core/claim.py 以外に見つかった "
        f"（C5: レンダラ1箇所に統合する約束が崩れている）: {offenders}"
    )


def test_guard_actually_fires_when_marker_placed_outside_renderer():
    """★ DoD5: 番人の発火実証。レンダラ以外の場所に一時的に該当リテラルを置いた文字列を
       直接 _offending_lines に通し、赤くなる（＝検出される）ことを確認する
       （実ファイルは書き換えない・関数を直接叩く自己検証）。"""
    poisoned_source = 'def f():\n    return "✓ 達成を機械検証済み（テスト用）"\n'
    hits = _offending_lines(poisoned_source, filename="fake_module.py")
    assert hits, "番人が発火しなかった（自己検証に失敗）"


def test_guard_ignores_docstrings_by_design():
    """docstring 中の言及（文書化目的の引用）は許容する、という設計判断そのものを検査する。
       ailine.py には『機械検証済み』を説明する docstring/コメントが複数あるが、これらは
       出力そのものではないため対象外でよい。"""
    docstring_only_source = (
        'def f():\n'
        '    """ここでは『機械検証済み』とは名乗らない、という説明。"""\n'
        '    return 1\n'
    )
    hits = _offending_lines(docstring_only_source, filename="fake_module.py")
    assert hits == [], "docstring 中の言及まで拾ってしまっている（誤検知）"
