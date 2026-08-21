"""番人: 拡張子判定は ailine_core/filetypes.py の登録簿に一元化されている
（単位F: 拡張子判定の登録簿統合）。

★ 主眼: `ailine.py` / `ailine_core/*.py`（filetypes.py 自身は白名簿）に
  suffix 判定の拡張子リテラル直書きが**再発しない**こと。AST で見る ── 文字列一致の
  grep だと「対象の文書 (.xlsx / .ods)」のような地の文（help メッセージ・docstring）まで
  誤検知するため、コード上の判定パターンだけを狙う:

  1. `path.suffix.lower() == ".xlsx"` のように、`.suffix` を含む式が拡張子リテラルと
     比較されている（in / == / != のいずれか）。
  2. `_SOME_EXTS = {".xlsx", ".xls"}` のように、拡張子っぽい文字列だけからなる
     集合/リスト/タプルが登録簿の外で新規に定義されている。
"""
import ast
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
AILINE_PY = REPO_ROOT / "ailine.py"
CORE_DIR = REPO_ROOT / "ailine_core"
REGISTRY_FILE = CORE_DIR / "filetypes.py"

_EXT_LITERAL_RE = re.compile(r"^\.[A-Za-z0-9]{1,6}$")


def _files_to_scan():
    yield AILINE_PY
    for py in sorted(CORE_DIR.glob("*.py")):
        if py.resolve() != REGISTRY_FILE.resolve():
            yield py


def _contains_suffix_attr(node: ast.AST) -> bool:
    return any(isinstance(n, ast.Attribute) and n.attr == "suffix" for n in ast.walk(node))


def _extension_literal_strings(node: ast.AST) -> list:
    """node 内の文字列定数のうち拡張子らしいもの（例: ".xlsx"）を列挙。"""
    found = []
    if isinstance(node, ast.Constant) and isinstance(node.value, str) and _EXT_LITERAL_RE.match(node.value):
        found.append(node.value)
    elif isinstance(node, (ast.Set, ast.List, ast.Tuple)):
        for elt in node.elts:
            found.extend(_extension_literal_strings(elt))
    return found


def _find_offenses(path: Path) -> list:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    offenses = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Compare):
            operands = [node.left] + list(node.comparators)
            if _contains_suffix_attr(node) and any(_contains_suffix_attr(o) for o in operands):
                literals = []
                for o in operands:
                    literals.extend(_extension_literal_strings(o))
                if literals:
                    offenses.append((node.lineno,
                                      f"suffix 判定に拡張子リテラルが直書き: {literals}"))
        if isinstance(node, ast.Assign) and isinstance(node.value, (ast.Set, ast.List, ast.Tuple)):
            elts = node.value.elts
            if len(elts) >= 2:
                literals = _extension_literal_strings(node.value)
                if len(literals) == len(elts):
                    offenses.append((node.lineno,
                                      f"拡張子リテラルの集合が登録簿の外で定義: {literals}"))
    return offenses


def test_no_inline_extension_suffix_literals_outside_registry():
    """`ailine.py` / `ailine_core/*.py` に拡張子判定のリテラル直書きが再発したら赤くする。
       新しい判定が要る時は ailine_core/filetypes.py に名前つきで足し、そこから import する。"""
    offenders = {}
    for path in _files_to_scan():
        offenses = _find_offenses(path)
        if offenses:
            offenders[str(path.relative_to(REPO_ROOT))] = offenses
    assert not offenders, (
        f"拡張子判定のリテラル直書きが見つかった（ailine_core/filetypes.py の登録簿に"
        f"追加し、そこから import すること）: {offenders}"
    )
