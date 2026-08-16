"""C1-F8: 全サブコマンドの --help 出力を丸ごと凍結する。CLI は買い手との契約面。

対象: トップレベル・run・stop・doctor・history・restore・undo・vocab・vocab add・
vocab list の10通り。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402

from golden._harness import GOLDEN_ROOT, assert_golden_text  # noqa: E402

F8_DIR = GOLDEN_ROOT / "f8_help"

CASES = {
    "top_level": [],
    "run": ["run"],
    "stop": ["stop"],
    "doctor": ["doctor"],
    "history": ["history"],
    "restore": ["restore"],
    "undo": ["undo"],
    "vocab": ["vocab"],
    "vocab_add": ["vocab", "add"],
    "vocab_list": ["vocab", "list"],
}


@pytest.mark.parametrize("name", sorted(CASES.keys()))
def test_help_output_golden(capsys, name):
    argv = CASES[name] + ["--help"]
    with pytest.raises(SystemExit) as exc:
        ailine.build_parser().parse_args(argv)
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert_golden_text(F8_DIR / f"{name}.txt", out, label=name)
