"""C1-F8: 全サブコマンドの --help 出力を丸ごと凍結する。CLI は買い手との契約面。

対象: トップレベル・run・stop・doctor・history・restore・undo・vocab・vocab add・
vocab list・alias・alias add・alias list・alias remove・alias undo の14通り。
★ W10 便A: alias 系4本（別名ストア）を追加。トップレベルの一覧に "alias" が
足されたことで top_level.txt も同時に更新が要る（サブコマンドを足せば必ず動く面）。
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
    "alias": ["alias"],
    "alias_add": ["alias", "add"],
    "alias_list": ["alias", "list"],
    "alias_remove": ["alias", "remove"],
    "alias_undo": ["alias", "undo"],
}


@pytest.mark.parametrize("name", sorted(CASES.keys()))
def test_help_output_golden(capsys, name):
    argv = CASES[name] + ["--help"]
    with pytest.raises(SystemExit) as exc:
        ailine.build_parser().parse_args(argv)
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert_golden_text(F8_DIR / f"{name}.txt", out, label=name)
