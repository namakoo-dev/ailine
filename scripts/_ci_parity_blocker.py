# -*- coding: utf-8 -*-
"""CI と同じ素の環境を作る遮断器。scripts/ci_parity.py が ALLOWED を書き込んで走らせる。

★ 2026-08-24: 初版は find_module（Python 3.12 で**廃止済み**）で書いていて、
  meta_path に入れても一度も呼ばれなかった ── 遮断しているつもりで何も遮断して
  いない番人だった。しかも『Pillow 無しで緑』という俺の実証も、実は空振りだった。
  変異試験（Pillow を要求する検体を足して赤くなるか）で発覚。現行 API は find_spec。
"""
import sys

ALLOWED = set()          # ci_parity.py が上書きする
STDLIB = set(sys.stdlib_module_names)


class Block:
    def find_spec(self, name, path=None, target=None):
        top = name.split('.')[0]
        if top in STDLIB or top in ALLOWED:
            return None
        raise ImportError(
            f'{name} は requirements-dev.txt に無い（CI には入っていません）。'
            ' 手元に在るから通っているだけです ── 依存を外すか、宣言してください')


def run(allowed):
    ALLOWED.update(allowed)
    sys.meta_path.insert(0, Block())
    import pytest
    return pytest.main(['tests', '-q', '-m', 'not local'])
