"""ailine の保存先（~/.ailine の下）を、名簿を手で持たずに**全部**切り離す（2026-09-23）。

★★ なぜ在るか（実害）: 導線の台帳の歩き方を載せ替えるサブエージェントが、歩き手を pytest の外から
  呼び、**本物の ~/.ailine に履歴 103 行・誤分類 2 行・undo の控え 4 つ**を書いた（Namakoo の許可を
  得て消した）。歩き手は製品を同じプロセスの中で走らせ、切り離しを**呼ぶ側に任せていた**。
★ しかも試験の番人（conftest）の切り離しは手で並べた 6 つで、ATTRIBUTES_FILE と HISTORY_DIR が
  抜けていた（手で並べた名簿の片配線）。→ ここで**モジュールから導く**（保存先が増えても自動で守られる）。
"""
from __future__ import annotations

import contextlib
import os
from pathlib import Path


def home_bound_paths(mod) -> dict:
    """モジュールの大域変数のうち、ホーム（HISTORY_DIR）の下にある Path を全部 {名前: Path}。"""
    home = Path(mod.HISTORY_DIR)
    out = {}
    for name, value in vars(mod).items():
        if name.isupper() and isinstance(value, Path):
            try:
                value.relative_to(home)
            except ValueError:
                continue
            out[name] = value
    return out


def is_isolated(mod, real_home: Path) -> bool:
    """保存先が 1 つも本物のホームの下に無いか。"""
    for p in home_bound_paths(mod).values():
        try:
            p.relative_to(real_home)
            return False
        except ValueError:
            continue
    return True


@contextlib.contextmanager
def isolated_home(mod, root: Path):
    """保存先を全部 root の下へ向け、終わったら元に戻す（AILINE_HOME も root に）。"""
    root.mkdir(parents=True, exist_ok=True)
    home = Path(mod.HISTORY_DIR)
    saved = home_bound_paths(mod)
    old_env = os.environ.get("AILINE_HOME")
    try:
        for name, p in saved.items():
            setattr(mod, name, root / p.relative_to(home))
        os.environ["AILINE_HOME"] = str(root)
        yield root
    finally:
        for name, p in saved.items():
            setattr(mod, name, p)
        if old_env is None:
            os.environ.pop("AILINE_HOME", None)
        else:
            os.environ["AILINE_HOME"] = old_env
