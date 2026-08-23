"""`python -m ailine` の入口。

★ wheel 化（2026-08-23）で ailine.py は src/ailine/__init__.py になった。
   subprocess から叩く経路（テスト・利用者のスクリプト）はこの形を使う ──
   ファイルの置き場所に依存しない唯一の呼び方。install 済みなら `ailine` コマンドも同じ。
"""
from . import main

raise SystemExit(main())
