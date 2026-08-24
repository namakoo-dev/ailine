#!/bin/sh
# ★ 2026-08-24: CI で永久に走らない試験（-m local・22 ファイル）を、押す前に必ず走らせる。
#
# なぜ在るか: 実機 LLM/LibreOffice が要る試験は CI が `-m "not local"` で外している。
# その日のうちに 2 件（幻覚封鎖の誤爆・決裁③に追随しない古い契約）を CI が
# 一度も知らせないまま通した。番人は在ったが、CI という舞台に立っていなかった
# ──「在っても鳴らない」。
#
# fail closed: 走らせられなかった時も止める（出ないことは信号でない）。
# どうしても押すなら AILINE_SKIP_LOCAL=1 を明示的に付ける（理由が記録に残る）。
if [ "$AILINE_SKIP_LOCAL" = "1" ]; then
    echo "⚠ pre-push: 実機テスト(-m local)を **明示的に飛ばして** 押しています" >&2
    exit 0
fi
# ★ 2026-08-24（盲検査定の指摘）: リンタが CI にも pre-push にも入っていなかった。
# 製品コードに死んだ変数・未使用 import が 12 件溜まっていた ── 動作は壊さないが
# 「11,653 行の 1 ファイルが道具で手入れされていない」と読まれる。事実だった。
echo "▶ pre-push: 製品コードのリンタ…"
python scripts/lint_product.py
if [ $? -ne 0 ]; then
    echo "" >&2
    echo "✗ pre-push: リンタの指摘が残っています。" >&2
    exit 1
fi

echo "▶ pre-push: 実機テスト(-m local)を走らせます（CI では走らない分）…"
PYTHONPATH=src python -m pytest tests -q -m local
rc=$?
if [ $rc -ne 0 ]; then
    echo "" >&2
    echo "✗ pre-push: 実機テストが通っていません（exit $rc）。" >&2
    echo "  CI は -m \"not local\" なので、ここで止めないと誰も気づきません。" >&2
    echo "  ollama/LibreOffice が落ちているだけなら AILINE_SKIP_LOCAL=1 git push で越えられます。" >&2
    exit 1
fi
echo "✓ pre-push: 実機テスト通過"

# ★ 2026-08-24: 「手元に在って CI に無いもの」で CI が落ちる事故がこの repo で 4 度
#   起きている（lxml / ollama / LibreOffice / Pillow）。どれも手元では緑だった。
#   宣言していない依存を遮断して、CI と同じ素の環境で走らせる。
echo "▶ pre-push: CI と同じ素の環境で走らせます（宣言外の依存を遮断）…"
python scripts/ci_parity.py
rc=$?
if [ $rc -ne 0 ]; then
    echo "" >&2
    echo "✗ pre-push: 素の環境で落ちました（exit $rc）。" >&2
    echo "  手元に在るものに依存しています ── 依存を外すか requirements-dev.txt に宣言を。" >&2
    echo "  ★ 出力に returncode=3221225794 (0xC0000142) が並んでいる場合は別件です:" >&2
    echo "    Windows のプロセス資源の枯渇（重い走行を並行させすぎ）。" >&2
    echo "    他の pytest / push を止めてから、もう一度 push してください。" >&2
    exit 1
fi
echo "✓ pre-push: 素の環境でも通過"
