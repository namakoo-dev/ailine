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
