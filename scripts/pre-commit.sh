#!/bin/sh
# ★★ 2026-09-15: 改行コードと制御文字の番人を **commit の瞬間** に走らせる（2.4 秒）。
#
# なぜ在るか（今日の実測）:
#   番人（tests/test_no_control_chars.py）は**元から在って、毎回ちゃんと鳴った**。
#   問題は鳴る場所が**全件テストの末尾**だったこと ── 1 日に 3 回、Python から
#   `pathlib.Path.write_text()` で repo のファイルを書き戻して LF を CRLF に変え、
#   そのたび 5 分 20 秒の全件を走らせ切ってから気づいた（2 回ぶん＝約 11 分を捨てた）。
#   ★ 検出は失敗していない。**検出が遅い**ことが損だった ── 直すのはループの長さ。
#   ★ 「気をつける」は効かなかった（記憶に `write_text は改行を壊す` と書いてあり、
#     セッション冒頭で読んだうえで 3 回踏んだ）。規律でなく機械にする ──
#     実機の鍵（scripts/machine_lock.py）を作った時と同じ結論。
#
# 走らせるもの: 速い番人だけ（重い番人は pre-push と CI が受け持つ・ここで遅くすると外される）。
# fail closed: pytest が走らせられなかった時も止める（出ないことは信号でない）。
# バイパス: git commit --no-verify（理由が記録に残る）
#
# ★ global の ~/.githooks/pre-commit（機密語スキャン）が、この repo 固有フックを
#   **自分の後ろで**呼ぶ作りになっている（core.hooksPath を設定すると git は呼ばないため）。
#   だから機密語の番人を壊さずに足せる。順は 機密語 → ここ。
set -e
root=$(git rev-parse --show-toplevel)
cd "$root"

printf '▶ pre-commit: 改行コードと制御文字…'
if ! out=$(python -m pytest -q -p no:cacheprovider --no-header tests/test_no_control_chars.py 2>&1); then
    printf '\n%s\n' "$out" >&2
    echo "" >&2
    echo "✗ pre-commit: 改行コード/制御文字の番人が止めました。" >&2
    echo "  よくある原因: Python から repo のファイルを書き戻す時に write_text() を使った" >&2
    echo "  （Windows では LF が CRLF に化けます）── 書き戻しは **write_bytes** で。" >&2
    echo '  直す: python -c "import pathlib;p=pathlib.Path(FILE);p.write_bytes(p.read_bytes().replace(b(CR LF),b(LF)))"' >&2
    echo '        （素直に書くなら: 読んだ bytes の \r\n を \n に置換して write_bytes）' >&2
    echo "  意図的に通すなら: git commit --no-verify" >&2
    exit 1
fi
printf ' ✓\n'
exit 0
