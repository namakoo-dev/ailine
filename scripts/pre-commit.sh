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
# ★★ 2026-09-21 追記: 改行の**混在**（1 ファイルの中に CRLF と LF が同居）も
#   ここで見る。同じ日に 2 回、CRLF のファイルへ LF で 1 行足して pre-push まで
#   気づかなかった（17 分 x 2）。番人は在ったが、鳴る場所が push の中だった。
#   ★ この番人は **git の index** を見る（tests/test_line_endings_stay_consistent.py の
#     docstring 参照）。commit の瞬間なら staged の中身が index に在るので捕まる。
if ! out=$(python -m pytest -q -p no:cacheprovider --no-header tests/test_no_control_chars.py tests/test_line_endings_stay_consistent.py 2>&1); then
    printf '\n%s\n' "$out" >&2
    echo "" >&2
    echo "✗ pre-commit: 改行コード/制御文字/改行の混在の番人が止めました。" >&2
    echo "  よくある原因: Python から repo のファイルを書き戻す時に write_text() を使った" >&2
    echo "  （Windows では LF が CRLF に化けます）── 書き戻しは **write_bytes** で。" >&2
    echo '  直す: python -c "import pathlib;p=pathlib.Path(FILE);p.write_bytes(p.read_bytes().replace(b(CR LF),b(LF)))"' >&2
    echo '        （素直に書くなら: 読んだ bytes の \r\n を \n に置換して write_bytes）' >&2
    echo "  意図的に通すなら: git commit --no-verify" >&2
    exit 1
fi
printf ' ✓\n'

# --- 2 本目: 図・行数・試験の本数が実体とずれていないか（2026-09-16）-------------
# ★ Namakoo「自動じゃなくてもフックで気付ける？ 実際の構成と図やグラフが一致してないと
#   それを追うのが不可能になる」── そのとおりで、この番人も**元から在った**。
#   居場所が全件の中、つまり **push の 30 分後**に鳴っていた。同じ日に 2 回それで止まった。
#   1 本目と同じ処方: 検出は失敗していない、**鳴る場所を手前に出す**。
# ★ 直さない（自動更新しない）。気づかせるだけにして、直す一行を画面に出す ──
#   commit の最中にファイルを書き換えると staged と working tree がずれる。
# ★ 費用は変わったファイルで絞る: 行数 0.1 秒 / 図 3.4 秒 / 試験の本数 6.8 秒。
#   絞らないと毎 commit 10 秒になり、いずれ外される（外された番人は在っても鳴らない）。
# ★ 測定の記録（効果の行列・翻訳精度）はここでは見ない ── 実機を 25 分回して初めて出る数字。
staged=$(git diff --cached --name-only)
parts=""
case "$staged" in *src/ailine/__init__.py*) parts="lines";; esac
case "$staged" in *tests/*.py*) parts="${parts:+$parts,}tests";; esac
case "$staged" in
    *src/ailine/*.py*|*src/ailine_core/*.py*) parts="${parts:+$parts,}graph";;
esac

if [ -n "$parts" ]; then
    printf '▶ pre-commit: 記録と実体の一致（%s）…' "$parts"
    if ! out=$(python scripts/refresh_records.py --parts "$parts" 2>&1); then
        printf '\n%s\n' "$out" >&2
        echo "" >&2
        echo "✗ pre-commit: 図や数が実体とずれています。" >&2
        echo "  ★ ずれたまま積むと、図を見ても実装を追えなくなります（それが図の存在理由）。" >&2
        echo "  直す:  python scripts/refresh_records.py --write" >&2
        echo "         作り直したら git add してから commit し直してください。" >&2
        echo "  ★ 測定の記録（効果の行列・翻訳精度）はこの道具では直りません ──" >&2
        echo "    実機を回して、出た数を人が tests/battery_recorded.json に書きます。" >&2
        echo "  意図的に通すなら: git commit --no-verify" >&2
        exit 1
    fi
    printf ' ✓\n'
fi
exit 0
