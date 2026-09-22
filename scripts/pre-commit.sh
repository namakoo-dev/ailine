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

# ★★ 2026-09-22 追記: **公開面の凍結**もここで見る（0.19 秒）。この日、署名を 2 つ
#   変えたまま push し、素の環境で 35 分走らせた末に止められた。番人は正しく鳴った ──
#   鳴る場所が push の中だっただけ。★ しかもその時、**意図していなかった 3 つ目**を
#   捕まえている（`from typing import NoReturn` で `ailine.NoReturn` が生えていた）。
#   安い番人ほど手前に置く。
printf '▶ pre-commit: 改行コード・制御文字・公開面…'
# ★★ 2026-09-21 追記: 改行の**混在**（1 ファイルの中に CRLF と LF が同居）も
#   ここで見る。同じ日に 2 回、CRLF のファイルへ LF で 1 行足して pre-push まで
#   気づかなかった（17 分 x 2）。番人は在ったが、鳴る場所が push の中だった。
#   ★ この番人は **git の index** を見る（tests/test_line_endings_stay_consistent.py の
#     docstring 参照）。commit の瞬間なら staged の中身が index に在るので捕まる。
if ! out=$(python -m pytest -q -p no:cacheprovider --no-header tests/test_no_control_chars.py tests/test_line_endings_stay_consistent.py tests/test_public_surface_is_frozen.py 2>&1); then
    printf '\n%s\n' "$out" >&2
    echo "" >&2
    echo "✗ pre-commit: 改行コード/制御文字/改行の混在/公開面の番人が止めました。" >&2
    echo "  ★ 公開面（名前と署名）を意図して変えたなら、記録を作り直してください:" >&2
    echo "    AILINE_REGEN_SURFACE=1 python -m pytest tests/test_public_surface_is_frozen.py" >&2
    echo "    ★ 作り直したら git diff で中身を読むこと（増えたぶんは意図した追加か）。" >&2
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
# --- 3 本目: バグ系のリンタと、絞った型検査（2026-09-22）------------------------
# ★ なぜ在るか: この日、私は 2 つの失敗をした ── 存在しない名前を書いたのと、
#   `urllib.error` を名指しで輸入せずに使っていたの。どちらも**書いた瞬間**に道具が
#   指したもので、**全件テストでは出なかった**（試験が通る経路では踏まないため）。
#   ★ 5091 件の試験が緑でも見えない族が在る ── だから足した。
# ★ 版を固定する。固定しないと、道具が上がった日に門が勝手に変わる
#   （requirements-dev.txt と同じ作法）。uvx なので環境には入れない。
# ★ 何を入れ、**何を入れなかったか（とその実測値）**は tests/pyright_gate.md に書いてある。
#   0 件でない規則は入れない ── 初日から赤い門は、そのうち誰も見なくなる。
# ★ 変異試験で鳴ることを確かめてある: 未代入・未定義・無い属性・無効なエスケープ・self の誤り
#   の 5 つを入れた検体で 5 件すべて鳴り exit 1、外して 0 件 exit 0。
# ★ 実測（暖まった状態）: ruff 0.16 秒 / pyright 8.5 秒。
py_staged=$(echo "$staged" | grep -E '^(src|gui)/.*[.]py$' || true)
if [ -n "$py_staged" ]; then
    if ! command -v uvx >/dev/null 2>&1; then
        echo "✗ pre-commit: uvx が無いので、リンタと型検査を走らせられません。" >&2
        echo "  ★ 走らせられなかったことを『指摘なし』と読みません（出ないことは信号でない）。" >&2
        echo "  直す: winget install astral-sh.uv   もしくは   pip install uv" >&2
        echo "  意図的に通すなら: git commit --no-verify" >&2
        exit 1
    fi

    printf '▶ pre-commit: バグ系のリンタ…'
    if ! out=$(uvx ruff@0.16.8 check --select=F,E9 --output-format=concise src/ gui/ 2>&1); then
        printf '\n%s\n' "$out" >&2
        echo "" >&2
        echo "✗ pre-commit: ruff（F=バグ系 / E9=構文）が止めました。" >&2
        echo "  ★ ここは書式の好みではありません ── 未定義の名前・重複した定義・" >&2
        echo "    使われない輸入など、**動きに関わる**ものだけを見ています。" >&2
        echo "  意図的に残すなら、その行に理由つきの # noqa を付けてください。" >&2
        echo "  意図的に通すなら: git commit --no-verify" >&2
        exit 1
    fi
    printf ' ✓\n'

    printf '▶ pre-commit: 型の門（8 秒）…'
    if ! out=$(uvx pyright@1.1.414 -p tests/pyright_gate.json 2>&1); then
        printf '\n%s\n' "$out" >&2
        echo "" >&2
        echo "✗ pre-commit: 型の門が止めました。" >&2
        echo "  見ているのは 5 つだけです: 経路によっては未代入 / 知らない名前 /" >&2
        echo "  無い属性 / 無効なエスケープ / self・cls の書き間違い。" >&2
        echo "  ★ 契約と、入れなかった規則の理由: tests/pyright_gate.md" >&2
        echo "  道具の型情報の限界なら、その行に理由を書いて" >&2
        echo "  # pyright: ignore[規則名] を付けてください（数を緩めるためには使わない）。" >&2
        echo "  意図的に通すなら: git commit --no-verify" >&2
        exit 1
    fi
    printf ' ✓\n'
fi

exit 0
