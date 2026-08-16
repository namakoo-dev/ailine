"""cli_render — C8: ailine.py に散った `print` のうち、複数経路で同じ形を手書きしていた
   ものを「文字列を組み立てる純関数」に寄せる。呼び側は組み立てた文字列/行リストを
   `print()` するだけにする（claim.py が『✓ 機械検証済み』でやった流儀を広げる）。

★ なぜ（C8 ブリーフより）: queue してある挙動変更2件（--dry プレビューの verified=True
誤認・差分番人と ✓ の矛盾警告）はどちらも表示に触る。表示が ailine.py 内に散っている
うちに直すと同じ場所を二度触ることになるため、先に層を分ける。

★ ここに移した範囲（と、あえて移さなかった範囲）: 実測で **複数の呼び出し箇所が同じ形を
手書きしていた** ものだけを対象にした（生成 .bas のコード表示ブロック3箇所・「続けるには
以下のいずれかを指定して」の再試行案内3箇所・「× 中止した」3箇所・restore/undo の
バックアップ一覧2箇所・run 見出し行3箇所・vocab の一覧/追加結果）。cmd_run_dsl や
cmd_run_plan の中に残る他の print（`？ {質問}` `× {エラー}` 等）は、その場限りの
1回書き（他に同形の呼び出し箇所が無い）で、関数に括り出しても呼ばれる場所が1つのままの
薄いラッパーにしかならないため、今回は動かしていない（「呼ばれていない/1箇所しか呼ばない
関数を作らない」という C8 ブリーフの縛り）。

★★ 出力は1バイトも変えない（純リファクタ）。retry-options のフラグ列は元の手書き文言が
グループごとに右端を揃えていた実測の桁（fidelity ゲート15桁・overwrite ゲート13桁・
freeform ゲート18桁）を、`max(len(flag)) + 2` で再現する形にした（ハードコードではなく
計算にした理由: 3ゲートで桁が違う＝手書きの揃え幅そのものに規則性があったため、その規則を
関数に持たせるほうが「桁だけ後で増える4つ目のゲート」が来ても自動で揃う）。
"""
from __future__ import annotations

from pathlib import Path

# --- 生成 .bas コード表示ブロック（単発 DSL / 自由生成の試行ループ / 複合計画の語彙外段） --

_CODE_BLOCK_FOOTER = "──────────────────────────────────────────"


def render_code_block(header: str, code: str, step_prefix: str = "") -> list:
    """生成コードの表示ブロック（見出し行・コード本体・区切り線）を行のリストにする。
       header は呼び出し側が組んだ見出し行そのもの（先頭の `\\n` の有無や文言は経路ごとに
       違う＝呼び出し側の責務のまま）。footer の区切り線だけが3経路で共通（実測で確認済み）。"""
    return [header, code, f"{step_prefix}{_CODE_BLOCK_FOOTER}"]


# --- 「続けるには以下のいずれかを指定して」再試行案内（忠実度/上書き/自由生成の3ゲート） --

def render_retry_options(step_prefix: str, options: list) -> list:
    """options: [(flag, 説明), ...]。フラグ列は同ブロック内の最長 flag + 2桁の空白幅で
       右側の説明を揃える（3ゲートとも手書きの揃え幅がこの計算と一致することを実測で
       確認済み — docs/behavior-corpus 側の挙動そのものは変えていない）。"""
    width = max(len(flag) for flag, _desc in options) + 2
    lines = [f"{step_prefix}この処理を続けるには、以下のいずれかを指定して再実行してください:"]
    for flag, desc in options:
        lines.append(f"{step_prefix}  {flag:<{width}}{desc}")
    return lines


def render_aborted(step_prefix: str = "") -> str:
    """対話で拒否された/非対話で確認できなかった時の中止行（3ゲートで同一文言）。"""
    return f"{step_prefix}× 中止した"


# --- run 見出し行（単発 DSL / 自由生成 / 複合計画で共通の骨格） ------------------------

def render_run_header(label: str, model: str, book_name: str) -> str:
    """「■ ailine（〜）  model=...  book=...」の骨格。label だけが経路ごとに違う
       （例: "DSL 経路" / "AI が直接作成・機械保証なし" / "複合計画・N 段"）。"""
    return f"■ ailine（{label}）  model={model}  book={book_name}"


# --- restore/undo のバックアップ一覧・復元結果（cmd_restore と cmd_undo が手書きで重複） --

def render_backup_list(book_name: str, backups: list) -> list:
    """`ailine restore --list` / `ailine undo --list` の一覧表示。backups は新しい順。"""
    if not backups:
        return [f"{book_name} のバックアップは無い"]
    lines = [f"{book_name} のバックアップ（{len(backups)} 世代・新しい順）:"]
    lines.extend(p.name for p in backups)
    return lines


def render_restore_done(book_name: str, used_name: str, remaining: int | None = None) -> str:
    """復元完了行。remaining（undo のみ、まだ戻せる回数）が渡されたときだけ注記を足す
       （restore は remaining=None のまま＝従来どおり注記なし）。"""
    suffix = f"（あと {remaining} 回戻せます）" if remaining is not None else ""
    return f"✓ {book_name} を {used_name} から復元した{suffix}"


# --- 用語集（vocab）コマンドの表示 -----------------------------------------------------

def render_vocab_add_result(ok: bool, msg: str) -> str:
    return ("✓ " if ok else "× ") + msg


def render_vocab_listing(vocab: dict, vocab_file: Path) -> list:
    """`ailine vocab list`。空なら登録方法の案内1行だけ返す。"""
    if not vocab:
        return [f"（用語集は空。{vocab_file} に登録するか `ailine vocab add <語> <値>` で追加）"]
    lines = [f"用語集（{vocab_file}・{len(vocab)}件）:"]
    lines.extend(f"  {term} = {vocab[term]:g}" for term in sorted(vocab))
    return lines
