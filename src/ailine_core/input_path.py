"""入力のパスを受け取る唯一の入口 ── 無いものは**名指しで断る**（2026-09-13・需要① B2）。

★★ なぜ 1 本にするか（測ってから作った・分母つき）: 「パスを 1 文字打ち間違えた」という
  非プログラマが最も高い確率でやる間違いに、**同じ事故へ 4 通りの返事**が出ていた
  （12 経路を実測・2026-09-13）:

    scan / stack / forms          Python の**トレースバック**（FileNotFoundError）・exit 1
    csv / export-csv / export-pdf 「文書が無い」・exit 1
    split / accounts / verify      名指しの断り・exit 4
    run                            「文書が無い」・exit 9（`ENGINEERING.md` の表どおり）

  トレースバックは買い手に「道具が壊れた」と読める（自分の打ち間違いだと分からない）。
  番号が 4 通りだと、自動化から「打ち間違い」と「中身が決まらない」を見分けられない。

★ 意味と番号は `docs/ENGINEERING.md` の表が正 ── **「文書が無い」は 9**（実行の前提が無い）。
  「在るが中身から決まらない」は 4（関所で止めた・何も作っていない）で、別の事故。
★ ここは**判断しない**。在るかどうかを見て、無ければ人の言葉で断るだけ
  （直し方まで言う ── 打ち間違いは「何が違うか」が分かれば自分で直せる）。
"""
from __future__ import annotations

from pathlib import Path

#: 「入力が無い」の終了コード（`ENGINEERING.md` の 9「実行の前提が無い／文書が無い」）。
MISSING_INPUT_EXIT = 9

#: 打ち間違いの心当たり（★ 実務で本当に多い 3 つだけ ── 一般論を並べない）。
HINTS = (
    "・パスの打ち間違い（全角の ￥ ／ 全角の空白 ／ 末尾の空白）を確かめてください",
    "・エクスプローラからこの窓へ落とすと、正しいパスがそのまま入ります",
)


class MissingInput(Exception):
    """入力のパスが無い／種類が違う（打ち間違い・移動・まだ作っていない）。

    ★ これを捕まえて終わらせるのは `ailine.main` の 1 箇所だけ（番号の出口を二重化しない）。
    """


def _refuse(head: str, path, *extra) -> "MissingInput":
    lines = [f"× {head}: {path}", *extra, *HINTS]
    return MissingInput("\n".join(lines))


def require_folder(path) -> Path:
    """フォルダとして受け取れることを確かめて返す（無い／ファイルだった → 断る）。"""
    folder = Path(path).resolve()
    if not folder.exists():
        raise _refuse("フォルダが見つかりません", folder)
    if not folder.is_dir():
        raise _refuse("フォルダではなくファイルが指定されています", folder,
                      "・入っている「親フォルダ」を指定します（中の 1 冊ではありません）")
    return folder


#: 「開けない」を人の言葉にする（★ 例外の**型名**で分ける ── openpyxl を import しない）。
#: ★★ 2026-09-13（買い手役 3 体の初見・事務職が離脱を宣言した所）: 中身がテキストなのに
#:   拡張子だけ .xlsx のファイルで `run` が**英語のトレースバック 30 行**（BadZipFile）を出し、
#:   `forms` は「読み込み失敗: BadZipFile」と例外名を生で見せていた。1 つ前の `.xls` では
#:   日本語で完璧に案内していたのに ── 「無い」の門（上）の隣に「開けない」の門が無かった。
_UNREADABLE = (
    ("BadZipFile", "中身が Excel の形ではありません（拡張子だけ .xlsx になっている・"
                   "ダウンロードが途中で切れた・パスワード付き、など）"
                   "── Excel か LibreOffice で開いて .xlsx として保存し直してください"),
    ("InvalidFileException", "この形式は開けません{old}"
                             "── Excel か LibreOffice で .xlsx として保存し直してください"),
    ("PermissionError", "開けません（別のプログラムが使っている・読み取りの権限が無い）"
                        "── Excel で開いていれば閉じてください"),
)


def explain_unreadable(exc: BaseException, path) -> str:
    """開けなかった理由を、次の一手つきの 1 文にする（★ 例外名は出さない）。"""
    name = type(exc).__name__
    suffix = Path(path).suffix.lower()
    old = "（旧形式の .xls）" if suffix == ".xls" else f"（{suffix or '拡張子なし'}）"
    if suffix == ".csv":
        # ★ 2026-09-13（3 回目の買い手役・会計）: split が CSV を「保存し直して」と断る一方で、
        #   同じ道具の `ailine csv` が変換できる ── 自分の入口を案内する。
        return ("この入口は CSV を直接は受け取りません ── `ailine csv <ファイル>` で xlsx にすると"
                "（0 落ちを守って・原本は無変更）そのまま渡せます")
    for kind, text in _UNREADABLE:
        if kind == name or kind in str(exc):
            return text.format(old=old)
    return ("読めません（壊れているかもしれません）"
            "── Excel か LibreOffice で開いて保存し直すと直ることがあります")


def require_file(path, *, what: str = "文書") -> Path:
    """ファイルとして受け取れることを確かめて返す（無い／フォルダだった → 断る）。"""
    target = Path(path).resolve()
    if not target.exists():
        raise _refuse(f"{what}が見つかりません", target)
    if target.is_dir():
        raise _refuse(f"{what}ではなくフォルダが指定されています", target,
                      "・この入口は 1 冊を受け取ります（フォルダを渡す入口は `ailine ops` に出ています）")
    return target
