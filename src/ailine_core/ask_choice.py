"""ask_choice — 「選択肢を出して選ばせる」を1箇所に閉じ込めた対話部品。

★ なぜ部品にするか（オーナーの指示そのまま）: 棚に「受付層と語彙登録の UX 設計」
「`--assist` 対話 CLI」が積んである。今回のシート衝突の3択を**一回限りの分岐**として
ailine.py に書くと、次に対話が要るときに同じものがもう一度書かれる（この repo の欠陥の
共通の根＝「器官が呼び出し規約でなく記憶で適用されている」の再演。
ailine_core/stage_organs.py の docstring 参照）。シート衝突は**この部品の最初の利用者**
であって、部品の持ち主ではない。

★ 設計（純粋寄り・テストで stdin を触らずに検証できること）:
  - 入出力は必ず注入できる（input_fn / print_fn）。既定は組み込みの input/print。
  - 「聞いてよいか」の判定(interactive)は**呼び出し側が渡す**真偽値であって、この
    モジュールが sys.stdin を見に行くことはしない（is_interactive() は判定に必要な
    事実を引数で受け取る純関数。sys.stdin.isatty() を実際に呼ぶのは ailine.py 側）。
  - 聞けない場面（非対話）は**必ず既定で素通り**する。ここで止めると、パイプや CI で
    動いていたスクリプトが黙って固まる/壊れる（オーナーの縛り: 「止めると動いていた
    スクリプトが黙って壊れる」）。
  - ★ operator 盲検9回目 CONFUSING②: 素通りする際、以前はメニューも告知も一切出さず
    黙って既定へ進んでいた。実測では非対話でも `1) 2) 3)` のメニュー＋入力待ち風の
    記号が出た後で待たずに進む形になっており（スクリプト/CI 用途で紛らわしい）、
    メニュー自体は`interactive=False`なら出ないのが正しいが、それなら今度は
    「なぜ既定になったか」が一切見えない黙る素通りだった。★ メニューは出さないまま、
    既定を選んだ事実だけを1行で告知する形に直した（render_default_notice）。

★ 移植可能性（tests/test_line_budget.py が機械で守る）: ailine を import しない。
標準ライブラリのみ。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

# 不正な入力を何度まで聞き直すか。無限ループにしない（自動実行に紛れ込んだ場合に
# 固まらせない・超えたら「聞けなかった」扱いで既定へ倒す）。
MAX_REPROMPTS = 3


@dataclass(frozen=True)
class Choice:
    """1つの選択肢。key は利用者が打つ文字列（"1"/"2"/"3"）、text は表示文。"""
    key: str
    text: str


@dataclass(frozen=True)
class ChoiceResult:
    """ask_choice の戻り値。
       key: 選ばれた Choice.key。**聞かなかった/聞けなかった**場合は None。
       asked: 実際に画面へ出して聞いたか（非対話で素通りしたのか、聞いたが答えが
              得られなかったのかを呼び出し側が区別できるようにする）。"""
    key: str | None
    asked: bool


def is_interactive(*, stdin_isatty: bool, json_mode: bool = False, dry: bool = False) -> bool:
    """対話してよい場面かどうかの判定（純関数）。

    ★ 3つとも「聞かない」理由が違う:
      - stdin_isatty=False … パイプ/リダイレクト/CI。input() は EOF か無限待ちになる。
      - json_mode          … 出力を機械が読む。プロンプトを混ぜてはいけない。
      - dry                … もともと適用しないので、選ばせる意味が無い。
    """
    return bool(stdin_isatty) and not json_mode and not dry


def render_choice_block(lines: list, choices: list, indent: str = "  ") -> list:
    """前置きの行（lines）＋「  N) 文」の選択肢行を、印字用の行リストにして返す（純関数）。
       印字そのものはしない ── テストは出力を組み立てた文字列として検証できる。"""
    out = list(lines)
    out.extend(f"{indent}{c.key}) {c.text}" for c in choices)
    return out


def render_default_notice(choices: list) -> str:
    """★ operator 盲検9回目 CONFUSING②の直し: 非対話で素通りする時にメニューの代わりに
       出す1行。既定は**呼び出し側の並び順の先頭(choices[0])**という既存の慣例をそのまま使う
       （sheet_conflict_choice_lines の「1」＝「上の解釈のとおり実行する」がその慣例の由来。
       ask_choice はここでも choices の中身の意味を知らない・先頭という位置だけで決める）。"""
    return f"（非対話のため既定で続行: {choices[0].text}）"


def ask_choice(lines: list, choices: list, *, interactive: bool,
               input_fn: Callable | None = None, print_fn: Callable | None = None,
               prompt: str = "> ") -> ChoiceResult:
    """選択肢を出して1つ選ばせる。

    interactive=False なら**メニューは出さず**、既定へ倒した事実だけを1行印字して
    ChoiceResult(key=None, asked=False) を返す（呼び出し側はこの key=None を既定の
    挙動として扱う ── 何が起きたかを黙らない、というだけで判定ロジックは変えない）。
    choices が空なら（選ばせる対象が無い）その1行も出さずそのまま返す。
    EOF/入力不能・不正入力が MAX_REPROMPTS 回続いた場合も key=None（asked=True）を返し、
    決して例外を投げない。
    """
    if not choices:
        return ChoiceResult(key=None, asked=False)
    say = print_fn or print
    if not interactive:
        say(render_default_notice(choices))
        return ChoiceResult(key=None, asked=False)
    ask = input_fn or input
    for ln in render_choice_block(lines, choices):
        say(ln)
    valid = {c.key for c in choices}
    for _ in range(MAX_REPROMPTS):
        try:
            answer = (ask(prompt) or "").strip()
        except (EOFError, KeyboardInterrupt):
            return ChoiceResult(key=None, asked=True)
        if answer in valid:
            return ChoiceResult(key=answer, asked=True)
        say(f"{'/'.join(c.key for c in choices)} のいずれかで答えてください。")
    return ChoiceResult(key=None, asked=True)


def ask_yes_no(question: str, *, interactive: bool, default: bool = False,
               input_fn: Callable | None = None) -> bool:
    """y/N の確認（既存の関所と同じ流儀: y/yes だけを肯定とみなす）。
       interactive=False / EOF なら default（既定は False＝やめる側）。"""
    if not interactive:
        return default
    ask = input_fn or input
    try:
        answer = (ask(question) or "").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return default
    return answer in ("y", "yes")
