"""C1（振る舞いの凍結）: ゴールデンテスト共通ハーネス。

★ 目的: 流れ層の再設計に入る前に「今の振る舞い」を機械可読な形で固定する。
本番コード(ailine.py)は一切変更しない — ここは tests/ 配下だけの追加。

## 更新の作法（承認なしに黙って更新されるゴールデンは番人ではない）

1. 挙動を意図的に変える commit を作る前に、まず `AILINE_REGEN_GOLDEN=1` を立てて
   該当テストを実行し、ゴールデンファイルを再生成する。
   例（PowerShell）: `$env:AILINE_REGEN_GOLDEN=1; python -m pytest tests/test_golden_codegen.py -q; Remove-Item Env:AILINE_REGEN_GOLDEN`
   例（bash）: `AILINE_REGEN_GOLDEN=1 python -m pytest tests/test_golden_codegen.py -q`
2. `git diff -- tests/golden` で差分を人が読み、意図した変更だけかを確認する
   （意図しない差分が混じっていたら regenerate をやり直す）。
3. 差分が意図どおりであることを確認してから commit する。
   ★ このハーネス自身は絶対に「差分があれば自動で追随する」ことをしない
   （AILINE_REGEN_GOLDEN が立っていない限り、ゴールデンと実際の出力が食い違えば
   必ずテスト失敗にする＝これが番人の唯一の仕事）。

## ファイル形式

- .bas: codegen_dsl の生バイト列（byte-for-byte 比較。改行は LF 固定）
- .json: 構造化データ（ensure_ascii=False・indent=2 で人が読める形に整形）
- .txt: 文字列そのもの（--help 出力・端末トランスクリプト等）

いずれも write_bytes/read_bytes で扱い、CRLF 化を起こさない（Windows の text mode で
open すると LF→CRLF に化けるため、開発機がどの OS でも同じバイト列になることを保証する）。
"""
import json
import os
from pathlib import Path

GOLDEN_ROOT = Path(__file__).resolve().parent
REGEN = os.environ.get("AILINE_REGEN_GOLDEN") == "1"


def assert_golden_bytes(path: Path, actual: bytes, label: str = "") -> None:
    """バイト列を golden ファイルと突合する。REGEN=1 なら書いて即 pass。"""
    if REGEN:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(actual)
        return
    if not path.exists():
        raise AssertionError(
            f"ゴールデンが無い: {path}\n"
            f"AILINE_REGEN_GOLDEN=1 で再生成してから diff を確認すること。"
            f"{('(' + label + ')') if label else ''}")
    expected = path.read_bytes()
    if expected != actual:
        raise AssertionError(
            f"ゴールデンと不一致: {path}\n"
            f"--- expected ---\n{expected.decode('utf-8', errors='replace')}\n"
            f"--- actual ---\n{actual.decode('utf-8', errors='replace')}\n"
            f"意図した変更なら AILINE_REGEN_GOLDEN=1 で再生成し、git diff を確認して"
            f"commit すること。{('(' + label + ')') if label else ''}")


def assert_golden_text(path: Path, actual_text: str, label: str = "") -> None:
    assert_golden_bytes(path, actual_text.encode("utf-8"), label=label)


def assert_golden_json(path: Path, actual_obj, label: str = "") -> None:
    text = json.dumps(actual_obj, ensure_ascii=False, indent=2, sort_keys=False) + "\n"
    assert_golden_bytes(path, text.encode("utf-8"), label=label)


def sorted_list(s) -> list:
    """set を JSON 化できる決定論的な list に変換する（順序ゆらぎで golden が偽陽性で
       赤くならないように必ずソートする）。"""
    return sorted(s)
