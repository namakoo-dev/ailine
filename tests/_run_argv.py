"""C2: `argparse.Namespace(...)` の手組みを `ailine.main(argv)` 用の argv へ変換する共有ヘルパ。

★★ 本番コード(ailine.py)は一切変更していない。ここは build_parser()（ailine.py 5537行目）
が定義する `ailine run` の CLI フラグと、手組み Namespace のフィールド名の対応を機械的に
持つだけで、意味は変えない（C1-F9 の test_golden_transcripts.py が確立した
`ailine.main(argv)` 経由の流儀を、test_ailine.py 側の cmd_run 直呼びテストにも広げる）。

`inplace` は受け取っても無視する。cmd_run 冒頭（ailine.py:4410）で
`a.inplace = not getattr(a, "copy", False)` が無条件に上書きするため、Namespace に
inplace=... を混ぜても cmd_run に渡った時点で意味を持たない（実測済み・旧テストが
inplace=True/False を書いていたのは死んだ引数だった）。--copy の有無だけが効く。
"""

_DEFAULTS = dict(
    model="qwen2.5-coder:7b", refs=None, helpers=None, repair=2, temperature=0.2,
    dry=False, json=False, timeout=180.0, ask=False, copy=False, values=False,
    header_row=None, accept_loss=False, overwrite=False, allow_freeform=False,
    keep_backups=None, sheet=None,   # ★ 挙動変更#2: --sheet（build_parser() に追加）
)


def run_argv(book, task, **overrides):
    """`["run", <book>, <task>, --flag ...]` を組み立てる。

    overrides には手組み Namespace と同じキー名をそのまま渡せる
    （book/task 以外は build_parser() の dest 名と一致）。未知のキーは
    「CLI フラグが存在しない＝発見」として素通しせず AssertionError で止める
    （黙って無視すると変換ミスが緑のまま埋もれるため）。`inplace` だけは
    上記の理由で明示的に許可して無視する。
    """
    overrides.pop("inplace", None)
    unknown = set(overrides) - set(_DEFAULTS)
    assert not unknown, f"run_argv: 未知のフィールド {sorted(unknown)}（CLI フラグ対応表を確認）"
    cfg = dict(_DEFAULTS)
    cfg.update(overrides)

    argv = ["run", str(book), task]
    if cfg["model"] is not None:
        argv += ["--model", str(cfg["model"])]
    if cfg["refs"] is not None:
        argv += ["--refs", str(cfg["refs"])]
    if cfg["helpers"] is not None:
        argv += ["--helpers", str(cfg["helpers"])]
    argv += ["--repair", str(cfg["repair"])]
    argv += ["--temperature", str(cfg["temperature"])]
    if cfg["dry"]:
        argv.append("--dry")
    if cfg["json"]:
        argv.append("--json")
    argv += ["--timeout", str(cfg["timeout"])]
    if cfg["ask"]:
        argv.append("--ask")
    if cfg["keep_backups"] is not None:
        argv += ["--keep-backups", str(cfg["keep_backups"])]
    if cfg["values"]:
        argv.append("--values")
    if cfg["header_row"] is not None:
        argv += ["--header-row", str(cfg["header_row"])]
    if cfg["sheet"] is not None:
        argv += ["--sheet", str(cfg["sheet"])]
    if cfg["accept_loss"]:
        argv.append("--accept-loss")
    if cfg["copy"]:
        argv.append("--copy")
    if cfg["overwrite"]:
        argv.append("--overwrite")
    if cfg["allow_freeform"]:
        argv.append("--allow-freeform")
    return argv
