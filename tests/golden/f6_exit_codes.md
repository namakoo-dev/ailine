# C1-F6: `ailine run` の終了コード表

★ 本番コード(ailine.py)は変更していない。この表は既存コードの実測（grep + 実行）から
組んだドキュメントで、`tests/test_golden_exit_codes.py` が表の主張を1つずつ生の関数呼び出し
で裏取りする（表だけ書いて終わりにしない）。

## 表

| code | 意味 | 発生箇所（関数） | 裏取りテスト |
|---|---|---|---|
| 0 | 成功（DSL/plan/freeform（複合計画の語彙外段）いずれかの達成。「機械保証なし」の警告つき成功や、doctor/history/vocab の正常終了も含む。★M2: `run <フォルダ>` の抽出集約は一致 0 行でも 0＝分母つきで「0 行」を明示して成功。★ freeform 最終決定(2026-08-21): 単発の語彙外はもう成功しない（`cmd_run_freeform` は廃止・0を返す経路から消えた）ので、単発 op の成功は `cmd_run_dsl` のみ） | `cmd_run_dsl` / `cmd_run_plan` / ★`cmd_run_folder` / `cmd_doctor` / `cmd_history` / `cmd_vocab` / `cmd_stop` 他 | `test_exit_0_success` |
| 1 | 汎用の失敗（事後条件 fail・y/N 確認で拒否・doctor の任意チェック失敗・restore/undo の対象無し・★W11 undo が履歴の端＝最も古い状態に着いた 等） | `_confirm_overwrite_or_gate`（対話 no）/ `_confirm_freeform_apply`（対話 no）/ `cmd_run_dsl` 等（事後条件未達成）/ ★挙動変更#3 `_sheet_conflict_gate`（3択の③「やめる」）・`_preview_and_run_on_alternative_sheet`（②のプレビュー後に N）/ ★W11 `cmd_undo`・`cmd_restore`（`NoOlderBackupError` を捕捉）/ ★M2 `cmd_run_folder`（書いた直後の独立読み検算が破れた＝出力は書かずに捨てる。stack の 5 は持ち込まない）・`cmd_undo`（フォルダには戻す対象が無い） | `test_exit_1_generic_failure` / `test_exit_1_undo_at_the_oldest_generation` / `tests/test_sheet_conflict.py::test_choice_3_stops_without_doing_anything` |
| **2** | **★ 欠番。ailine.py 自身はこのコードを一度も使わない。** | （下の「なぜ2が欠番か」参照） | `test_exit_2_is_argparse_reserved_not_ailine_own` |
| 3 | CLARIFY（見出し行推定の自信不足、または翻訳が確認質問を返した）／★ freeform 最終決定(2026-08-21): 単発の語彙外(FREEFORM/OUT_OF_VOCAB)も同じ exit 3（`cmd_refuse_vocab_miss` ── 既存の CLARIFY 系に合わせた。理由・vocab_miss 記録・次の手を言って即座に断る。旧 `cmd_run_freeform` の生成→適用ループ・exit 8 の関所は単発からは消えた）／★M2: `cmd_run_folder` は exit=3 を返す経路が6箇所あり、うち **`_run_folder_refuse` を経由するのは2箇所だけ**（op非対応・複数段の計画）。残り4箇所（CLARIFY passthrough・cmp不正・col不明・value読み取り不能/数値変換不能）は `_run_folder_refuse` を経由しない直接の `print(...); return 3` （★ review3#4: 以前の版はここを「`_run_folder_refuse`・OP_META の folder 宣言が唯一の根拠」と一枚岩に書いていたが実装・テストと一致しなかった＝言い過ぎ） | `_cmd_run_body`（見出し行推定）/ `cmd_run_dsl` / `cmd_refuse_vocab_miss`（★ 単発の語彙外・旧 `cmd_run_freeform` の後継）/ `cmd_run_plan`（翻訳結果が CLARIFY）/ ★`cmd_run_folder`（`_run_folder_refuse` 経由: op非対応・複数段計画／直接 return 3: CLARIFY passthrough・cmp不正・col不明・value不能） | `test_exit_3_clarify`（単一ブックの見出し行推定 CLARIFY のみ・`cmd_run_folder` には触れない）/ `tests/test_freeform_out_only.py`（★ 単発の語彙外の断り・3本）/ `tests/test_run_folder.py::test_unsupported_op_on_folder_is_refused_by_name`（`_run_folder_refuse`・単一op拒否）/ `tests/test_run_folder.py::test_folder_condition_validation_paths_exit_3`（残り5経路をパラメタライズで裏取り: CLARIFY passthrough・複数段計画・cmp・col・value×2） |
| 4 | 往復忠実度ゲート（LibreOffice 往復で失われる飾りを検出・`--accept-loss`/`--copy` 未指定） | `_cmd_run_body` | `test_exit_4_fidelity_gate` |
| 5 | Excel ロック検出（同フォルダの `~$` ロックファイル） | `_cmd_run_body`（`check_excel_lock`） | `test_exit_5_excel_lock` |
| 6 | グローバル run ロック取得失敗（別プロセスが `ailine run` 実行中） | `cmd_run`（`acquire_run_lock`） | `test_exit_6_run_lock_busy` |
| 7 | 破壊の関所・非対話で確認できない（既存データを持つ列への上書き）／★M2: 出力先に人のファイル（または ailine の別コマンドの出力）がある（`run` には `--overwrite` が無いので、退けるか消してもらう） | `_confirm_overwrite_or_gate`（`EOFError`）/ ★`cmd_run_folder`（書き込みの関所・移す直前にも再判定） | `test_exit_7_overwrite_gate_noninteractive` |
| 8 | 自由生成の関所・非対話で確認できない（機械検証できないコードの適用） | `_confirm_freeform_apply`（`EOFError`） | `test_exit_8_freeform_gate_noninteractive` |

## なぜ 2 が欠番か（調査結果）

**結論: `ailine.py` 自身のコードではなく、Python 標準ライブラリ `argparse` が予約している。**

`argparse.ArgumentParser.error()`（CPython 標準ライブラリの実装、`ailine.py` には実装が無い）は、
不正な引数・未知のサブコマンド等で呼ばれると `self.exit(2, ...)` を無条件に呼ぶ
（標準ライブラリのソース: `Lib/argparse.py` の `ArgumentParser.error`）。`build_parser()`
（ailine.py 5537行目）は `argparse.ArgumentParser` を素の設定（`prog`/`description`/
`add_subparsers(required=True)` のみ）で使っており、`error()` をオーバーライドしていない
（`grep -n "ArgumentParser(\|def error"` で ailine.py 内に override が無いことを確認済み）。
そのため:

- `python ailine.py --bogus-flag` → `exit=2`
- `python ailine.py badsubcommand` → `exit=2`

が両方とも実測で確認できる（`test_exit_2_is_argparse_reserved_not_ailine_own` が同じことを
`SystemExit` 捕捉で再現する）。

`ailine.py` 自身の `return`/`sys.exit` を全数 grep した結果（`grep -n "return [0-9]\|sys.exit"`）、
使われているのは `{0, 1, 3, 4, 5, 6, 7, 8}` のみで `2` は一度も出てこない。これは
「2 番を明示的に避けている」というより、**argparse が既に 2 を『コマンドライン自体が
不正』の意味で使っているため、ailine 独自の終了コード体系（0=成功、1=汎用失敗、
3以降=各種の関所/確認）を設計した時点で 2 と衝突させないよう暗黙に空けている**、
というのが実態に近い（コード中に「2 を避ける」という明示コメントは無く、単に
argparse の既定動作を上書きしていないだけ）。★ 他のレビューでも理由が見つからなかった
という前提の追記どおり、コード内にこれを明言する一次資料（コメント・commit メッセージ）
は見つからなかった。上記は grep/実行による外形的な調査結果であり、設計者の意図表明
そのものではない。

## 更新の作法

`ailine.py` の終了コードを追加/変更した場合は、この表と
`tests/test_golden_exit_codes.py` の両方を人力で更新する（このファイルは
`assert_golden_*` を通さない手書きドキュメントなので `AILINE_REGEN_GOLDEN` の対象外。
表を直したら必ず対応するテストも直し、`python -m pytest tests/test_golden_exit_codes.py -q`
が通ることを確認してから commit する）。
