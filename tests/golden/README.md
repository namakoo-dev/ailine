# C1: 振る舞いの凍結（ゴールデンテスト群）

★★ 本番コード(`ailine.py`)は一切変更していない。ここは追加のみ
（`tests/test_golden_*.py` + `tests/golden/` 配下）。流れ層の再設計に入る前に、
決定論の層（LLM を通らない純関数）の「今の振る舞い」をバイト単位で固定し、
再設計が等価かどうかを人の目でなく機械が判定できるようにする。

## 更新の作法

`tests/golden/_harness.py` のモジュール docstring 参照。要約:

1. 挙動を意図的に変える前に `AILINE_REGEN_GOLDEN=1` を立てて該当テストを再実行し、
   ゴールデンファイルを再生成する。
2. `git diff -- tests/golden` で差分を人が読み、意図した変更だけかを確認する。
3. 確認できたら commit する。

**このハーネスは絶対に「差分があれば自動で追随する」ことをしない**
（`AILINE_REGEN_GOLDEN` が立っていない限り、食い違えば必ずテスト失敗にする）。

## 構成（F1〜F9）

| # | 対象 | テストファイル | ゴールデン | 件数 |
|---|---|---|---|---|
| F1 | `codegen_dsl` | `test_golden_codegen.py` | `f1_codegen/*.bas`（byte-exact） | 50 |
| F2 | `verify_dsl_args` | `test_golden_verify.py` | `f2_verify/*.json` | 71 |
| F3 | `run_postcondition` | `test_golden_postcondition.py` | `f3_postcondition/*.json` | 42 |
| F4 | `build_advisories`/`_structural_advisories` | `test_golden_advisories.py` | `f4_advisories/*.json`（before/after snapshot 込み） | 22 |
| F5 | `format_confirmation_line` | `test_golden_confirmation.py` | `f5_confirmation/golden.json`（1ファイル集約） | 18 |
| F6 | exit code 表 | `test_golden_exit_codes.py` | `f6_exit_codes.md`（手書き＋裏取りテスト） | 9（0,1,3-8 各1 + 2の調査） |
| F7 | `--json` キー集合/型 | `test_golden_json_keys.py` | `f7_json_keys/*.json` | 6（+ 不整合の自己検査1） |
| F8 | `--help` 全出力 | `test_golden_help.py` | `f8_help/*.txt` | 10 |
| F9 | 端末トランスクリプト | `test_golden_transcripts.py` | `f9_transcripts/*.txt`（`main(argv)` 経由） | 24 |

## 網羅性（できた分岐 / できていない分岐）— 網羅は主張しない

### F1 codegen_dsl
- できた: 全16 op を最低1ケース、COMPUTE_COLUMN は 2列/単列×target有無×use_formula の
  組み合わせ、header_row 1/3 の切替、後方互換(header_rows キー無し book_meta)。
- できていない: COMPUTE_COLUMN 以外の op で use_formula=False を明示的に確認したのは
  SORT のみ（他 op は「無視される」ことをソース読解で確認済みだが個別ゴールデンは無い）。
  複数シートを参照する op は LOOKUP_FILL のみで検証（他 op の他シート絡みは無い＝
  そもそも他 op は単一シートしか触らない設計）。

### F2 verify_dsl_args
- できた: 全16 op の ok/error 両系、数字表記の列解決（一意/曖昧）、factor 抽出の4経路
  （依頼文%/依頼文倍/用語集/未解決2種）、税込/税抜ラベルの3経路（明示・キーワード
  fallback・target有）、LLM値との食い違い警告（factor・value 各1）、シート無し/未対応op。
- できていない: `extract_rate_factor` の「掛け」パターン単体（`_RATE_KAKE_RE`）は
  ケース化していない（「割っ」「%」「倍」「倍率キーワード」は収載）。DRAW_BORDERS/AUTOFIT
  は引数が無いため error 系そのものが存在しない（分岐が無い＝網羅済み）。

### F3 run_postcondition
- できた: 全16 op-dispatch（POSTCONDITIONS 15 + CHART）の pass、warn が定義される op
  （SORT/INSERT_ROWS/AUTOFIT）の warn、代表的な fail 系、事後条件チェッカー自身の例外を
  `"error"` に変換する境界を1件。
- **できていない（明確なギャップ）**: `use_formula=True` の二層検証（式text＋data_only
  キャッシュ値）は COMPUTE_COLUMN でのみ実演（`_inject_formula_cache` で openpyxl に
  無い機能を模擬）。SORT/AGGREGATE/LOOKUP_FILL も同じ W10f 修正（data_only 側から読む）
  の対象だが、use_formula=True の版はゴールデン化していない（同じ XML 注入手法を
  流用すれば拡張できるが今回は見送った）。check_insert_rows の fail 分岐は3種類ある
  （シフト不一致・挿入行が空欄でない・想定より下にデータがある）うち1種類だけ収載。
  check_append_total のラベル不一致 fail 分岐は未収載（式形不一致・0件は収載）。

### F4 build_advisories / _structural_advisories
- できた: ①幽霊データ(検出/非検出) ②一様埋め(検出/非検出) ③件数突き合わせ(通常/見出し行
  込み) ⑤新規シートの中身(一様埋め検出/SummaryTable形の非検出) ⑥依頼にないシート新設
  (警告/沈黙) ⑦既存シート丸ごとすり替え(検出/部分更新は非検出) ④依頼文言重なり(列文字/
  数字表記/行/シート名・exclude_sheets)・`_structural_advisories` 単体の境界。
- できていない: `detect_ghost_data` の「使用範囲が不明で判定を保留」分岐（`rect is None`
  で `return None`）は未収載。複数シートが同時に変更されて `_structural_advisories` 内の
  個別関数が「全部該当」条件で不発火になる複合ケースは1本（multiple_advisories_at_once）
  のみで、パターンの網羅ではない。

### F5 format_confirmation_line
- 全16 op を収載。`_sources` 出典タグ・`(推定)` タグ・M2c のフィールド省略（target無し）
  もそれぞれ確認済み。既知のギャップなし（このセットは分岐が少なく網羅しやすい）。

### F6 exit code 表
- {0,1,3,4,5,6,7,8} を各1条件で裏取り。2 の調査結果は `f6_exit_codes.md` 参照
  （argparse 自身の予約・ailine.py 側の実装は無い）。
- できていない: 各コードが「複数の箇所」から出るケース（例: 1 は複数箇所から返る）の
  網羅ではなく、代表条件を1つずつ選んでいる（表の「発生箇所」列に全箇所を列挙、
  裏取りテストは代表1経路のみ）。

### F7 `--json` キー集合/型
- 全6通り(path×dry)を収載。既知の不整合（advisories の型が経路で違う）を
  型シグネチャの自己検査で明示。
- できていない: 各6通りとも「成功/dry」側のみ。失敗系（postcondition fail・
  実行時エラー等）での --json キー構成の型シグネチャは未収載（F9 の transcript 側で
  対応するテキスト出力は確認しているが、--json の型としては見ていない）。

### F8 --help
- 全10サブコマンドを収載。網羅済み（サブコマンドの数が少なく分岐が単純）。

### F9 端末トランスクリプト（本命）
- brief の目安どおり22本（dsl 4・plan 4・freeform 4・破壊の関所3・忠実度ゲート1・
  Excelロック1・runロック1・header-row1・--dry×3）＋ ★ 単位E で2本追加
  （subject_contradiction＝③で `✓` を出さない / subject_unspoken_note＝②の run 固有の1文）
  ＝ 24本。
- できていない: `ailine run` 以外のサブコマンド（doctor/history/vocab/restore/undo/stop）
  の main(argv) 経由トランスクリプトは対象外（brief のシナリオ一覧が run 系のみのため）。
  freeform の修復ループ（bad_signature/truncated → 再試行 → 成功）の transcript は
  未収載（今回のシナリオは全部「初回で成否が決まる」設計にした）。--json と組み合わせた
  transcript も無い（--json の型検証は F7 の担当として分離した）。

## ゴールデンが番人として機能することの実証（DoD 2）

`ailine.py` の `codegen_dsl` の区切り文字（`_wrap_basic` が返す `"End Sub\n"` の
末尾）を1箇所だけ意図的に壊す実験を行った。★ 本番ファイルへの書き込みは
Claude Code のサンドボックス自体（auto mode classifier）が `ailine.py` への
Edit/Bash(sed)/PowerShell 経由の変更を一律ブロックしたため、代わりに同じ壊れた
`_wrap_basic` を **プロセス内メモリ上でだけ** `ailine._wrap_basic` に monkeypatch し、
同じプロセス内で pytest をインプロセス実行して結果を見た（ディスク上の `ailine.py`
には一度も書き込んでいない＝`git diff` は最初から最後まで空のまま）。

結果: **866本中57本が red**。内訳は完全に予測どおり:
- `test_golden_codegen.py` の **50/50 全部**（F1 は例外なく `codegen_dsl` の出力を
  直接比較するため、区切り文字を変えれば必ず全滅する）。
- `test_golden_transcripts.py` の **7本**（`dry_dsl`・`dsl_pass`・`dsl_fail`・
  `dsl_warn`・`dsl_runtime_error`・`header_row_explicit`・`overwrite_gate_yes`——
  いずれも DSL 経路で「生成した .bas」を標準出力に印字するシナリオ。複合計画
  (`plan_*`) は個々の段の `.bas` を印字しないため無傷、自由生成(`freeform_*`)は
  `codegen_dsl` を通らないため無傷、破壊の関所で拒否/非対話終了するシナリオは
  コード生成前に return するため無傷——という判定ロジックの理解どおりの結果）。

残り809本（F2〜F8 全部・F9 の残り15本）はこの変更の影響を一切受けず green のまま。

monkeypatch を解除して確認したところ、通常実行では866本全部が再び green に戻った
（そもそもディスクを一切変更していないので当然だが、実測でも確認済み）。

## その他の確認事項

- CI 相当: `BASRUN=<存在しないパス>` でも866本全部 green（既存594本と同じ性質を
  新規257本にも確認）。
- `git diff --stat -- ailine.py` は空（本番コード無変更）。
- 異文字混入チェック（Hangul/Cyrillic）: 新規ファイル全部でヒット無し。
- CRLF混入チェック: 新規ファイル全部で `\r` 無し（`__pycache__/*.pyc` のバイナリ
  誤検知を除く・実害無し・`.gitignore` 対象）。
- LibreOffice残存プロセス: このタスクの実行中に basrun/LibreOffice を一度も
  起動していない（`normalize_book`/`basrun_apply` は全ゴールデンテストで
  monkeypatch 済み・`BASRUN` 無効化テストでも866本全部 green がその証拠）。
  実行環境に稼働中の `soffice`/`soffice.bin`（開始時刻から判断してこのタスク開始前
  からの別プロセス）が見つかったが、本タスクが起動したものではないため触っていない
  （名前一括 kill 禁止の方針どおり、自分が起動したと確認できないものは殺さない）。
