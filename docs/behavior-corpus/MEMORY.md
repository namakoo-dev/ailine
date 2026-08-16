# ailine 挙動コーパス（B 棚・建てながら書く方式）

**方式宣言（2026-08-14 開始）**: このコーパスは製品開発と**並走**で書く。
basrun-ts / ailine-ts の移行実験では完成後に遡って掘った（考古学）が、ここでは
挙動が確定した日にノードを書く（同時代史）。目的は二つ:
①将来の言語移行の費用をほぼゼロにする ②「なぜこの挙動か」の一次資料を散逸させない。

書き方の約束（移行実験 #1/#2 で確立した流儀）:
- 1 ノード = 1 挙動単位。言語中立（Python の実装詳細でなく、挙動と意図を書く）
- 出典行を必ず付ける（commit・検証ログ・監査記録のどれか）
- 渡河分類: あ=そのまま渡る / い=構造は渡るが再設計要 / う=プラットフォーム結合

## ノード索引

- [noop-guard-normalization](nodes/noop-guard-normalization.md) — 適用前の LibreOffice 正規化パス。初回保存の行高実体化が no-op ガードを盲目にする問題への恒久対処
- [apply-timeout](nodes/apply-timeout.md) — 適用タイムアウト既定 180s・PID kill・★固まった LO は自動復旧しない既知の限界つき
- [run-history](nodes/run-history.md) — 実行履歴 = ヘルパ昇格経済学の需要センサー（失敗種別を構造化記録）
- [doctor](nodes/doctor.md) — セットアップ診断 7 項目。買った人の最初の 5 分を守る器官

- [dsl-pipeline](nodes/dsl-pipeline.md) — 中間命令言語の二段構え（翻訳→決定論 codegen→事後条件・列名を正とする・段階的劣化）
- [plan-execution](nodes/plan-execution.md) — 複合依頼の計画実行・黙落の構造的禁止・「言い切る範囲=検証した範囲」の正直バナー
- [header-detection-structdump](nodes/header-detection-structdump.md) — 見出し行検出と StructDump（三層が同じ推定を共有・CLARIFY 退避・gotoEndOfUsedArea(True) の罠・子見出し先頭空セルの親行フォールバック）
- [formula-dialect-conversion](nodes/formula-dialect-conversion.md) — COMPUTE_COLUMN/APPEND_TOTAL の式化（setFormula は LO 方言のセミコロン区切りのみ・保存時 Excel 方言(カンマ)へ自動変換・挿入耐性 SUM・二層事後条件）
- [empty-verification-ban](nodes/empty-verification-ban.md) — 空虚な検証合格の禁止（検証対象 0 件の ✓ を構造的に禁止・MAX_ROWS 切詰時の経路別の正直な注記）
- [progress-and-diff-humanization](nodes/progress-and-diff-humanization.md) — 進捗表示（stderr）と差分の人間可読化（生 tuple の追放）
- [destruction-gate-declarative](nodes/destruction-gate-declarative.md) — 破壊の関所と OP_WRITE_TARGET の宣言駆動（op ごとの if でなく宣言を読む・全 op 宣言の番人テスト）
- [verification-scope-honesty](nodes/verification-scope-honesty.md) — 検証が主張する範囲（「計画どおり」≠「依頼どおり」・複合計画にも解釈行・決定論警告）
- [formula-readback-duality](nodes/formula-readback-duality.md) — 式で書いた列を読み戻す二重性（openpyxl 式ビュー/値ビュー・行独立は部分採点、全行またぎは打ち切り）
- [freeform-gate-helper-sweep](nodes/freeform-gate-helper-sweep.md) — 自由生成の関所とヘルパ総なめの機械検出（閾値はプロンプトでなく Call 数の決定論検出）
- [lo-basic-native-formatting-quirks](nodes/lo-basic-native-formatting-quirks.md) — 太字は CharWeightAsian でしか効かない・PivotSum の再描画癖と SummaryTable の使い分け

保留: 無し（進捗表示/差分表示の人間化は GUI 非導入方針の確定によりノード化済み — progress-and-diff-humanization 参照）

ノード化待ち: 無し（2026-08-16 時点で上記 9 ノードへ全て収載済み）
