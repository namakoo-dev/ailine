# ailine

**自然言語のタスクを、ローカル LLM が LibreOffice Basic に書き起こし、[basrun](https://github.com/namakoo-dev/basrun) で文書に適用し、★ 効果を読み戻して検証する。**

「Excel の見積に金額と合計を入れて」を、平文の Basic に翻訳し、`.xlsx` に適用する。
書いたコードは平文で残り、`git diff` で読める。外部にデータは送らない。

> **状態: 骨格（PoC）。** 中核のパイプライン（生成 → 適用 → 検証 → 修復）は動き、
> 純ロジックのユニットテスト 20 件は緑。参照ライブラリの拡充と、実機 LibreOffice を
> 通した自動の通し試験はこれから。実運用の前に下の「限界」を必ず読むこと。

> **姉妹版**: [ailine-ts](https://github.com/namakoo-dev/ailine-ts) — この repo の Python
> ソースを一切読まずに、挙動コーパスだけから実装された TypeScript 移植（移行手法実験の
> 成果物・テスト 61 件）。本 repo が正典で、両者は独立に保守される。

---

## 何をするか

```
    自然言語のタスク
        │  ローカル LLM (ollama, 既定 qwen2.5-coder:7b)
        ▼
    LibreOffice Basic  Sub Run(oDoc As Object)   ← 平文。レビューできる
        │  basrun apply（LibreOffice を headless で駆動）
        ▼
    文書に適用（★ 既定=原本に直接。検品ゲート+世代バックアップ+undo の3重安全網つき。
                --copy で従来のコピー .out 方式も選べる）
        │  適用の前後で snapshot して差分を取る
        ▼
    ★ 変化したか？ ── しなければ「何もしていない」とみなし修復に回す
```

## 使い方

```bash
# 生成 → コピーに適用 → 変化を検証 → 差分を表示（原本は触らない）
python ailine.py run demo/sample.xlsx "各行の 列5 に 売上(列3) − 原価(列4) を入れる"

# 生成して見せるだけ（レビュー用。適用しない）
python ailine.py run demo/sample.xlsx "..." --dry

# ★ W8b-2: 既定で原本に直接反映する（反映前に自動でバックアップを作る）。
# 何もつけなくてよい
python ailine.py run demo/sample.xlsx "..."

# 原本には触らず <book>.out に結果を作りたいとき（旧既定・原本は無変更）
python ailine.py run demo/sample.xlsx "..." --copy

# LibreOffice 往復だけで失われる飾り（条件付き書式・図形・ピボット・VBA 等）を
# 検出したら、原本に触る前に申告して止まる（exit 4）。承知の上で続けるか
# （バックアップから ailine undo で戻せる）、.out に切り替えるか選ぶ
python ailine.py run demo/sample.xlsx "..." --accept-loss
python ailine.py run demo/sample.xlsx "..." --copy   # ゲートも走らせず原本に触らない

# 別のモデルに載せ替える（天井を上げたいとき）
python ailine.py run demo/sample.xlsx "..." --model qwen2.5-coder:32b

# 参照ライブラリ / ヘルパのディレクトリを差し替える（既定は ./refs, ./helpers）
python ailine.py run demo/sample.xlsx "..." --refs my_refs --helpers my_helpers

# 生成の温度・修復の最大回数・適用タイムアウト秒を調整する
python ailine.py run demo/sample.xlsx "..." --temperature 0.1 --repair 3 --timeout 60

# 見出しが何行目か機械が確信を持てず「？ 見出しが何行目か分かりません」で止まったとき、
# 見出し行(1起点)を明示して自動検出をスキップする
python ailine.py run demo/sample.xlsx "..." --header-row 3

# 結果を JSON でも出す（changes/advisories/out などを機械可読で受け取る）
python ailine.py run demo/sample.xlsx "..." --json

# 起動した LibreOffice を落とす
python ailine.py stop

# セットアップを診断する（python/openpyxl/ollama/モデル/LibreOffice/basrun/demo）
python ailine.py doctor

# 実行履歴を見る（新しい順。既定 10 件）
python ailine.py history --max 20

# 原本への反映前のバックアップから復元する（復元前の現状も自動で退避＝復元自体も可逆）。
# ailine undo が restore の昇格版（あと何回戻せるかを表示・restore は互換のため残す）
python ailine.py undo demo/sample.xlsx
python ailine.py undo demo/sample.xlsx --list   # 一覧だけ表示（復元しない）

# 用語集（税率等の取り決め値）に語を登録する。「税込み合計」等で率が本文にも
# 用語集にも無い場合、この形のコピペ可能な1行が CLARIFY のメッセージに出る
python ailine.py vocab add 消費税 1.1
python ailine.py vocab list
```

`~/.ailine/vocab.json` に平坦な `{"語": 値}` で保持する（グローバルのみ・ブック別上書きは
無い）。テンプレは `vocab.template.json`（値は空・古びた法定値を焼き込まない。値を埋めて
`~/.ailine/vocab.json` にコピーするか `vocab add` で登録する）。

試せる表を `demo/` に同梱している: `sample.xlsx`（商品×金額の一覧）・
`sales.xlsx`（部門×金額 — ピボット/集計表向き）・`lookup.xlsx`（明細＋単価表 — VLOOKUP 向き）。

## 設計判断（なぜこうしたか）

計測に基づく（`basrun_spike` 2026-08-10、qwen2.5-coder:7b）。

- **モデル非依存** — `--model` で差し替える。天井はモデルの大きさでなく
  「正しい参照例の供給 ＋ 効果の検証」で上げる。実測で **7B が正解例 1 本で
  苦手層（新シート・ソート・グラフ）を 0% → 67%** まで上げた（残ったソートの
  取り違えは、のちに下の「ヘルパ」方式で解決）。
- ★ **検証をループに（no-op ガード）** — 適用の前後で文書を snapshot し、
  値・数値書式・背景色・太字・**罫線・結合・列幅・行高・水平配置**・シート・グラフの変化を見る
  （構造や装飾だけの変更も取りこぼさない）。LibreOffice + LLM は
  **「実行時エラー無しで成功と報告し、実際は何もしない」**ことがある（もっともらしい
  UNO の幻覚）。変化ゼロなら失敗として修復に回す。
- ★ **既定で原本にそのまま反映（W8b-2）** — 「壊さない」の担保は、もう「コピーにしか
  書かない」ことではなく、3重の安全網でまかなう:
  ①**検品ゲート（往復忠実度ゲート）** — LibreOffice に一度通すだけ(何もしないマクロ)
  で失われる飾り（条件付き書式・入力規則・図形・ピボット・VBA・_rels）を検出したら、
  原本に触る前に申告して止まる（exit 4。`--accept-loss`/`--copy` で選ぶ・喪失ゼロなら
  無言）。②**世代バックアップ + `ailine undo`** — 反映のたびに自動でバックアップし、
  何段でも `ailine undo` で1段ずつ遡れる（`~/.ailine/backups/<フォルダのハッシュ>/`
  ＝同名ファイルが別フォルダにあっても取り違えない）。③**機械検証** — no-op ガード・
  事後条件チェッカーで「変化したか」「（DSL経路は）正しいか」を見る。原本には触らず
  従来どおりコピーで試したい時は `--copy` を使う（`<book>.out.xlsx` に生成・原本は無変更・
  ゲートも走らせない）。
- ★ **その他の安全機構（W8b）** — ①**Excel ロック検出**: 同フォルダの `~$`
  ロックファイルと書き込み可否を run の最初に見る。②**原子的な置換**: `os.replace()`
  で torn write（書きかけの状態が外から見える窓）を塞ぐ（失敗時は copy2 にフォールバック
  し、その旨を表示）。③**グローバル run ロック**: 基盤の LibreOffice が単一インスタンス
  前提のため、`ailine run` は同時に1本だけ（別プロセスが実行中なら待たせず即エラー）。
- **参照ライブラリ** — `refs/*.bas` を few-shot に供給。苦手な操作（並べ替え・
  グラフ・新シート）の**動作検証済みの正解例**を渡して補う。追加は `refs/` に置くだけ。
- ★ **達成検証層（M2a）** — 「✓ の下に隠れた失敗」（幽霊データ・無関係なすり替え・
  1 行の静かな欠落）への機械的対抗。差分の後に助言として表示する（ブロックはしない）。
  - **幽霊データ検出** — 変更セルの全部が原本の使用範囲外に集中している場合だけ
    「★ 疑わしい: 変更が元データの範囲外です（Z2:Z6）」。
  - **一様埋め検出** — 変化前が全部空欄・変化後が全部同一値（特に 0/空文字）の場合だけ
    「★ 疑わしい: 空欄への同一値の一括書き込みです（値 0 × 5 セル）」。
  - **件数の突き合わせ** — 変更が単一列に集中している場合、「列 C: データ 3 行のうち
    2 行を変更（1 行は未変更）」を添える（りんご欠落型を1秒で見えるように）。
  - **依頼文と変更範囲の重なりチェック** — タスク文言に「列Z」「行5」「シート名」等の
    明示的な言及があるのに、変更が全く重ならない場合だけ「★ 依頼で言及された『列Z』
    は存在しません/変更されていません」。言及ゼロのタスクでは何も言わない。
  - どれも保守的（両条件とも全セルがそれに該当した時だけ発火）＝誤検知回避優先。
- ★ **倍率(税率等)は LLM に計算させない** — 「消費税10%」を LLM に自分で 1.1 へ換算
  させると幻覚する（実測: 8% と誤ることがあった）。`APPEND_TOTAL` の倍率は
  ①依頼文の明示率（10%/1.1倍 等）を機械の正規表現で抽出 ②無ければ用語集
  （`vocab add`）を引く、の2段で**LLM を介さず**確定する。ラベルが「税込み」等
  なのに倍率が確定できない場合は断定せず確認を返す（税抜き金額に「税込み」の
  ラベルが付く恒真の誤りを防ぐ）。

## ★ 限界（正直に）

- **珍しい UNO 操作は外しやすい。** LibreOffice Basic + UNO は学習データが薄い。
  参照ライブラリで補う設計だが、載っていない操作は当たり率が落ちる。
- **no-op ガードが保証するのは「変化したこと」だけ。**「**正しいか**」は保証しない ──
  出力の差分を人が見て判断すること。ツールは必ず「差分を見て判断せよ」と促す。
  M2a の助言（★ 疑わしい 等）も同じ層 — 「変化」の機械保証であって、「依頼を
  達成したか」は助言＋人の確認が要る。
- **timeout kill 後に固まった LibreOffice** は、稀に手動で `python ailine.py stop`
  （または OS のタスクマネージャ）から止める必要がある場合がある。
- ★ **倍率の機械確定は DSL 経路(APPEND_TOTAL)に限る。** 依頼が `APPEND_TOTAL` として
  分類されなかった場合（列が曖昧・見慣れない書き方等で翻訳が `OUT_OF_VOCAB`/自由生成に
  退避した場合）、倍率の機械抽出・用語集・CLARIFY 番人は一切効かず、自由生成の LLM が
  税率を自分で仮定することがある（実機確認: 同じ「小計」列を持つ表でも "税込み合計を
  出して" のように依頼文だけでは対象列が一意に決まらない書き方だと自由生成に落ち、
  「消費税率は10%として計算」等をコードに直接書くことがあった）。
  出力の確認行が出ない（`■ ailine（AI が直接作成・機械保証なし）` と表示される）ことで
  気づける設計に加え、★ 生成コード中の率らしい数値リテラル（0.05〜0.2 / 1.05〜1.2 の
  小数）が依頼文にも用語集にも出典が無い場合は「★ 率らしい数値 (0.08) が依頼に無いのに
  使われています — 検算してください」を助言として添える（`scan_rate_literals`）。
  ただし機械的なブロックではなく、あくまで検算を促す助言止まり。
- **人が同ファイルを開いたまま実行した場合は未検証。** ロックや競合の扱いは
  基本の LibreOffice/OS の挙動に委ねている。
- **ローカル完結。** ollama と LibreOffice が要る。**外部にデータは送らない。**

## 必要なもの

- Python 3.10+ と `openpyxl`
- [ollama](https://ollama.com/) ＋ コード生成モデル（既定 `qwen2.5-coder:7b`）
- LibreOffice ＋ [basrun](https://github.com/namakoo-dev/basrun)（環境変数 `BASRUN` で場所を指定可）

## 参照ライブラリ

`refs/` に置いた `.bas` が few-shot として渡る。同梱:

| ファイル | 教える API |
|---|---|
| `01_value_format.bas` | セルの読み書き・四則・数値書式 |
| `02_new_sheet.bas` | 新シート作成（`insertNewByName`） |
| `05_cell_color.bas` | 条件付き背景色（★色は `&HRRGGBB&` の16進で。`RGB()` は VBASupport 下で BGR になり色が入れ替わる） |

**追加する参照は、必ず basrun で動作検証してから置くこと。** 動かない例は few-shot を毒する。

## ヘルパ（`helpers/`）— 難所は「呼ぶだけ」

参照例（few-shot）だけでは確度が上がりきらない操作がある。典型が**並べ替え**で、
7B は正しい例を見せても `ContainsHeader` の真偽を半分ほど滑らせた（判断ミスで、知識では直らない）。

そこで **arcane な操作は人が検証したヘルパに閉じ込め、モデルには `Call` で呼ぶだけ**させる。
`helpers/*.bas` は生成コードと同じライブラリに同梱され、`Sub Run` から呼べる。

| ヘルパ | モデルが書くのは | 隠している難所 |
|---|---|---|
| `SortByColumn(oDoc, col, ascending)` | `Call SortByColumn(oDoc, 1, False)` | 範囲検出・`SortFields`・`ContainsHeader=False` |
| `InsertBarChart(oDoc, valCol)` | `Call InsertBarChart(oDoc, 1)` | 見栄えのする棒グラフ。タイトル・横軸・系列色を見出しから自動導出（データラベルは付けず縦軸で読ませる清潔な既定。LO native の表現力を自前で引き出す。項目名は列0固定） |
| `MergeCells(oDoc, c1, r1, c2, r2)` | `Call MergeCells(oDoc, 0, 0, 1, 0)` | 範囲を渡さず単一セルに merge する誤りを封じる |
| `InsertRows(oDoc, atRow, count)` | `Call InsertRows(oDoc, 1, 1)` | `Rows.insertByIndex`・0起点の位置 |
| `DrawTableBorders(oDoc)` | `Call DrawTableBorders(oDoc)` | データ範囲を自動検出・`TableBorder2` の格子 |
| `AutoFitColumns(oDoc)` | `Call AutoFitColumns(oDoc)` | 使用列を自動検出・`OptimalWidth` |
| `AlignCenter(oDoc)` | `Call AlignCenter(oDoc)` | 表全体を中央揃え。セル配置は `HoriJustify`（`CharHorizontalAlignment` は段落用で効かず 7B が滑る罠を封じる） |
| `FormatThousands(oDoc, col)` | `Call FormatThousands(oDoc, 4)` | 指定列に3桁区切り `#,##0`。`queryKey` の -1 を `addNew` で拾い Locale を正しく構築（7B は addNew を落として滑る） |
| `VLookupFromTable(oDoc, keyCol, resultCol, lookupSheet)` | `Call VLookupFromTable(oDoc, 0, 2, "単価表")` | Basic 側で照合（数式 `=VLOOKUP` はこの経路で `#VALUE!`）。参照表は 列0=キー/列1=値 |
| `PivotSum(oDoc, groupCol, valueCol)` | `Call PivotSum(oDoc, 0, 1)` | 本物のピボット（DataPilot）を新「ピボット」シートに。分類×合計を自動。★LO は開くたび再描画してセル書式を撥ねる（罫線・カンマが出ない）＝清潔な表が欲しければ下の `SummaryTable` |
| `SummaryTable(oDoc, groupCol, valueCol)` | `Call SummaryTable(oDoc, 0, 1)` | 分類×合計を新「集計」シートに**普通の表**として出す（格子罫線・カンマ・中央揃え・見出し/合計太字を native で。DataPilot でないので全部残る） |
| `StyleBold(oDoc, c1, r1, c2, r2)` | `Call StyleBold(oDoc, 0, 0, 4, 0)` | native 太字。★`CharWeight`＋**`CharWeightAsian`**（日本語）＋`CharWeightComplex` をセルに直接。text cursor は数値を壊すので使わない |

ユーザは `SortByColumn` を知らなくてよい。**「金額で降順に並べ替えて」と自然文で頼むだけ**で、
モデルが列と向きを選んでヘルパを呼ぶ。難所に触れないので滑らない（実測で完全降順を確認）。

★ **W9: `InsertRows`/`DrawTableBorders`/`AutoFitColumns`/`PivotSum` は DSL 語彙にも昇格済み**
（`INSERT_ROWS`/`DRAW_BORDERS`/`AUTOFIT`/`PIVOT`）。「3行目の前に1行挿入して」「表にけい線を
引いて」「列幅を内容に合わせて」「部門ごとにピボットテーブルで集計して」のように頼むと、
自由生成(LLM任せ)でなく決定論の DSL パイプライン（②検証→③確認→④codegen→⑤適用→⑥事後条件）
で機械検証まで通る。「ピボット」と明示しない集計依頼（「まとめて」「小計」等）は引き続き
`AGGREGATE`（`SummaryTable`・書式つき）になる。

**ヘルパも必ず basrun で動作検証してから置くこと。** 呼び方は `Call 名前(引数)`
（括弧つきで `Call` 無しは LibreOffice Basic が誤動作する）。
