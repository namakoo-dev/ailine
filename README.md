# ailine

**自然言語のタスクを、ローカル LLM が LibreOffice Basic に書き起こし、[basrun](https://github.com/watasisaikou/basrun) で文書に適用し、★ 効果を読み戻して検証する。**

「Excel の見積に金額と合計を入れて」を、平文の Basic に翻訳し、`.xlsx` に適用する。
書いたコードは平文で残り、`git diff` で読める。外部にデータは送らない。

> **状態: 骨格（PoC）。** 中核のパイプライン（生成 → 適用 → 検証 → 修復）は動くが、
> 参照ライブラリの拡充・対応書式・テストはこれから。実運用の前に下の「限界」を必ず読むこと。

---

## 何をするか

```
    自然言語のタスク
        │  ローカル LLM (ollama, 既定 qwen2.5-coder:7b)
        ▼
    LibreOffice Basic  Sub Run(oDoc As Object)   ← 平文。レビューできる
        │  basrun apply（LibreOffice を headless で駆動）
        ▼
    文書に適用（★ 原本でなくコピー .out に）
        │  適用の前後で snapshot して差分を取る
        ▼
    ★ 変化したか？ ── しなければ「何もしていない」とみなし修復に回す
```

## 使い方

```bash
# 生成 → コピーに適用 → 変化を検証 → 差分を表示（原本は触らない）
python ailine.py run 見積.xlsx "各行の 列5 に 売上(列3) − 原価(列4) を入れる"

# 生成して見せるだけ（レビュー用。適用しない）
python ailine.py run 見積.xlsx "..." --dry

# 原本を上書きしてよいとき
python ailine.py run 見積.xlsx "..." --inplace

# 別のモデルに載せ替える（天井を上げたいとき）
python ailine.py run 見積.xlsx "..." --model qwen2.5-coder:32b

# 起動した LibreOffice を落とす
python ailine.py stop
```

## 設計判断（なぜこうしたか）

計測に基づく（`basrun_spike` 2026-08-10、qwen2.5-coder:7b）。

- **モデル非依存** — `--model` で差し替える。天井はモデルの大きさでなく
  「正しい参照例の供給 ＋ 効果の検証」で上げる。実測で **7B が正解例 1 本で
  苦手層（新シート・グラフ）を 0% → ほぼ解ける**まで上がった。
- ★ **検証をループに（no-op ガード）** — 適用の前後で文書を snapshot し、
  値・数値書式・背景色・太字・**罫線・結合・列幅・行高**・シート・グラフの変化を見る
  （構造や装飾だけの変更も取りこぼさない）。LibreOffice + LLM は
  **「実行時エラー無しで成功と報告し、実際は何もしない」**ことがある（もっともらしい
  UNO の幻覚）。変化ゼロなら失敗として修復に回す。
- **コピー安全** — 原本は触らず `<book>.out.xlsx` に適用する。壊さない。
- **参照ライブラリ** — `refs/*.bas` を few-shot に供給。苦手な操作（並べ替え・
  グラフ・新シート）の**動作検証済みの正解例**を渡して補う。追加は `refs/` に置くだけ。

## ★ 限界（正直に）

- **太字は「自作の道」で対応。** LibreOffice は CharWeight の太字を xlsx に書き出せない
  （store でも storeToURL でも。描画で二重確認）。そこで `StyleBold` ヘルパは Basic 側では
  何もせず、**ailine が basrun 適用後に openpyxl で太字を後付け**する。フォント色・サイズも
  同じ仕組みで足せる。値・数式・数値書式・背景色は Basic で直接動く。
- **珍しい UNO 操作は外しやすい。** LibreOffice Basic + UNO は学習データが薄い。
  参照ライブラリで補う設計だが、載っていない操作は当たり率が落ちる。
- **no-op ガードが保証するのは「変化したこと」だけ。**「**正しいか**」は保証しない ──
  出力の差分を人が見て判断すること。ツールは必ず「差分を見て判断せよ」と促す。
- **ローカル完結。** ollama と LibreOffice が要る。**外部にデータは送らない。**

## 必要なもの

- Python 3.10+ と `openpyxl`
- [ollama](https://ollama.com/) ＋ コード生成モデル（既定 `qwen2.5-coder:7b`）
- LibreOffice ＋ [basrun](https://github.com/watasisaikou/basrun)（環境変数 `BASRUN` で場所を指定可）

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
| `VLookupFromTable(oDoc, keyCol, resultCol, lookupSheet)` | `Call VLookupFromTable(oDoc, 0, 2, "単価表")` | Basic 側で照合（数式 `=VLOOKUP` はこの経路で `#VALUE!`）。参照表は 列0=キー/列1=値 |
| `PivotSum(oDoc, groupCol, valueCol)` | `Call PivotSum(oDoc, 0, 1)` | 本物のピボット（DataPilot）を新「ピボット」シートに。分類×合計を自動 |
| `StyleBold(oDoc, c1, r1, c2, r2)` | `Call StyleBold(oDoc, 0, 0, 4, 0)` | ★Basic では no-op。太字は **ailine が openpyxl で後付け**（LO は太字を xlsx に書けないため） |

ユーザは `SortByColumn` を知らなくてよい。**「金額で降順に並べ替えて」と自然文で頼むだけ**で、
モデルが列と向きを選んでヘルパを呼ぶ。難所に触れないので滑らない（実測で完全降順を確認）。

**ヘルパも必ず basrun で動作検証してから置くこと。** 呼び方は `Call 名前(引数)`
（括弧つきで `Call` 無しは LibreOffice Basic が誤動作する）。
