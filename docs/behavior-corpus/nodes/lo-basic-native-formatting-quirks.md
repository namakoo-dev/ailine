# lo-basic-native-formatting-quirks — LibreOffice Basic のネイティブ書式付けの癖

**分類: う（LibreOffice UNO API のプロパティ名・DataPilot 実装固有の再描画挙動に結合）**

## 挙動

**太字は `CharWeightAsian` でしか効かない**: セルへ太字を当てる際、`CharWeight`
だけを設定しても**日本語（アジア言語）テキストは太らない**。`CharWeight` /
`CharWeightAsian` / `CharWeightComplex` の 3 プロパティを同時に
`com.sun.star.awt.FontWeight.BOLD` へ設定して初めて、日本語を含むセルが太字に
なる（数値セルも壊さず太字にできる — text cursor 経由は数値を文字列化するので
使わない）。ヘルパ層はセルへ直接プロパティを当てる方式（native 書き）で、xlsx に
太字が書き出せることを openpyxl 読み戻し・描画の両方で実測確認済み。

**`PivotSum`（本物の DataPilot）は再描画で書式が撥ねる**: `PivotSum` が作る
DataPilot は本物のピボットテーブルだが、LibreOffice が開くたび**再描画して
セル書式を撥ねる**（罫線・カンマ区切りが表示されない）。書式つきの見栄えのする
集計表が欲しい場合は、`SummaryTable` を使う — こちらは普通のセルへ直接書くため、
格子罫線・カンマ・中央揃え・太字が native のままそのまま残る（描画で確認済み）。
製品側はこの使い分けを、DSL の確認行と結果表示の両方で一言添えて
（`PIVOT_CAVEAT = "書式なしの素の表になります。書式つきは『集計表』"`）促す。
DSL 語彙では「ピボット」と明示された依頼だけが `PIVOT`（DataPilot・
`PivotSum` ヘルパ）に、それ以外の「集計/まとめ/小計」的な依頼は既定で
`AGGREGATE`（`SummaryTable` ヘルパ）に翻訳される。

## なぜ（一次資料つき）

- 太字の日本語不発現は当初「LibreOffice Basic + UNO は学習データが薄く、珍しい
  操作は外しやすい」という一般論から**「環境不可」と誤断**されていた。実際には
  `CharWeight`+`CharWeightAsian` の native 書きで解決済みという実測結果が
  README の「正直な限界」節に残っている（誤断の訂正そのものが一次資料）
- `PivotSum`/`SummaryTable` の使い分けは、DataPilot が LibreOffice の再描画時に
  書式を失うという UNO 実装固有の癖への対処として `helpers/AiLineHelpers.bas`
  のコメントに実測記録されている。製品側の `PIVOT_CAVEAT` 文言・DSL 語彙分岐
  （「ピボット」の語の有無）はこの癖を前提に設計されている

- 出典: `helpers/AiLineHelpers.bas`（`BoldRange`/`SummaryTable`/`PivotSum` の
  Sub 直前コメント）・README「正直な限界」節（太字の誤断の訂正）・ailine.py
  `PIVOT_CAVEAT` 定義・commit 03e3d11 周辺（W6・W9 で「ピボット」の語の有無による
  DSL 分岐を強化）

## 検証挙動（GOLDEN 相当）

| 入力 | 期待 |
|---|---|
| 日本語の見出し行を太字にする依頼 | `CharWeightAsian` を含む 3 プロパティ設定で実際に太る（描画確認済み） |
| 「ピボットで部門ごとに集計」 | `PivotSum`（本物の DataPilot）・確認行に PIVOT_CAVEAT が付く |
| 「部門ごとにまとめて」（「ピボット」の語なし） | `SummaryTable`（罫線・カンマ・太字つきの見栄えする表） |
| PivotSum の出力を LibreOffice で再度開く | 再描画で罫線・カンマが撥ねることがある（既知の癖・SummaryTable は撥ねない） |
