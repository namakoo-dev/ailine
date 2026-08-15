# B1 忠実度ベースライン（★ W8b 項目7）

`bench/fidelity_baseline.py` の実測結果。各検体を `ailine.normalize_book`（LibreOffice で一度・何もしない空マクロで開いて保存するだけ）に通し、`check_round_trip_fidelity` で LO 往復だけによる喪失を見る。

| 検体 | ファイル | 結果 | 内訳 |
|---|---|---|---|
| A title_rows | title_rows.xlsx | 喪失なし | - |
| B large(3000行) | large.xlsx | 喪失なし | - |
| C formulas | formulas.xlsx | 喪失なし | - |
| D merged_head | merged_head.xlsx | 喪失なし | - |
| E cf(条件付き書式) | cf.xlsx | 喪失なし | - |
| F datavalidation(入力規則) | datavalidation.xlsx | 喪失なし | - |
| G shapes(図形) | shapes.xlsx | 喪失あり | 図形/描画 1件 |
| H rows1500(1500行) | rows1500.xlsx | 喪失なし | - |

喪失あり: 1 / 測定できた検体 8（開けなかった/作れなかった検体: 0）