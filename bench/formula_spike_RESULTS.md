# 式化スパイク結果

| ケース | 保存された式 | キャッシュ値 | 判定 |
|---|---|---|---|
| P1 D2 (=B2*C2) | `=B2*C2` | 360 | ✓ (期待 360) |
| P1 D3 | `=B3*C3` | 400 | ✓ (期待 400) |
| P1 D4 | `=B4*C4` | 300 | ✓ (期待 300) |
| P2 E2 (=SUM) | `=SUM(D2:D4)` | 1060 | ✓ (期待 1060) |
| P3 F2 (comma+bang) | `=VLOOKUP(A2,単価表 A2:B4,2,0)` | #VALUE! | × (期待 120) |
| P3 F3 (semi+bang) | `=VLOOKUP(A3,単価表 A2:B4,2,0)` | #NAME? | × (期待 80) |
| P3 G2 (comma+dot) | `=VLOOKUP(A2,単価表!A2:B4,2,0)` | #VALUE! | × (期待 120) |
| P3 G3 (semi+dot) | `=VLOOKUP(A3,単価表!A2:B4,2,0)` | 80 | ✓ (期待 80) |

## INDEX 挿入耐性式（第二次スパイク・`formula_spike_work2/`）

`APPEND_TOTAL` の合計式を「挿入しても追従する」形にできるか（設計 B）。
`=SUM(D2:INDEX(D:D;ROW()-1))*1.1` を setFormula の区切り(セミコロン/カンマ)違いで
2通り書いて実測（`formula_spike_work2/src/Gen.bas`・元表は D2:D4=360/400/300、
期待合計 (360+400+300)*1.1=1166）。

| 変種 | setFormula に渡した式 | 保存後の式 | キャッシュ値 | 判定 |
|---|---|---|---|---|
| セミコロン (LO方言) | `=SUM(D2:INDEX(D:D;ROW()-1))*1.1` | `=SUM(D2:INDEX(D:D,ROW()-1))*1.1` | 1166 | ✓ (期待 1166) |
| カンマ (Excel方言) | `=SUM(E2:INDEX(E:E,ROW()-1))*1.1` | `=SUM(E2:INDEX(E:E,ROW()-1))*1.1` | `#VALUE!` | × |

結論: P3 の VLOOKUP と同じ罠 — `setFormula` に渡す時点では **LO 方言（セミコロン）で
書かないと動かない**。保存後は両方ともカンマ区切りの Excel 方言に変換されるので、
`openpyxl` で読む事後条件（`check_append_total`）は保存後のカンマ形と照合すればよい。
`codegen_dsl` は前者（セミコロン）で `setFormula` を呼ぶ。
