Option VBASupport 1
Option Explicit

' ────────────────────────────────────────────────────────────────
'  ailine の検証済みヘルパ集。arcane な UNO 操作を「呼ぶだけ」にする。
'  モデルはこれらを呼ぶだけ。中の難所（ソートの ContainsHeader 等）は触らせない。
'  ★ ここは人が検証して固定する。生成物ではない。
' ────────────────────────────────────────────────────────────────

' 1枚目シートのデータ行（見出し行 headerRow を除く）を、col 列で並べ替える。
' 範囲と見出し扱いは内部で自動処理する。呼び側は見出し行・列幅・列・向きを渡す。
'   headerRow  : 見出し行（0 起点。W3: StructDump が推定した実際の見出し行。物理1行目なら0）
'   lastCol    : 表の最終列（0 起点。★ W3: 多段見出しでは先頭列が見出し行で空欄になり得る
'                （例: 親見出し行にのみ商品名があり子見出し行の同じ列は空）ため、Basic で
'                見出し行を走査して求めず、接地済みの列数から Python 側(codegen_dsl)が
'                決定論的に渡す）
'   col        : 並べ替えの基準列（0 起点）
'   ascending  : True=昇順, False=降順
Sub SortByColumn(oDoc As Object, headerRow As Integer, lastCol As Integer, col As Integer, ascending As Boolean)
    Dim oSheet As Object, oRange As Object
    Dim lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（A 列を見出しの直下から走査）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    If lastRow < headerRow + 1 Then Exit Sub   ' データが無い

    ' 見出しの直下..lastRow（見出しを含めない）を範囲にする
    oRange = oSheet.getCellRangeByPosition(0, headerRow + 1, lastCol, lastRow)

    Dim aFields(0) As New com.sun.star.util.SortField
    aFields(0).Field = col                 ' 範囲は列0起点なので絶対列＝相対列
    aFields(0).SortAscending = ascending

    Dim aDesc(1) As New com.sun.star.beans.PropertyValue
    aDesc(0).Name = "SortFields"
    aDesc(0).Value = aFields()
    aDesc(1).Name = "ContainsHeader"       ' ★ 範囲に見出しを含めていないので必ず False
    aDesc(1).Value = False

    oRange.sort(aDesc())
End Sub


' 1枚目シートに「見栄えのする」グラフを1つ挿入する（棒/折れ線/円）。範囲もスタイルも内部で組む。
' ★ グラフ段: 項目名の列(catCol)は呼び側が指定する（旧 InsertBarChart は列0固定だったが、
'    非隣接列（例: 横軸=A・値=C）も選べるよう引数化）。
' ★ spike 実測: line/pie は addNewByName（棒と同じ2本の CellRangeAddress）の後に
'    createInstance("com.sun.star.chart.{Line,Pie}Diagram") + setDiagram の差し替えだけでよい。
' ★ タイトル・系列色を見出しから自動導出して styling する（棒/折れ線はさらに横軸タイトルも）
'    ＝ LibreOffice native チャートの表現力を自前で引き出す（外部依存なし・ours）。
'    line/pie にも棒と同等のスタイリングを適用する（商品の顔を揃える・spike の素の line/pie は
'    タイトル無しだった）。データラベルは付けない（値は縦軸/凡例で読める。全要素に数字を
'    振らない方が清潔＝プロの既定）。
' ★ 円グラフは「1系列を複数スライスに割る」形なので、スライスごとの区別はグラフ自体の
'    自動配色（visualColors）に任せる ── 棒/折れ線と同じ単色の FillColor を当てると
'    全スライスが同じ色になり、円グラフの用を成さなくなるため当てない。同じ理由で
'    横軸(カテゴリ軸)自体が無いので HasXAxisTitle も設定しない。凡例はスライスの
'    区別に要るため棒/折れ線と逆に立てる。
'   headerRow  : 見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   catCol     : 項目名にする列（0 起点）
'   valCol     : 値にする列（0 起点。例: 金額=1, 売上=3）
'   sKind      : "bar" / "line" / "pie"（省略不可。既定は呼び側の codegen_dsl が "bar" を渡す）
'   maxDataRow : ★ operator10 ①(片配線の解消): データ範囲の上限行（0起点・省略可）。
'                合計行のある表で codegen_dsl（Python 側・ailine_core/chart_range.py の
'                total_row.py 再利用判定）が渡す ── 自前走査の lastRow がこれを超えたら
'                切り詰め、合計行がグラフの第4の柱として混入するのを防ぐ。省略時
'                (IsMissing) は従来どおり自前走査の結果をそのまま使う（後方互換）。
Sub InsertChart(oDoc As Object, headerRow As Integer, catCol As Integer, valCol As Integer, sKind As String, Optional maxDataRow As Variant)
    Dim oSheet As Object, oCharts As Object, oChart As Object, oDiag As Object
    Dim lastRow As Long
    Dim sCat As String, sVal As String
    Dim sName As String
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（項目名の列を見出しの直下から走査）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(catCol, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If Not IsMissing(maxDataRow) Then
        If lastRow > CLng(maxDataRow) Then lastRow = CLng(maxDataRow)
    End If
    If lastRow < headerRow + 1 Then Exit Sub

    oCharts = oSheet.Charts
    sName = "Chart_" & valCol
    If oCharts.hasByName(sName) Then Exit Sub   ' 既にあれば作り直さない

    Dim oRect As New com.sun.star.awt.Rectangle
    oRect.X = 9000 : oRect.Y = 400 : oRect.Width = 14000 : oRect.Height = 8500

    ' 項目名の列 と 値の列（見出し行を含める＝ラベルになる）の2範囲（非隣接でもよい）
    Dim oRanges(1) As New com.sun.star.table.CellRangeAddress
    oRanges(0).Sheet = 0
    oRanges(0).StartColumn = catCol : oRanges(0).StartRow = headerRow
    oRanges(0).EndColumn = catCol   : oRanges(0).EndRow = lastRow
    oRanges(1).Sheet = 0
    oRanges(1).StartColumn = valCol : oRanges(1).StartRow = headerRow
    oRanges(1).EndColumn = valCol   : oRanges(1).EndRow = lastRow
    ' True,True = 先頭行=系列名・先頭列=項目名。既定(addNewByName)は縦棒グラフ。
    oCharts.addNewByName(sName, oRect, oRanges(), True, True)

    oChart = oCharts.getByName(sName).getEmbeddedObject()

    ' ── 種別の差し替え（spike 実測: 棒は既定のまま・line/pie だけ Diagram を差し替える） ──
    If sKind = "line" Then
        oChart.setDiagram(oChart.createInstance("com.sun.star.chart.LineDiagram"))
    ElseIf sKind = "pie" Then
        oChart.setDiagram(oChart.createInstance("com.sun.star.chart.PieDiagram"))
    End If

    ' ── styling（見出しから導出。LO native は色/ラベル/タイトル/軸/フォントを honor する） ──
    sCat = oSheet.getCellByPosition(catCol, headerRow).getString()
    sVal = oSheet.getCellByPosition(valCol, headerRow).getString()

    ' タイトル＝値の見出し。太字・濃色（種別によらず商品の顔を揃える）
    oChart.HasMainTitle = True
    oChart.Title.String = sVal
    oChart.Title.CharColor = &H1B2B49&      ' 濃紺
    oChart.Title.CharHeight = 15
    oChart.Title.CharWeight = com.sun.star.awt.FontWeight.BOLD
    ' 棒/折れ線は単系列なので凡例は畳む。円グラフはスライスの区別に凡例が要るので立てる。
    oChart.HasLegend = (sKind = "pie")

    If sKind <> "pie" Then
        oDiag = oChart.getDiagram()
        oDiag.HasXAxisTitle = True : oDiag.XAxisTitle.String = sCat   ' 横軸＝項目名の見出し
        ' 系列色（★16進 RRGGBB。VBASupport の RGB は BGR になるので使わない）
        oDiag.getDataRowProperties(0).FillColor = &H2E86C1&            ' 落ち着いた青
    End If
End Sub


' セル範囲を1つに結合する。★ 単一セルでなく必ず「範囲」で呼ぶこと。
'   col1,row1 = 左上（0起点）  col2,row2 = 右下
Sub MergeCells(oDoc As Object, col1 As Integer, row1 As Integer, col2 As Integer, row2 As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.getCellRangeByPosition(col1, row1, col2, row2).merge(True)
End Sub


' 行を挿入する。atRow の位置に count 行入り、既存の行は下へずれる。
'   atRow : 挿入位置（0 起点。例: 先頭データ行=見出しの次=1 の前に入れるなら atRow=1）
'   count : 挿入する行数
Sub InsertRows(oDoc As Object, atRow As Integer, count As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.Rows.insertByIndex(atRow, count)
End Sub


' 名前で行を探す（2026-08-27 追加）。0 起点の行 index / 見つからない -1 / 複数 -2。
' ★ 走査は gotoEndOfUsedArea（物理の使用範囲）── Python 側の走査は 1 列目の最初の空で
'   止まるので、表の途中に空行があるとその下を見失う。Basic 側ではその穴が構造的に無い。
Function FindRowByName(oSheet As Object, sName As String, nCol As Integer, nHeaderRow As Integer) As Integer
    Dim oCur As Object, lastRow As Long, r As Long, hit As Integer, n As Integer
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    hit = -1 : n = 0
    For r = nHeaderRow + 1 To lastRow
        If Trim(oSheet.getCellByPosition(nCol, r).getString()) = sName Then
            hit = r : n = n + 1
        End If
    Next r
    If n = 0 Then
        FindRowByName = -1
    ElseIf n > 1 Then
        FindRowByName = -2
    Else
        FindRowByName = hit
    End If
End Function


' 名前で指した行の、指定した列に**1 セルだけ**書く（2026-08-27 追加）。
' ★ なぜ 1 セル専用のヘルパか: 列全体を書く SetColumnValue を流用して走査範囲を
'   間違えると「1 セルのはずが列を潰す」── この機能で最も起きやすい壊れ方で、
'   しかも列全体の事後条件では**潰した方が pass する**（逆向きの検算になる）。
' sKind: "n"=数値 / それ以外=文字。
Sub SetCellByName(oDoc As Object, sName As String, nKeyCol As Integer, _
                   nCol As Integer, sValue As String, sKind As String, nHeaderRow As Integer)
    Dim oSheet As Object, r As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    r = FindRowByName(oSheet, sName, nKeyCol, nHeaderRow)
    If r < 0 Then Exit Sub
    If sKind = "n" Then
        oSheet.getCellByPosition(nCol, r).setValue(CDbl(sValue))
    Else
        oSheet.getCellByPosition(nCol, r).setString(sValue)
    End If
End Sub


' 行番号と列番号で指した**1 セル**に書く（2026-08-28 追加）。どちらも 0 起点。
' ★ 名前で指す SetCellByName と別に在る理由: 人が「7行目のF列」と**座標で**言った時は、
'   探し直す相手が無い（番号そのものが依頼）。同名の行が 2 つある表でも狙いが定まる。
' ★ 名前で指された時は今までどおり SetCellByName を使う ── あちらは Basic が実文書を
'   走査して自分で位置を見つけるので、Python の事後条件が独立した検算になる。
Sub SetCellAt(oDoc As Object, nRow As Integer, nCol As Integer, _
               sValue As String, sKind As String)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    If sKind = "n" Then
        oSheet.getCellByPosition(nCol, nRow).setValue(CDbl(sValue))
    Else
        oSheet.getCellByPosition(nCol, nRow).setString(sValue)
    End If
End Sub


' 行を1本挿し、指定した列に値を書く（2026-08-26 追加）。
' ★ 既存の InsertRows は**空行を挿すだけ**で、値を入れる手段が 1 つも無かった
'   （21 op のどれにも「データを 1 行足す」が無いことを実測で確認）。
' colIdxCsv: 0起点の列番号を "0,1,2" の形で／valuesCsv: 同じ並びの値／
' typesCsv:  各値の型 "s"(文字) か "n"(数値)。区切りは Chr(1)（値にカンマが入りうるため）。
Sub AddRowWithValues(oDoc As Object, atRow As Integer, colIdxCsv As String, _
                      valuesCsv As String, typesCsv As String)
    Dim oSheet As Object, oCell As Object
    Dim cols() As String, vals() As String, kinds() As String
    Dim i As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.Rows.insertByIndex(atRow, 1)
    If Len(colIdxCsv) = 0 Then Exit Sub
    cols = Split(colIdxCsv, ",")
    vals = Split(valuesCsv, Chr(1))
    kinds = Split(typesCsv, ",")
    For i = 0 To UBound(cols)
        oCell = oSheet.getCellByPosition(CInt(cols(i)), atRow)
        If i <= UBound(kinds) And kinds(i) = "n" Then
            oCell.setValue(CDbl(vals(i)))
        Else
            oCell.setString(vals(i))
        End If
    Next i
End Sub


' 名前で指した 2 行を入れ替える（2026-08-27 追加）。
' ★★ 実測で設計が決まった（bench/swap_formula_spike_RESULTS.md）:
'   セルの値を「文字として交換」すると **式が壊れる** ── みかんの行の =B3*C3 が
'   りんごの金額を出すようになり、見た目は正しく並んでいるので人が気づけない。
'   この repo が最も嫌う「静かに壊れる」形。だから値の交換では実装しない。
' ★ moveRange を使う: LibreOffice 自身が参照を付け替えるので、式は自分の行を指し続ける。
'
' ★★ 2026-08-27 に実測した落とし穴（黙って壊れる）:
'   **変数名 oR は使えない。** Basic は大文字小文字を区別しないので **予約語 Or** と
'   衝突し、**モジュールごとコンパイルに失敗**する。すると Call は何も起こさず、
'   basrun は「適用した」と言う ── **エラーが 1 行も出ないまま何も起きない**。
'   （事後条件が「変化なし」で落としたので嘘の ✓ にはならなかった。命綱は効いた）
'   ★ 切り分けの記録: 初めは createUnoStruct / Dim ... As New を疑ったが、**それは無実**
'     だった（このファイルは PivotSum 等 8 箇所で As New を使っていて動いている）。
'     1 行ずつ足して切り分けたら、犯人は名前だった。疑う順番を間違えていた。
' ★ ここではアドレスを生きたオブジェクトから貰う（.RangeAddress / .CellAddress）──
'   Sheet 番号を自分で埋めなくて済むので、対象シートの取り違えが構造的に起きない。
' ★ 空き地は使用範囲の外に取れない ── 先に insertByIndex で行を作り、最後に消す。
Sub SwapRowsByName(oDoc As Object, sA As String, sB As String, _
                    nKeyCol As Integer, nHeaderRow As Integer)
    Dim oSheet As Object, oCur As Object, oRng As Object, oDst As Object
    Dim r1 As Integer, r2 As Integer, park As Integer, lastCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    r1 = FindRowByName(oSheet, sA, nKeyCol, nHeaderRow)
    r2 = FindRowByName(oSheet, sB, nKeyCol, nHeaderRow)
    ' 見つからない(-1)/複数(-2)/同じ行 は何もしない ── Python 側の関所が先に断るが、
    ' ここでも黙って別の行を動かさない（二重の歯止め）。
    If r1 < 0 Or r2 < 0 Or r1 = r2 Then Exit Sub
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastCol = oCur.RangeAddress.EndColumn
    park = oCur.RangeAddress.EndRow + 1
    oSheet.Rows.insertByIndex(park, 1)
    ' ① A を空き地へ ② B を A の跡地へ ③ 空き地の A を B の跡地へ
    oRng = oSheet.getCellRangeByPosition(0, r1, lastCol, r1).RangeAddress
    oDst = oSheet.getCellByPosition(0, park).CellAddress
    oSheet.moveRange(oDst, oRng)
    oRng = oSheet.getCellRangeByPosition(0, r2, lastCol, r2).RangeAddress
    oDst = oSheet.getCellByPosition(0, r1).CellAddress
    oSheet.moveRange(oDst, oRng)
    oRng = oSheet.getCellRangeByPosition(0, park, lastCol, park).RangeAddress
    oDst = oSheet.getCellByPosition(0, r2).CellAddress
    oSheet.moveRange(oDst, oRng)
    oSheet.Rows.removeByIndex(park, 1)
End Sub


' 名前（見出し）で指した 2 列を入れ替える（2026-08-27 追加）。行版と同じ理屈・同じ手順。
Sub SwapColumnsByName(oDoc As Object, sA As String, sB As String, nHeaderRow As Integer)
    Dim oSheet As Object, oCur As Object, oRng As Object, oDst As Object
    Dim c1 As Integer, c2 As Integer, park As Integer, lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    c1 = FindColByNameAt(oSheet, sA, nHeaderRow)
    c2 = FindColByNameAt(oSheet, sB, nHeaderRow)
    If c1 < 0 Or c2 < 0 Or c1 = c2 Then Exit Sub
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    park = oCur.RangeAddress.EndColumn + 1
    oSheet.Columns.insertByIndex(park, 1)
    oRng = oSheet.getCellRangeByPosition(c1, 0, c1, lastRow).RangeAddress
    oDst = oSheet.getCellByPosition(park, 0).CellAddress
    oSheet.moveRange(oDst, oRng)
    oRng = oSheet.getCellRangeByPosition(c2, 0, c2, lastRow).RangeAddress
    oDst = oSheet.getCellByPosition(c1, 0).CellAddress
    oSheet.moveRange(oDst, oRng)
    oRng = oSheet.getCellRangeByPosition(park, 0, park, lastRow).RangeAddress
    oDst = oSheet.getCellByPosition(c2, 0).CellAddress
    oSheet.moveRange(oDst, oRng)
    oSheet.Columns.removeByIndex(park, 1)
End Sub


' 見出し行から列名を探す（0起点の列番号。見つからない=-1・複数=-2）。
' ★ FindRowByName の列版。見出し行を引数で受ける（多段見出しでも呼び側が決められる）。
Function FindColByNameAt(oSheet As Object, sName As String, nHeaderRow As Integer) As Integer
    Dim oCur As Object, lastCol As Integer, c As Integer, hit As Integer, n As Integer
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastCol = oCur.RangeAddress.EndColumn
    hit = -1 : n = 0
    For c = 0 To lastCol
        If Trim(oSheet.getCellByPosition(c, nHeaderRow).getString()) = sName Then
            hit = c : n = n + 1
        End If
    Next c
    If n = 0 Then
        FindColByNameAt = -1
    ElseIf n > 1 Then
        FindColByNameAt = -2
    Else
        FindColByNameAt = hit
    End If
End Function


' 行をまとめて消す（2026-08-26 追加）。atRow は 0起点、count は本数。
Sub DeleteRows(oDoc As Object, atRow As Integer, count As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.Rows.removeByIndex(atRow, count)
End Sub


' 列を1本、指定の位置に挿す（2026-08-27 追加）。colIdx は 0起点＝新しい列が入る位置。
' ★ Columns.insertByIndex を使う ── 実測で、右にあった列を参照する式は
'   **LibreOffice が自動で付け替える**（=B2*C2 → =B2*D2）。値の書き写しでは絶対にやらない。
' ★ sName が空なら見出しも空のまま入れる（人が「列を追加して」としか言わない場合）。
' ★ 変数名に oR のような予約語（Or）を使わないこと ── モジュールごと黙って死ぬ。
Sub InsertColumnAt(oDoc As Object, colIdx As Integer, sName As String, nHeaderRow As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.Columns.insertByIndex(colIdx, 1)
    If Len(sName) > 0 Then
        oSheet.getCellByPosition(colIdx, nHeaderRow).setString(sName)
    End If
End Sub


' 列を1本消す（2026-08-26 追加）。colIdx は 0起点。
Sub DeleteColumn(oDoc As Object, colIdx As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.Columns.removeByIndex(colIdx, 1)
End Sub


' データ範囲（見出し行0〜最終データ行・0列〜最終列）に格子の罫線を引く。
' ★ 範囲は自動検出する。呼び側は引数なしでよい（迷わせない）。
Sub DrawTableBorders(oDoc As Object)
    Dim oSheet As Object, oRange As Object
    Dim lastRow As Long, lastCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)

    lastRow = 0
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1
    If lastRow < 0 Or lastCol < 0 Then Exit Sub

    oRange = oSheet.getCellRangeByPosition(0, 0, lastCol, lastRow)

    Dim ln As New com.sun.star.table.BorderLine2
    ln.LineStyle = 0 : ln.LineWidth = 26      ' 細い実線
    Dim bd As New com.sun.star.table.TableBorder2
    bd.TopLine = ln : bd.BottomLine = ln : bd.LeftLine = ln : bd.RightLine = ln
    bd.HorizontalLine = ln : bd.VerticalLine = ln
    bd.IsTopLineValid = True : bd.IsBottomLineValid = True
    bd.IsLeftLineValid = True : bd.IsRightLineValid = True
    bd.IsHorizontalLineValid = True : bd.IsVerticalLineValid = True
    oRange.TableBorder2 = bd
End Sub


' 使用中の各列の幅を、内容に合わせて自動調整する。
' ★ 対象列は自動検出（見出し行0の埋まっている列）。引数なし。
Sub AutoFitColumns(oDoc As Object)
    Dim oSheet As Object, oCols As Object
    Dim i As Integer, lastCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1
    If lastCol < 0 Then Exit Sub
    oCols = oSheet.Columns
    For i = 0 To lastCol
        oCols.getByIndex(i).OptimalWidth = True
    Next i
End Sub


' 使用中の表（見出し行 headerRow〜最終データ行・0列〜最終列）の全セルを中央揃えにする。
' ★ 範囲は自動検出（最終行のみ）。セル配置は HoriJustify で設定する
'   （CharHorizontalAlignment は段落用で Calc のセルには効かない ── 7B が滑りやすい罠）。
'   headerRow : 見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   lastCol   : 表の最終列（0 起点。★ W3: 多段見出しでは見出し行の先頭列が空欄になり得るため
'               Basic で走査せず、接地済みの列数から Python 側(codegen_dsl)が渡す）
Sub AlignCenter(oDoc As Object, headerRow As Integer, lastCol As Integer)
    Dim oSheet As Object, oRange As Object
    Dim lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    ' 最終データ行（A 列を見出しの直下から走査。0行でも見出し行自体は対象にする）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastCol < 0 Then Exit Sub
    oRange = oSheet.getCellRangeByPosition(0, headerRow, lastCol, lastRow)
    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
End Sub


' 指定列のデータセル（見出し行 headerRow を除く）に3桁区切りのカンマ書式 #,##0 を付ける。
' ★ queryKey の -1（未登録）を addNew で拾う所と Locale の構築を内部で正しく処理する
'   （7B は queryKey だけ書いて addNew を落とし、Locale() を関数呼びして滑る）。
'   headerRow : 見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   col       : カンマを付ける列（0起点。例: 単価=3, 金額=4）
Sub FormatThousands(oDoc As Object, headerRow As Integer, col As Integer)
    Dim oSheet As Object, oFormats As Object
    Dim lastRow As Long, nFmt As Long
    Dim aLocale As New com.sun.star.lang.Locale
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < headerRow + 1 Then Exit Sub
    oFormats = oDoc.getNumberFormats()
    nFmt = oFormats.queryKey("#,##0", aLocale, False)
    If nFmt = -1 Then nFmt = oFormats.addNew("#,##0", aLocale)
    oSheet.getCellRangeByPosition(col, headerRow + 1, col, lastRow).NumberFormat = nFmt
End Sub


' VLOOKUP 相当。1枚目シートの各データ行について、keyCol の値をキーに
' 別表シートを照合し、見つけた値を resultCol に書く（静的な値として）。
' ★ 数式の =VLOOKUP は この経路で #VALUE! になるため、Basic 側で照合する。
' ★ 参照表(lookupSheet)は「列0=キー・列1=値」の2列表を前提にする（物理1行目が見出し・
'   検出対象外）。
'   headerRow   : 1枚目シートの見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   keyCol      : 1枚目シートの、キーが入っている列（例: 商品名=0）
'   resultCol   : 1枚目シートの、引いた値を書き込む列（例: 単価=2）
'   lookupSheet : 参照表のシート名（例: "単価表"）
Sub VLookupFromTable(oDoc As Object, headerRow As Integer, keyCol As Integer, resultCol As Integer, lookupSheet As String)
    Dim oSheet As Object, oLook As Object
    Dim lastRow As Long, lastLook As Long, i As Long, j As Long
    Dim key As String
    Dim oSrc As Object, oDst As Object

    oSheet = oDoc.Sheets.getByIndex(0)
    If Not oDoc.Sheets.hasByName(lookupSheet) Then Exit Sub
    oLook = oDoc.Sheets.getByName(lookupSheet)

    ' 対象シートの最終データ行
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    ' 参照表の最終行（参照表は常に物理1行目が見出し）
    lastLook = 1
    Do While oLook.getCellByPosition(0, lastLook).getString() <> ""
        lastLook = lastLook + 1
    Loop
    lastLook = lastLook - 1

    For i = headerRow + 1 To lastRow
        key = oSheet.getCellByPosition(keyCol, i).getString()
        For j = 1 To lastLook
            If oLook.getCellByPosition(0, j).getString() = key Then
                oSrc = oLook.getCellByPosition(1, j)      ' 参照表 列1=値
                oDst = oSheet.getCellByPosition(resultCol, i)
                If oSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                    oDst.setString(oSrc.getString())
                Else
                    oDst.setValue(oSrc.getValue())
                End If
                Exit For
            End If
        Next j
    Next i
End Sub


' ピボット集計。1枚目シートのデータを groupCol で分類し、valueCol の合計を
' 新しい「ピボット」シートに本物のピボットテーブル(DataPilot)として出す。
' ★ 出力シートと範囲は内部で組み立てる。呼び側は「分類する列」と「合計する列」だけ。
'   groupCol : 分類の基準列（0起点。例: 部門=0）
'   valueCol : 合計する値の列（例: 金額=1）
Sub PivotSum(oDoc As Object, groupCol As Integer, valueCol As Integer)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, lastCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)

    lastRow = 0
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1
    If lastRow < 1 Then Exit Sub

    Dim oSrc As New com.sun.star.table.CellRangeAddress
    oSrc.Sheet = 0 : oSrc.StartColumn = 0 : oSrc.StartRow = 0
    oSrc.EndColumn = lastCol : oSrc.EndRow = lastRow

    If Not oDoc.Sheets.hasByName("ピボット") Then
        oDoc.Sheets.insertNewByName("ピボット", oDoc.Sheets.Count)
    End If
    oOut = oDoc.Sheets.getByName("ピボット")

    Dim oDest As New com.sun.star.table.CellAddress
    oDest.Sheet = oDoc.Sheets.Count - 1 : oDest.Column = 0 : oDest.Row = 0

    Dim oTables As Object, oDesc As Object, oFields As Object, oData As Object
    oTables = oOut.DataPilotTables
    oDesc = oTables.createDataPilotDescriptor()
    oDesc.SourceRange = oSrc
    oFields = oDesc.DataPilotFields
    oFields.getByIndex(groupCol).Orientation = com.sun.star.sheet.DataPilotFieldOrientation.ROW
    oData = oFields.getByIndex(valueCol)
    oData.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.DATA
    oData.Function = com.sun.star.sheet.GeneralFunction.SUM

    If Not oTables.hasByName("Pivot1") Then
        oTables.insertNewByName("Pivot1", oDest, oDesc)
    End If
End Sub


' 集計表。1枚目シートのデータを groupCol で分類し valueCol の合計を、新しい「集計」シートに
' 見栄えのする普通の表として出す（分類×合計＋総合計行）。★ PivotSum が作る本物の DataPilot は
' LibreOffice が開くたび再描画してセル書式を撥ねる（罫線・カンマが出ない）。こちらは普通のセルに
' 書くので、格子罫線・カンマ・中央揃え・太字が native でそのまま残る（描画で確認済み）。
'   headerRow : 集計元(1枚目シート)の見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   groupCol  : 分類の基準列（0起点。例: 部門=1）
'   valueCol  : 合計する値の列（例: 金額=4）
Sub SummaryTable(oDoc As Object, headerRow As Integer, groupCol As Integer, valueCol As Integer)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, i As Long, j As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> "" : lastRow = lastRow + 1 : Loop
    lastRow = lastRow - 1
    If lastRow < headerRow + 1 Then Exit Sub

    Dim gHead As String, vHead As String
    gHead = oSheet.getCellByPosition(groupCol, headerRow).getString()
    vHead = oSheet.getCellByPosition(valueCol, headerRow).getString()

    ' 分類ごとの合計（出現順を保つ）
    Dim keys(1000) As String, sums(1000) As Double
    Dim nKeys As Integer : nKeys = 0
    Dim total As Double : total = 0
    Dim k As String, v As Double, found As Integer
    For i = headerRow + 1 To lastRow
        k = oSheet.getCellByPosition(groupCol, i).getString()
        v = oSheet.getCellByPosition(valueCol, i).getValue()
        found = -1
        For j = 0 To nKeys - 1
            If keys(j) = k Then found = j : Exit For
        Next j
        If found = -1 Then
            keys(nKeys) = k : sums(nKeys) = v : nKeys = nKeys + 1
        Else
            sums(found) = sums(found) + v
        End If
        total = total + v
    Next i

    If oDoc.Sheets.hasByName("集計") Then oDoc.Sheets.removeByName("集計")
    oDoc.Sheets.insertNewByName("集計", oDoc.Sheets.Count)
    oOut = oDoc.Sheets.getByName("集計")

    oOut.getCellByPosition(0, 0).setString(gHead)
    oOut.getCellByPosition(1, 0).setString("合計 - " & vHead)
    For j = 0 To nKeys - 1
        oOut.getCellByPosition(0, j + 1).setString(keys(j))
        oOut.getCellByPosition(1, j + 1).setValue(sums(j))
    Next j
    Dim totalRow As Integer : totalRow = nKeys + 1
    oOut.getCellByPosition(0, totalRow).setString("合計")
    oOut.getCellByPosition(1, totalRow).setValue(total)

    ' ── native 整形（普通のセルなので全部残る） ──
    Dim oRange As Object
    oRange = oOut.getCellRangeByPosition(0, 0, 1, totalRow)
    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    Dim ln As New com.sun.star.table.BorderLine2
    ln.LineStyle = 0 : ln.LineWidth = 26
    Dim bd As New com.sun.star.table.TableBorder2
    bd.TopLine = ln : bd.BottomLine = ln : bd.LeftLine = ln : bd.RightLine = ln
    bd.HorizontalLine = ln : bd.VerticalLine = ln
    bd.IsTopLineValid = True : bd.IsBottomLineValid = True
    bd.IsLeftLineValid = True : bd.IsRightLineValid = True
    bd.IsHorizontalLineValid = True : bd.IsVerticalLineValid = True
    oRange.TableBorder2 = bd
    ' 値列のデータ行＋合計にカンマ
    Dim oFormats As Object, nFmt As Long, aLocale As New com.sun.star.lang.Locale
    oFormats = oDoc.getNumberFormats()
    nFmt = oFormats.queryKey("#,##0", aLocale, False)
    If nFmt = -1 Then nFmt = oFormats.addNew("#,##0", aLocale)
    oOut.getCellRangeByPosition(1, 1, 1, totalRow).NumberFormat = nFmt
    ' 見出し行と合計行を native 太字
    Call BoldRange(oOut, 0, 0, 1, 0)
    Call BoldRange(oOut, 0, totalRow, 1, totalRow)
    ' 列幅
    oOut.Columns.getByIndex(0).OptimalWidth = True
    oOut.Columns.getByIndex(1).OptimalWidth = True
End Sub


' 範囲を太字にする。★ セルに CharWeight / CharWeightAsian / CharWeightComplex を直接当てる。
'   ★ 日本語は CharWeightAsian が効く（CharWeight だけだと日本語が太らない）。数値セルも
'     壊さず太字にできる（text cursor 経由は数値を文字列化するので使わない）。実測で
'     xlsx に太字が書き出せることを openpyxl 読み戻し＋描画の両方で確認済み。
'   col1,row1 = 左上（0起点）  col2,row2 = 右下
Sub StyleBold(oDoc As Object, col1 As Integer, row1 As Integer, col2 As Integer, row2 As Integer)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    Call BoldRange(oSheet, col1, row1, col2, row2)
End Sub


' 指定シートのセル範囲を太字にする内部ヘルパ（StyleBold / SummaryTable が使う）。
Sub BoldRange(oSheet As Object, col1 As Integer, row1 As Integer, col2 As Integer, row2 As Integer)
    Dim oCell As Object, r As Integer, c As Integer
    For r = row1 To row2
        For c = col1 To col2
            oCell = oSheet.getCellByPosition(c, r)
            ' ★ 混在文字（日本語+英字）のセルは、LO が保存時にスクリプト別のリッチテキスト run
            '   に分割し（「注文」=和文フォント/「ID」=欧文フォント）、run の書式がセルレベルの
            '   太字より勝つため、xlsx 上で太字が立たない（実測 2026-08-19: 見出し 7 セル中
            '   混在の『注文ID』だけ落ちた）。文字列を張り直して run を潰してから太字を当てる。
            '   ★ TEXT 型に限定: 数式セルに setString(getString()) すると式が表示値に化ける。
            If oCell.getType() = com.sun.star.table.CellContentType.TEXT Then
                oCell.setString(oCell.getString())
            End If
            oCell.CharWeight = com.sun.star.awt.FontWeight.BOLD
            oCell.CharWeightAsian = com.sun.star.awt.FontWeight.BOLD
            oCell.CharWeightComplex = com.sun.star.awt.FontWeight.BOLD
        Next c
    Next r
End Sub


' EXTRACT: 単一条件（col × cmp × value）に一致する行を新しいシート(dstName)へ抜き出す。
' ★ 型を保つコピー: 数値セルは getValue/setValue・文字列セルは getString/setString で
'   分岐する（getType() で判定・VLookupFromTable と同じ作法）。自由生成の実弾で観測した
'   事故（全セルを setString で書き、'59,400' のようにカンマごと文字列として焼き込む）を
'   直接潰す設計 ―― ここでは逆を書く（値を型どおりに保つ）。
'   headerRow : 元シートの見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   colIdx    : 条件を判定する列（0 起点）
'   cmpCode   : 比較の種類（0=以上 1=以下 2=超 3=未満 4=等しい 5=を含む）
'   cmpValue  : 比較する値（gte/lte/gt/lt は数値・eq は数値か文字列・contains は文字列）
'   dstName   : 出力先シート名（Python 側 codegen が col/cmp/value から決め打ちで組む。
'               既に同名シートがあれば作り直す ―― 単位F/H が「前回の自分の出力の作り
'               直しか、人の物の破壊か」を before/after の見出し署名で見分ける）
' 1 セルが条件に一致するか（0=以上 1=以下 2=超 3=未満 4=等しい 5=含む）。
' ★ 2026-08-27: ExtractRows の中にあった Select Case をここへ畳んだ。条件つきの書き込み
'   （SetColumnValueWhere）が同じ判定を必要とし、**書き写すと必ず片方だけ直る**
'   （この repo は今日それを 3 回踏んだ）。呼ぶ側は 2 つ、判定は 1 つ。
' ★ Python 側 _extract_predicate が**別実装で同じ表**を持つ（tests/test_predicate_truth_table.py
'   が凍結）。意味論を変える時は必ず両方＋その表を一緒に直すこと。
Function RowMatches(oCell As Object, cmpCode As Integer, cmpValue As Variant) As Boolean
    Dim isNumericCell As Boolean, m As Boolean
    isNumericCell = (oCell.getType() <> com.sun.star.table.CellContentType.TEXT) And (oCell.getType() <> com.sun.star.table.CellContentType.EMPTY)
    Select Case cmpCode
        Case 0   ' 以上
            m = (oCell.getValue() >= CDbl(cmpValue))
        Case 1   ' 以下
            m = (oCell.getValue() <= CDbl(cmpValue))
        Case 2   ' 超
            m = (oCell.getValue() > CDbl(cmpValue))
        Case 3   ' 未満
            m = (oCell.getValue() < CDbl(cmpValue))
        Case 4   ' 等しい（数値セルは許容誤差 1e-6 の数値比較・それ以外は文字列比較）
            If isNumericCell Then
                m = (Abs(oCell.getValue() - CDbl(cmpValue)) <= 0.000001)
            Else
                m = (oCell.getString() = CStr(cmpValue))
            End If
        Case 5   ' を含む（文字列セルのみ ── 数値/空欄は対象外）
            m = (oCell.getType() = com.sun.star.table.CellContentType.TEXT) _
                And (InStr(oCell.getString(), CStr(cmpValue)) > 0)
        Case 6   ' どれか（Chr(2) 区切りの一覧に**丸ごと一致**するものが在るか）
            ' ★ 2026-08-27: 「みかんの行とりんごの行だけ抽出して」用。
            '   ★ InStr の部分一致にしない ── 「りんご」が「青りんご」に当たると
            '     頼んでいない行が黙って混じる（contains で実測した事故と同じ形）。
            '     両端に区切りを足して、**区切りごと**探すことで丸ごと一致にする。
            m = (InStr(Chr(2) & CStr(cmpValue) & Chr(2), _
                        Chr(2) & oCell.getString() & Chr(2)) > 0)
        Case Else
            m = False
    End Select
    RowMatches = m
End Function


' 1 本の列を、指定した位置へ動かす（2026-08-27 追加）。fromIdx/toIdx は 0 起点。
' ★★ なぜ在るか: 位置の言い回し（「原価の右に」「AとBの間に」）は**どの操作でも**出る
'   のに、位置を扱えるのが列追加だけだった。計算列・転記・分割が作る新しい列は右端固定。
'   ★ op ごとに codegen を書き換えると、op が増えるたびに配線が要る（今日 4 回踏んだ形）。
'   代わりに「作ったあとで動かす」1 本の手を全部の op が共有する。
' ★ 実測（bench/swap_formula_spike_RESULTS.md）: insertByIndex も moveRange も
'   **参照を自動で付け替える**ので、式は壊れない。値の書き写しでは絶対にやらない。
' ★ 変数名に oR を使わないこと（予約語 Or と衝突してモジュールごと黙って死ぬ）。
Sub MoveColumnTo(oDoc As Object, fromIdx As Integer, toIdx As Integer)
    Dim oSheet As Object, oCur As Object, oRng As Object, oDst As Object
    Dim src As Integer, lastRow As Long
    If fromIdx = toIdx Then Exit Sub
    oSheet = oDoc.Sheets.getByIndex(0)
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    ' ① 目的地に空きを作る（ここで右側の列は 1 つずつ右へずれる ── 参照は自動追随）
    oSheet.Columns.insertByIndex(toIdx, 1)
    ' ② 挿入で元の列がずれたか（目的地より右に在ったならずれる）
    src = fromIdx
    If fromIdx >= toIdx Then src = fromIdx + 1
    ' ③ 中身を空きへ移す（moveRange も参照を付け替える）
    oRng = oSheet.getCellRangeByPosition(src, 0, src, lastRow).RangeAddress
    oDst = oSheet.getCellByPosition(toIdx, 0).CellAddress
    oSheet.moveRange(oDst, oRng)
    ' ④ 空になった元の列を詰める
    oSheet.Columns.removeByIndex(src, 1)
End Sub


' 指定した列だけを新しいシートへ抜き出す（2026-08-27 追加）。
' ★ Namakoo「特定行や特定列の抜き出しができない」── 行は ExtractRows が持っていたが、
'   列を選ぶ手段が 1 つも無かった。
' colIdxCsv: 0起点の列番号を並び順で "0,2" のように。★ 並びは Python 側が依頼文の順で決める。
' ★ 型を保つ（ExtractRows と同じ規律）: 文字は setString・数値は setValue + NumberFormat・
'   空欄は触らない。0 で埋めない。
Sub ExtractColumns(oDoc As Object, headerRow As Integer, colIdxCsv As String, dstName As String)
    Dim oSheet As Object, oOut As Object, oCur As Object
    Dim cols() As String, i As Long, j As Integer
    Dim lastRow As Long, oSrc As Object, oDst As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    If Len(colIdxCsv) = 0 Then Exit Sub
    cols = Split(colIdxCsv, ",")
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    If lastRow < headerRow Then Exit Sub
    If oDoc.Sheets.hasByName(dstName) Then oDoc.Sheets.removeByName(dstName)
    oDoc.Sheets.insertNewByName(dstName, oDoc.Sheets.Count)
    oOut = oDoc.Sheets.getByName(dstName)
    For i = headerRow To lastRow
        For j = 0 To UBound(cols)
            oSrc = oSheet.getCellByPosition(CInt(cols(j)), i)
            oDst = oOut.getCellByPosition(j, i - headerRow)
            If oSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                oDst.setString(oSrc.getString())
            ElseIf oSrc.getType() = com.sun.star.table.CellContentType.EMPTY Then
                ' 何もしない（空欄のまま）
            Else
                oDst.setValue(oSrc.getValue())
                oDst.NumberFormat = oSrc.NumberFormat
            End If
        Next j
    Next i
End Sub


' 条件に一致する行だけ、指定した列に同じ値を書く（2026-08-27 追加）。
' ★ Namakoo「原価が500以上の項目のチェック列に◎を付けて」── 表計算のごく普通の操作で、
'   ここまで一覧に無かった（SET_COLUMN_VALUE は列を丸ごと同じ値にするだけ）。
' writeCol/condCol は 0 起点。sValue は書く文字（数値化はしない ── 印を付ける用途）。
' skipRowsCsv: 対象から外す行（0 起点・カンマ区切り）。合計行など「データ行でない行」を
'   Python 側が構造として見つけて渡す。★ **条件の判定は渡さない** ── そこは Basic が
'   自分で決めるからこそ、Python の事後条件が独立した検算になる（片方に寄せない）。
Sub SetColumnValueWhere(oDoc As Object, headerRow As Integer, writeCol As Integer, _
                         condCol As Integer, cmpCode As Integer, cmpValue As Variant, _
                         sValue As String, Optional skipRowsCsv As Variant)
    Dim oSheet As Object, oCur As Object, lastRow As Long, i As Long
    Dim skips As String
    skips = ""
    If Not IsMissing(skipRowsCsv) Then skips = "," & CStr(skipRowsCsv) & ","
    oSheet = oDoc.Sheets.getByIndex(0)
    ' ★ 走査範囲は**物理の使用範囲**から取る（1 列目の空で止まる罠を避ける ──
    '   Python 側で今週 3 度直した形。Basic 側は gotoEndOfUsedArea で素直に取れる）。
    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    If lastRow < headerRow + 1 Then Exit Sub
    For i = headerRow + 1 To lastRow
        If Len(skips) > 2 And InStr(skips, "," & CStr(i) & ",") > 0 Then
            ' 対象外の行（合計行など）── 触らない
        ElseIf RowMatches(oSheet.getCellByPosition(condCol, i), cmpCode, cmpValue) Then
            oSheet.getCellByPosition(writeCol, i).setString(sValue)
        End If
    Next i
End Sub


Sub ExtractRows(oDoc As Object, headerRow As Integer, colIdx As Integer, cmpCode As Integer, cmpValue As Variant, dstName As String)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, lastCol As Integer, i As Long, j As Integer
    Dim oCell As Object, oSrc As Object, oDst As Object
    Dim outRow As Long, matched As Boolean
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（A 列を見出しの直下から走査）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    ' 最終列（見出し行を走査）
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, headerRow).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1
    If lastRow < headerRow + 1 Or lastCol < 0 Then Exit Sub

    If oDoc.Sheets.hasByName(dstName) Then oDoc.Sheets.removeByName(dstName)
    oDoc.Sheets.insertNewByName(dstName, oDoc.Sheets.Count)
    oOut = oDoc.Sheets.getByName(dstName)

    ' 見出し行のコピー（単位H の署名: 出力の1行目 = 元シートの見出し行そのもの）
    For j = 0 To lastCol
        oOut.getCellByPosition(j, 0).setString(oSheet.getCellByPosition(j, headerRow).getString())
    Next j

    outRow = 1
    For i = headerRow + 1 To lastRow
        oCell = oSheet.getCellByPosition(colIdx, i)
        matched = RowMatches(oCell, cmpCode, cmpValue)

        If matched Then
            For j = 0 To lastCol
                oSrc = oSheet.getCellByPosition(j, i)
                oDst = oOut.getCellByPosition(j, outRow)
                If oSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                    oDst.setString(oSrc.getString())
                ElseIf oSrc.getType() = com.sun.star.table.CellContentType.EMPTY Then
                    ' 何もしない（空欄のまま。0 で埋めない＝型を保つコピーの一部）
                Else
                    oDst.setValue(oSrc.getValue())
                    ' ★ 数値セルは NumberFormat も一緒に運ぶ。実測 (2026-08-20): 日付セルは getValue() が
                    '   シリアル値 (46237) を返すため、書式を運ばないと日付がただの整数に化け、
                    '   check_extract の型保存検査が正しく fail した。書式キーは同一ドキュメント内で有効。
                    oDst.NumberFormat = oSrc.NumberFormat
                End If
            Next j
            outRow = outRow + 1
        End If
    Next i
End Sub


' ★ freeform 廃止バンドル前段: DEDUP op（重複行の除去・非破壊形）。ExtractRows と同族
'   （型保存コピー・NumberFormat 運搬・見出し行コピーは単位H の署名として同じ形）。
'   違いは「条件で選ぶ」でなく「キー列の組が既出かどうかで選ぶ」こと。
' keyIdxCsv: 判定キー列（0起点の列インデックス）をカンマ区切りにした文字列（例 "0" や "0,2"）。
'   Call の引数に配列リテラルを直接書けないため、Python 側(codegen_dsl)がここで文字列に
'   組み、このヘルパが Split で戻す。
' キーの正規化: 前後空白除去のみ・型が違えば別キー（テキストセルと数値セルは常に別キー。
'   ailine_core/match.py の normalize_key・check_dedup の _dedup_normalize_key_part と
'   同じ規則 ── 3箇所が同じ規則を独立に書く点は sum_identity/total_row と同じ作法）。
' ★ 実弾の教訓（2026-08-21）: 既出キー判定を最初 Collection.Add の重複キーエラー
'   （On Error Resume Next / Err 判定）で書いたところ、実機 LO で無限ループして
'   basrun ごと固まった（60秒で戻らず PID kill で回収）。LO では**エラー頼みの制御フロー
'   が固まる**── ここではエラーを一切発生させない文字列所属検査（InStr）に置き換えている。
'   On Error / Err / Collection は今後もこの Sub に持ち込まない。
Sub DedupRows(oDoc As Object, headerRow As Integer, keyIdxCsv As String, dstName As String)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, lastCol As Integer, i As Long, j As Integer, k As Integer
    Dim oCell As Object, oSrc As Object, oDst As Object
    Dim outRow As Long
    Dim keyIdxStrs() As String
    Dim nKeys As Integer
    Dim compositeKey As String, part As String
    Dim seen As String   ' 番兵(Chr(1))区切りの既出キー連結。O(n^2) だが第一波の行数では問題にしない。
    Dim isDup As Boolean

    oSheet = oDoc.Sheets.getByIndex(0)
    keyIdxStrs = Split(keyIdxCsv, ",")
    nKeys = UBound(keyIdxStrs) - LBound(keyIdxStrs) + 1

    ' 最終データ行（A 列を見出しの直下から走査）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    ' 最終列（見出し行を走査）
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, headerRow).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1
    If lastRow < headerRow + 1 Or lastCol < 0 Then Exit Sub

    If oDoc.Sheets.hasByName(dstName) Then oDoc.Sheets.removeByName(dstName)
    oDoc.Sheets.insertNewByName(dstName, oDoc.Sheets.Count)
    oOut = oDoc.Sheets.getByName(dstName)

    ' 見出し行のコピー（単位H の署名: 出力の1行目 = 元シートの見出し行そのもの）
    For j = 0 To lastCol
        oOut.getCellByPosition(j, 0).setString(oSheet.getCellByPosition(j, headerRow).getString())
    Next j

    outRow = 1
    seen = ""
    For i = headerRow + 1 To lastRow
        ' キー列の組から複合キー文字列を作る（型ごとに接頭辞を変え、型違いを別キーにする）。
        compositeKey = ""
        For k = 0 To nKeys - 1
            oCell = oSheet.getCellByPosition(CInt(keyIdxStrs(k)), i)
            If oCell.getType() = com.sun.star.table.CellContentType.TEXT Then
                part = "S:" & Trim(oCell.getString())
            ElseIf oCell.getType() = com.sun.star.table.CellContentType.EMPTY Then
                part = "E:"
            Else
                part = "V:" & CStr(oCell.getValue())
            End If
            compositeKey = compositeKey & Chr(9) & part
        Next k

        ' 既出キーの判定はエラーに一切依存しない文字列所属検査（InStr）で行う。
        ' 番兵は Chr(1)（compositeKey 内部の区切り Chr(9) とは衝突しない）。
        If InStr(seen, Chr(1) & compositeKey & Chr(1)) > 0 Then
            isDup = True
        Else
            isDup = False
            seen = seen & Chr(1) & compositeKey & Chr(1)
        End If

        If Not isDup Then
            For j = 0 To lastCol
                oSrc = oSheet.getCellByPosition(j, i)
                oDst = oOut.getCellByPosition(j, outRow)
                If oSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                    oDst.setString(oSrc.getString())
                ElseIf oSrc.getType() = com.sun.star.table.CellContentType.EMPTY Then
                    ' 何もしない（空欄のまま。0 で埋めない＝型を保つコピーの一部）
                Else
                    oDst.setValue(oSrc.getValue())
                    ' ★ ExtractRows と同じ理由: 日付セルは getValue() がシリアル値を返すため、
                    '   書式を運ばないと日付がただの整数に化ける。
                    oDst.NumberFormat = oSrc.NumberFormat
                End If
            Next j
            outRow = outRow + 1
        End If
    Next i
End Sub


' 帳票段: REPORT_PER_ROW（DESIGN-20260823-report-per-row.md）。
' 雛形シート(templateSheet)を1枚複製して newSheetName にし、複製したシートの中の
' 印({{列名}})を、参照シート(srcSheet)の srcRow 行の対応列の値で埋める。
' ★ 憲法の適用: 機械が触ってよいのは印のあるセルだけ ── ここでは印を含まないセルには
'   一切書き込まない（複製の時点で雛形の書式・値がそのまま残る）。
' ★ 訂正2: newSheetName は Python 側(verify_dsl_args の unique_sheet_name)が
'   呼ぶ前に一意・31文字以内・禁止文字なしを確定済み ── copyByName が失敗する余地は無い
'   （失敗したら別名で再試行、はしない。孤児シートが積み上がるため）。
' ★ 訂正3: 型は元データ列で決める。丸ごと一致(セル全体が"{{列名}}")の印だけ
'   setValue/setString を型で出し分ける。部分一致（他の文字と同居する印）は文字列置換の
'   み（Python 側の verify_dsl_args が部分一致+数値列の組を事前に拒むため、ここに来る
'   部分一致は常に文字列でよい）。雛形側の数値書式（NumberFormat）には一切触れない。
' ★ 印の探索は実行時に雛形の複製そのものを走査する（Python からセル位置ごとの写像を
'   渡さない）── VLookupFromTable と同じ「シートを見出しで引く」作法の踏襲。
'   used range は createCursor().gotoEndOfUsedArea で求める（AutoFitColumns 等の
'   単純な「列0が空になるまで」走査だと、雛形の飛び飛びの空白行で取り逃がすため）。
'   headerRow/srcRow は 0 起点（Python 側 codegen_dsl が Excel の1起点から変換して渡す）。
Sub FillReportSheet(oDoc As Object, templateSheet As String, newSheetName As String, srcSheet As String, srcRow As Long, headerRow As Long)
    Dim oNew As Object, oSrc As Object
    Dim oCursor As Object, oAddr As Object
    Dim r As Long, c As Long, lastRow As Long, lastCol As Long
    Dim srcLastCol As Long, hc As Long
    Dim s As String, inner As String
    Dim p1 As Long, p2 As Long
    Dim oCellNew As Object, oCellSrc As Object

    oDoc.Sheets.copyByName(templateSheet, newSheetName, oDoc.Sheets.Count)
    oNew = oDoc.Sheets.getByName(newSheetName)
    oSrc = oDoc.Sheets.getByName(srcSheet)

    oCursor = oNew.createCursor()
    oCursor.gotoEndOfUsedArea(True)
    oAddr = oCursor.RangeAddress
    lastRow = oAddr.EndRow
    lastCol = oAddr.EndColumn

    ' 参照シートの見出し列数（見出し行を左から走査）。
    srcLastCol = 0
    Do While oSrc.getCellByPosition(srcLastCol, headerRow).getString() <> ""
        srcLastCol = srcLastCol + 1
    Loop
    srcLastCol = srcLastCol - 1

    For r = 0 To lastRow
        For c = 0 To lastCol
            oCellNew = oNew.getCellByPosition(c, r)
            s = oCellNew.getString()
            p1 = InStr(s, "{{")
            If p1 > 0 Then
                p2 = InStr(p1, s, "}}")
                If p2 > 0 Then
                    inner = Mid(s, p1 + 2, p2 - p1 - 2)
                    For hc = 0 To srcLastCol
                        If oSrc.getCellByPosition(hc, headerRow).getString() = inner Then
                            oCellSrc = oSrc.getCellByPosition(hc, srcRow)
                            If s = "{{" & inner & "}}" Then
                                ' 丸ごと一致 ── 元データ列の型で出し分ける（数値書式は触らない）。
                                If oCellSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                                    oCellNew.setString(oCellSrc.getString())
                                ElseIf oCellSrc.getType() = com.sun.star.table.CellContentType.EMPTY Then
                                    oCellNew.setString("")
                                Else
                                    oCellNew.setValue(oCellSrc.getValue())
                                End If
                            Else
                                ' 部分一致 ── 文字列置換のみ（数値列は Python 側が事前に拒む）。
                                oCellNew.setString(Left(s, p1 - 1) & oCellSrc.getString() & Mid(s, p2 + 2))
                            End If
                            Exit For
                        End If
                    Next hc
                End If
            End If
        Next c
    Next r
End Sub


' 様式写像段: FORMAT_MAP（DESIGN-20260824-format-map.md）。REPORT_PER_ROW の兄弟
' （縦の展開）── 雛形シート(templateSheet)の1行（見出し行の直下・印({{列名}})の行）を、
' 参照シート(srcSheet)の複数行それぞれで埋め、新シート(dstSheetName)へ1行ずつ積む。
' ★ 憲法の適用: 雛形には一切書き込まない（読むだけ）。機械が触るのは新シートのセルだけ。
' ★ dstSheetName は Python 側(verify_dsl_args の unique_sheet_name)が呼ぶ前に一意・
'   31文字以内・禁止文字なしを確定済み（ExtractRows/DedupRows と同じ作法）。
' ★ 型は元データ列で決める（丸ごと一致の印だけ setValue/setString を型で出し分ける・
'   部分一致は文字列置換のみ ── Python 側が部分一致+数値列の組を事前に拒む）。
' headerTplRow/phTplRow: 雛形シートの見出し行/印行（0起点）。srcHeaderRow: 参照シートの
' 見出し行（0起点）。srcRowsCsv: 出力する参照シートの行（0起点）をカンマ区切りにした文字列
' （DedupRows の keyIdxCsv と同じ作法。Python 側が総合計行の除外まで決め切って渡す）。
Sub FillFormatMapSheet(oDoc As Object, templateSheet As String, srcSheet As String, dstSheetName As String, _
                       headerTplRow As Long, phTplRow As Long, srcHeaderRow As Long, srcRowsCsv As String)
    Dim oTpl As Object, oSrc As Object, oOut As Object
    Dim oCursor As Object, oAddr As Object
    Dim tplLastCol As Long, srcLastCol As Long
    Dim c As Long, outCol As Long, outRow As Long
    Dim s As String, inner As String
    Dim p1 As Long, p2 As Long
    Dim srcRowsStrs() As String
    Dim ri As Long, srcRow As Long
    Dim hc As Long
    Dim oCellOut As Object, oCellSrc As Object

    oTpl = oDoc.Sheets.getByName(templateSheet)
    oSrc = oDoc.Sheets.getByName(srcSheet)

    If oDoc.Sheets.hasByName(dstSheetName) Then oDoc.Sheets.removeByName(dstSheetName)
    oDoc.Sheets.insertNewByName(dstSheetName, oDoc.Sheets.Count)
    oOut = oDoc.Sheets.getByName(dstSheetName)

    oCursor = oTpl.createCursor()
    oCursor.gotoEndOfUsedArea(True)
    oAddr = oCursor.RangeAddress
    tplLastCol = oAddr.EndColumn

    ' 参照シートの見出し列数（見出し行を左から走査）。
    srcLastCol = 0
    Do While oSrc.getCellByPosition(srcLastCol, srcHeaderRow).getString() <> ""
        srcLastCol = srcLastCol + 1
    Loop
    srcLastCol = srcLastCol - 1

    ' 見出し行（雛形の印行にある列だけを左から詰めて1回だけ出力する）。
    outCol = 0
    For c = 0 To tplLastCol
        s = oTpl.getCellByPosition(c, phTplRow).getString()
        p1 = InStr(s, "{{")
        If p1 > 0 Then
            If InStr(p1, s, "}}") > 0 Then
                oOut.getCellByPosition(outCol, 0).setString(oTpl.getCellByPosition(c, headerTplRow).getString())
                outCol = outCol + 1
            End If
        End If
    Next c

    srcRowsStrs = Split(srcRowsCsv, ",")
    For ri = LBound(srcRowsStrs) To UBound(srcRowsStrs)
        srcRow = CLng(srcRowsStrs(ri))
        outRow = ri - LBound(srcRowsStrs) + 1
        outCol = 0
        For c = 0 To tplLastCol
            s = oTpl.getCellByPosition(c, phTplRow).getString()
            p1 = InStr(s, "{{")
            If p1 > 0 Then
                p2 = InStr(p1, s, "}}")
                If p2 > 0 Then
                    inner = Mid(s, p1 + 2, p2 - p1 - 2)
                    oCellOut = oOut.getCellByPosition(outCol, outRow)
                    For hc = 0 To srcLastCol
                        If oSrc.getCellByPosition(hc, srcHeaderRow).getString() = inner Then
                            oCellSrc = oSrc.getCellByPosition(hc, srcRow)
                            If s = "{{" & inner & "}}" Then
                                ' 丸ごと一致 ── 元データ列の型で出し分ける（数値書式も運ぶ）。
                                If oCellSrc.getType() = com.sun.star.table.CellContentType.TEXT Then
                                    oCellOut.setString(oCellSrc.getString())
                                ElseIf oCellSrc.getType() = com.sun.star.table.CellContentType.EMPTY Then
                                    oCellOut.setString("")
                                Else
                                    oCellOut.setValue(oCellSrc.getValue())
                                    oCellOut.NumberFormat = oCellSrc.NumberFormat
                                End If
                            Else
                                ' 部分一致 ── 文字列置換のみ（数値列は Python 側が事前に拒む）。
                                oCellOut.setString(Left(s, p1 - 1) & oCellSrc.getString() & Mid(s, p2 + 2))
                            End If
                            Exit For
                        End If
                    Next hc
                    outCol = outCol + 1
                End If
            End If
        Next c
    Next ri
End Sub

' SplitColumn ── 1セルに詰まった複数値を、区切りで右の新しい列へ割る（SPLIT_CELL）。
'  ★ 新しい見出し名は Python 側が全部決めて namesCsv で渡す（Basic 側で名前を作らない）。
'  ★ 元の列は残す ── 消すと『繋ぎ直して元と一致する』検算ができなくなる。
'  ★ On Error / Collection は使わない（例外駆動の制御は LO を固める。DedupRows のコメント参照）。
Sub SplitColumn(oDoc As Object, headerRow As Integer, colIdx As Integer, sep As String, namesCsv As String)
    Dim oSheet As Object, oCell As Object
    Dim lastRow As Long, lastCol As Integer, i As Long, k As Integer
    Dim names() As String, parts() As String
    Dim baseCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    names = Split(namesCsv, ",")

    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, headerRow).getString() <> ""
        lastCol = lastCol + 1
    Loop
    If lastRow < headerRow + 1 Then Exit Sub
    baseCol = lastCol   ' データの右端の次から新しい列を作る

    For k = 0 To UBound(names)
        oSheet.getCellByPosition(baseCol + k, headerRow).setString(names(k))
    Next k

    For i = headerRow + 1 To lastRow
        parts = Split(oSheet.getCellByPosition(colIdx, i).getString(), sep)
        For k = 0 To UBound(names)
            If k <= UBound(parts) Then
                oSheet.getCellByPosition(baseCol + k, i).setString(Trim(parts(k)))
            Else
                oSheet.getCellByPosition(baseCol + k, i).setString("")
            End If
        Next k
    Next i
End Sub

' ★ 2026-08-24（土台固め）: 検分シートを **LibreOffice 側で**作る。
'   なぜ: 旧実装は openpyxl でブックを開き直して検分シートを足していた。openpyxl の
'   往復は xl/drawings の中の**図形（描かれた角印・社判・テキストボックス）を捨てる**。
'   実測: 雛形に角印のある請求書ブックで、LO が正しく N 枚へ複製した角印を、最後の
'   openpyxl 往復が全部消したうえで ✓ を出していた（帳票段の主用途そのもの）。
'   LO 経路は図形を保つと実測済みなので、書き手を LO へ寄せて往復ごと無くす。
'
'   payload の形: レコード = Chr(30) 区切り / フィールド = Chr(31) 区切り。
'   types は 1 文字ずつ列に対応（"s"=文字列 / "n"=数値）。Excel のシート名もセル値も
'   制御文字を含めないので、この 2 文字は区切りとして安全。
Sub WriteInspectionSheet(oDoc As Object, sheetName As String, payload As String, types As String)
    Dim oSheet As Object, oCell As Object
    Dim recs() As String, flds() As String
    Dim r As Long, c As Long
    Dim nCols As Long

    If oDoc.Sheets.hasByName(sheetName) Then
        oDoc.Sheets.removeByName(sheetName)
    End If
    oDoc.Sheets.insertNewByName(sheetName, oDoc.Sheets.Count)
    oSheet = oDoc.Sheets.getByName(sheetName)

    recs = Split(payload, Chr(30))
    nCols = 0
    For r = 0 To UBound(recs)
        If recs(r) <> "" Then
            flds = Split(recs(r), Chr(31))
            If UBound(flds) + 1 > nCols Then nCols = UBound(flds) + 1
            For c = 0 To UBound(flds)
                oCell = oSheet.getCellByPosition(c, r)
                If r > 0 And c < Len(types) And Mid(types, c + 1, 1) = "n" Then
                    oCell.setValue(CDbl(flds(c)))
                Else
                    oCell.setString(flds(c))
                End If
            Next c
        End If
    Next r

    ' 見出し行を太字に（既存の BoldRange と同じ作法）。
    If nCols > 0 Then
        Call BoldRange(oSheet, 0, 0, nCols - 1, 0)
        For c = 0 To nCols - 1
            oSheet.Columns.getByIndex(c).OptimalWidth = True
        Next c
    End If
End Sub


' 帳票段（まとめ版）: REPORT_PER_GROUP。同じ取引先の発注が複数行あるとき、請求書を
' **1 枚にまとめる**（REPORT_PER_ROW は 1 行 = 1 枚なので 2 枚になってしまう）。
' ★ 憲法の適用は REPORT_PER_ROW と同じ: 雛形には一切書き込まない。触るのは新シートの
'   印の在るセルだけ。シート名は Python 側(unique_sheet_name)が呼ぶ前に確定済み。
' ★ 印は 3 種類（人が雛形に書く）:
'     {{列名}}      … そのグループで同じはずの値（Python 側が食い違いを先に断っている）
'     {{明細:列名}} … 発注 1 件ごと。この行が件数ぶん増える
'     {{合計:列名}} … そのグループの合計（ここで足す。Python 側は別実装で足し直して検算する）
' ★ 予約語との衝突を避けるため変数名は接頭辞つき（実測: oR が Or と衝突してモジュールごと
'   黙って落ちた ── basrun からは「適用した」に見える）。
' srcRowsCsv: このグループの元行（0起点）をカンマ区切り。detailRow0: 雛形の明細行（0起点・
' 明細の印が無ければ -1）。
Sub FillGroupReportSheet(oDoc As Object, templateSheet As String, newSheetName As String, _
                          srcSheet As String, srcRowsCsv As String, headerRow As Long, _
                          detailRow0 As Long)
    Dim oGrp As Object, oGSrc As Object
    Dim oGCur As Object, oGAddr As Object
    Dim oGRng As Object, oGDst As Object
    Dim gRow As Long, gCol As Long, gLastRow As Long, gLastCol As Long
    Dim gSrcLastCol As Long, gHc As Long
    Dim sCell As String, sInner As String, sColName As String
    Dim pA As Long, pB As Long
    Dim oCellG As Object, oCellS As Object
    Dim rowStrs() As String
    Dim nRows As Long, k As Long, lIdx As Long, srcRow As Long
    Dim dTotal As Double
    Dim isDetail As Boolean, isTotal As Boolean

    rowStrs = Split(srcRowsCsv, ",")
    nRows = UBound(rowStrs) - LBound(rowStrs) + 1

    oDoc.Sheets.copyByName(templateSheet, newSheetName, oDoc.Sheets.Count)
    oGrp = oDoc.Sheets.getByName(newSheetName)
    oGSrc = oDoc.Sheets.getByName(srcSheet)

    ' 明細行を件数ぶんに増やす（1 件なら何もしない）。
    ' ★ 先に行を挿してから雛形の明細行をコピーする ── 書式も印の文字も一緒に運ぶ。
    If detailRow0 >= 0 And nRows > 1 Then
        oGCur = oGrp.createCursor()
        oGCur.gotoEndOfUsedArea(False)
        gLastCol = oGCur.RangeAddress.EndColumn
        oGrp.Rows.insertByIndex(detailRow0 + 1, nRows - 1)
        oGRng = oGrp.getCellRangeByPosition(0, detailRow0, gLastCol, detailRow0).RangeAddress
        For k = 1 To nRows - 1
            oGDst = oGrp.getCellByPosition(0, detailRow0 + k).CellAddress
            oGrp.copyRange(oGDst, oGRng)
        Next k
    End If

    oGCur = oGrp.createCursor()
    oGCur.gotoEndOfUsedArea(True)
    oGAddr = oGCur.RangeAddress
    gLastRow = oGAddr.EndRow
    gLastCol = oGAddr.EndColumn

    gSrcLastCol = 0
    Do While oGSrc.getCellByPosition(gSrcLastCol, headerRow).getString() <> ""
        gSrcLastCol = gSrcLastCol + 1
    Loop
    gSrcLastCol = gSrcLastCol - 1

    For gRow = 0 To gLastRow
        For gCol = 0 To gLastCol
            oCellG = oGrp.getCellByPosition(gCol, gRow)
            sCell = oCellG.getString()
            pA = InStr(sCell, "{{")
            If pA > 0 Then
                pB = InStr(pA, sCell, "}}")
                If pB > 0 Then
                    sInner = Mid(sCell, pA + 2, pB - pA - 2)
                    isDetail = (Left(sInner, 3) = "明細:")
                    isTotal = (Left(sInner, 3) = "合計:")
                    If isDetail Or isTotal Then
                        sColName = Trim(Mid(sInner, 4))
                    Else
                        sColName = Trim(sInner)
                    End If
                    ' 何件目の明細か（明細行の何行下にいるか）。明細以外は 1 件目を使う。
                    If isDetail Then
                        lIdx = gRow - detailRow0
                    Else
                        lIdx = 0
                    End If
                    If lIdx < 0 Then lIdx = 0
                    If lIdx > nRows - 1 Then lIdx = nRows - 1
                    srcRow = CLng(Trim(rowStrs(lIdx)))
                    For gHc = 0 To gSrcLastCol
                        If oGSrc.getCellByPosition(gHc, headerRow).getString() = sColName Then
                            If isTotal Then
                                dTotal = 0
                                For k = 0 To nRows - 1
                                    dTotal = dTotal + oGSrc.getCellByPosition(gHc, CLng(Trim(rowStrs(k)))).getValue()
                                Next k
                                oCellG.setValue(dTotal)
                            Else
                                oCellS = oGSrc.getCellByPosition(gHc, srcRow)
                                If sCell = "{{" & sInner & "}}" Then
                                    If oCellS.getType() = com.sun.star.table.CellContentType.TEXT Then
                                        oCellG.setString(oCellS.getString())
                                    ElseIf oCellS.getType() = com.sun.star.table.CellContentType.EMPTY Then
                                        oCellG.setString("")
                                    Else
                                        oCellG.setValue(oCellS.getValue())
                                    End If
                                Else
                                    oCellG.setString(Left(sCell, pA - 1) & oCellS.getString() & Mid(sCell, pB + 2))
                                End If
                            End If
                            Exit For
                        End If
                    Next gHc
                End If
            End If
        Next gCol
    Next gRow
End Sub
