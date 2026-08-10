Option VBASupport 1
Option Explicit

' ────────────────────────────────────────────────────────────────
'  ailine の検証済みヘルパ集。arcane な UNO 操作を「呼ぶだけ」にする。
'  モデルはこれらを呼ぶだけ。中の難所（ソートの ContainsHeader 等）は触らせない。
'  ★ ここは人が検証して固定する。生成物ではない。
' ────────────────────────────────────────────────────────────────

' 1枚目シートのデータ行（見出し行0を除く）を、col 列で並べ替える。
' 範囲と見出し扱いは内部で自動処理する。呼び側は列と向きだけ渡す。
'   col        : 並べ替えの基準列（0 起点）
'   ascending  : True=昇順, False=降順
Sub SortByColumn(oDoc As Object, col As Integer, ascending As Boolean)
    Dim oSheet As Object, oRange As Object
    Dim lastRow As Long, lastCol As Integer
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（A 列を上から走査）
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    ' 最終列（見出し行0を左から走査）
    lastCol = 0
    Do While oSheet.getCellByPosition(lastCol, 0).getString() <> ""
        lastCol = lastCol + 1
    Loop
    lastCol = lastCol - 1

    If lastRow < 1 Then Exit Sub   ' データが無い

    ' 行1..lastRow（見出しを含めない）を範囲にする
    oRange = oSheet.getCellRangeByPosition(0, 1, lastCol, lastRow)

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


' 1枚目シートに「見栄えのする」棒グラフを1つ挿入する。範囲もスタイルも内部で組む。
' ★ 項目名は先頭列(列0)に固定。呼び側は「値の列」だけ渡す（迷わせない）。
' ★ タイトル・横軸タイトル・データラベル・系列色を見出しから自動導出して styling する
'    ＝ LibreOffice native チャートの表現力を自前で引き出す（外部依存なし・ours）。
'   valCol : 棒にする値の列（0 起点。例: 金額=1, 売上=3）
Sub InsertBarChart(oDoc As Object, valCol As Integer)
    Dim oSheet As Object, oCharts As Object, oChart As Object, oDiag As Object
    Dim lastRow As Long
    Dim catCol As Integer
    Dim sCat As String, sVal As String
    Dim sName As String
    catCol = 0                        ' 項目名は先頭列に固定
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（A 列を上から走査）
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 1 Then Exit Sub

    oCharts = oSheet.Charts
    sName = "Chart_" & valCol
    If oCharts.hasByName(sName) Then Exit Sub   ' 既にあれば作り直さない

    Dim oRect As New com.sun.star.awt.Rectangle
    oRect.X = 9000 : oRect.Y = 400 : oRect.Width = 14000 : oRect.Height = 8500

    ' 項目名の列 と 値の列（見出し行0を含める＝ラベルになる）の2範囲
    Dim oRanges(1) As New com.sun.star.table.CellRangeAddress
    oRanges(0).Sheet = 0
    oRanges(0).StartColumn = catCol : oRanges(0).StartRow = 0
    oRanges(0).EndColumn = catCol   : oRanges(0).EndRow = lastRow
    oRanges(1).Sheet = 0
    oRanges(1).StartColumn = valCol : oRanges(1).StartRow = 0
    oRanges(1).EndColumn = valCol   : oRanges(1).EndRow = lastRow
    ' True,True = 先頭行=系列名・先頭列=項目名。既定は縦棒グラフ。
    oCharts.addNewByName(sName, oRect, oRanges(), True, True)

    ' ── styling（見出しから導出。LO native は色/ラベル/タイトル/軸/フォントを honor する） ──
    sCat = oSheet.getCellByPosition(catCol, 0).getString()
    sVal = oSheet.getCellByPosition(valCol, 0).getString()
    oChart = oCharts.getByName(sName).getEmbeddedObject()

    ' タイトル＝値の見出し。太字・濃色
    oChart.HasMainTitle = True
    oChart.Title.String = sVal
    oChart.Title.CharColor = &H1B2B49&      ' 濃紺
    oChart.Title.CharHeight = 15
    oChart.Title.CharWeight = com.sun.star.awt.FontWeight.BOLD
    ' 単系列なので凡例は畳む（余計な要素を出さない）
    oChart.HasLegend = False

    oDiag = oChart.getDiagram()
    oDiag.DataCaption = com.sun.star.chart.ChartDataCaption.VALUE   ' 各棒に値
    oDiag.HasXAxisTitle = True : oDiag.XAxisTitle.String = sCat     ' 横軸＝項目名の見出し
    ' 系列色（★16進 RRGGBB。VBASupport の RGB は BGR になるので使わない）
    oDiag.getDataRowProperties(0).FillColor = &H2E86C1&            ' 落ち着いた青
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


' VLOOKUP 相当。1枚目シートの各データ行について、keyCol の値をキーに
' 別表シートを照合し、見つけた値を resultCol に書く（静的な値として）。
' ★ 数式の =VLOOKUP は この経路で #VALUE! になるため、Basic 側で照合する。
' ★ 参照表(lookupSheet)は「列0=キー・列1=値」の2列表を前提にする。
'   keyCol      : 1枚目シートの、キーが入っている列（例: 商品名=0）
'   resultCol   : 1枚目シートの、引いた値を書き込む列（例: 単価=2）
'   lookupSheet : 参照表のシート名（例: "単価表"）
Sub VLookupFromTable(oDoc As Object, keyCol As Integer, resultCol As Integer, lookupSheet As String)
    Dim oSheet As Object, oLook As Object
    Dim lastRow As Long, lastLook As Long, i As Long, j As Long
    Dim key As String
    Dim oSrc As Object, oDst As Object

    oSheet = oDoc.Sheets.getByIndex(0)
    If Not oDoc.Sheets.hasByName(lookupSheet) Then Exit Sub
    oLook = oDoc.Sheets.getByName(lookupSheet)

    ' 対象シートの最終データ行
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    ' 参照表の最終行
    lastLook = 1
    Do While oLook.getCellByPosition(0, lastLook).getString() <> ""
        lastLook = lastLook + 1
    Loop
    lastLook = lastLook - 1

    For i = 1 To lastRow
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


' 太字。★ このヘルパは Basic 側では何もしない（no-op）。
'   理由: LibreOffice は CharWeight の太字を xlsx に書き出せない（実測・描画で確認）。
'   そこで ailine が basrun 適用後に openpyxl で太字を後付けする（自作の道）。
'   モデルは太字にしたい範囲を渡してこれを呼ぶだけでよい。
'   col1,row1 = 左上（0起点）  col2,row2 = 右下
Sub StyleBold(oDoc As Object, col1 As Integer, row1 As Integer, col2 As Integer, row2 As Integer)
    ' 意図的に空。実体は ailine 側（Python/openpyxl）が適用する。
End Sub
