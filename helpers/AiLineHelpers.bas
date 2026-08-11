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
' ★ タイトル・横軸タイトル・系列色を見出しから自動導出して styling する
'    ＝ LibreOffice native チャートの表現力を自前で引き出す（外部依存なし・ours）。
'    データラベルは付けない（値は縦軸で読める。全棒に数字を振らない方が清潔＝プロの既定）。
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


' 使用中の表（見出し行0〜最終データ行・0列〜最終列）の全セルを中央揃えにする。
' ★ 範囲は自動検出。呼び側は引数なし。セル配置は HoriJustify で設定する
'   （CharHorizontalAlignment は段落用で Calc のセルには効かない ── 7B が滑りやすい罠）。
Sub AlignCenter(oDoc As Object)
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
    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
End Sub


' 指定列のデータセル（見出し行0を除く）に3桁区切りのカンマ書式 #,##0 を付ける。
' ★ queryKey の -1（未登録）を addNew で拾う所と Locale の構築を内部で正しく処理する
'   （7B は queryKey だけ書いて addNew を落とし、Locale() を関数呼びして滑る）。
'   col : カンマを付ける列（0起点。例: 単価=3, 金額=4）
Sub FormatThousands(oDoc As Object, col As Integer)
    Dim oSheet As Object, oFormats As Object
    Dim lastRow As Long, nFmt As Long
    Dim aLocale As New com.sun.star.lang.Locale
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 1 Then Exit Sub
    oFormats = oDoc.getNumberFormats()
    nFmt = oFormats.queryKey("#,##0", aLocale, False)
    If nFmt = -1 Then nFmt = oFormats.addNew("#,##0", aLocale)
    oSheet.getCellRangeByPosition(col, 1, col, lastRow).NumberFormat = nFmt
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


' 集計表。1枚目シートのデータを groupCol で分類し valueCol の合計を、新しい「集計」シートに
' 見栄えのする普通の表として出す（分類×合計＋総合計行）。★ PivotSum が作る本物の DataPilot は
' LibreOffice が開くたび再描画してセル書式を撥ねる（罫線・カンマが出ない）。こちらは普通のセルに
' 書くので、格子罫線・カンマ・中央揃え・太字が native でそのまま残る（描画で確認済み）。
'   groupCol : 分類の基準列（0起点。例: 部門=1）
'   valueCol : 合計する値の列（例: 金額=4）
Sub SummaryTable(oDoc As Object, groupCol As Integer, valueCol As Integer)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, i As Long, j As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> "" : lastRow = lastRow + 1 : Loop
    lastRow = lastRow - 1
    If lastRow < 1 Then Exit Sub

    Dim gHead As String, vHead As String
    gHead = oSheet.getCellByPosition(groupCol, 0).getString()
    vHead = oSheet.getCellByPosition(valueCol, 0).getString()

    ' 分類ごとの合計（出現順を保つ）
    Dim keys(1000) As String, sums(1000) As Double
    Dim nKeys As Integer : nKeys = 0
    Dim total As Double : total = 0
    Dim k As String, v As Double, found As Integer
    For i = 1 To lastRow
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
            oCell.CharWeight = com.sun.star.awt.FontWeight.BOLD
            oCell.CharWeightAsian = com.sun.star.awt.FontWeight.BOLD
            oCell.CharWeightComplex = com.sun.star.awt.FontWeight.BOLD
        Next c
    Next r
End Sub
