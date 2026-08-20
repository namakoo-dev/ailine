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


' 1枚目シートに「見栄えのする」棒グラフを1つ挿入する。範囲もスタイルも内部で組む。
' ★ 項目名は先頭列(列0)に固定。呼び側は「見出し行・値の列」を渡す（迷わせない）。
' ★ タイトル・横軸タイトル・系列色を見出しから自動導出して styling する
'    ＝ LibreOffice native チャートの表現力を自前で引き出す（外部依存なし・ours）。
'    データラベルは付けない（値は縦軸で読める。全棒に数字を振らない方が清潔＝プロの既定）。
'   headerRow : 見出し行（0 起点。W3: StructDump が推定した実際の見出し行）
'   valCol    : 棒にする値の列（0 起点。例: 金額=1, 売上=3）
Sub InsertBarChart(oDoc As Object, headerRow As Integer, valCol As Integer)
    Dim oSheet As Object, oCharts As Object, oChart As Object, oDiag As Object
    Dim lastRow As Long
    Dim catCol As Integer
    Dim sCat As String, sVal As String
    Dim sName As String
    catCol = 0                        ' 項目名は先頭列に固定
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行（A 列を見出しの直下から走査）
    lastRow = headerRow + 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < headerRow + 1 Then Exit Sub

    oCharts = oSheet.Charts
    sName = "Chart_" & valCol
    If oCharts.hasByName(sName) Then Exit Sub   ' 既にあれば作り直さない

    Dim oRect As New com.sun.star.awt.Rectangle
    oRect.X = 9000 : oRect.Y = 400 : oRect.Width = 14000 : oRect.Height = 8500

    ' 項目名の列 と 値の列（見出し行を含める＝ラベルになる）の2範囲
    Dim oRanges(1) As New com.sun.star.table.CellRangeAddress
    oRanges(0).Sheet = 0
    oRanges(0).StartColumn = catCol : oRanges(0).StartRow = headerRow
    oRanges(0).EndColumn = catCol   : oRanges(0).EndRow = lastRow
    oRanges(1).Sheet = 0
    oRanges(1).StartColumn = valCol : oRanges(1).StartRow = headerRow
    oRanges(1).EndColumn = valCol   : oRanges(1).EndRow = lastRow
    ' True,True = 先頭行=系列名・先頭列=項目名。既定は縦棒グラフ。
    oCharts.addNewByName(sName, oRect, oRanges(), True, True)

    ' ── styling（見出しから導出。LO native は色/ラベル/タイトル/軸/フォントを honor する） ──
    sCat = oSheet.getCellByPosition(catCol, headerRow).getString()
    sVal = oSheet.getCellByPosition(valCol, headerRow).getString()
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
Sub ExtractRows(oDoc As Object, headerRow As Integer, colIdx As Integer, cmpCode As Integer, cmpValue As Variant, dstName As String)
    Dim oSheet As Object, oOut As Object
    Dim lastRow As Long, lastCol As Integer, i As Long, j As Integer
    Dim oCell As Object, oSrc As Object, oDst As Object
    Dim outRow As Long, matched As Boolean, isNumericCell As Boolean
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
        isNumericCell = (oCell.getType() <> com.sun.star.table.CellContentType.TEXT) And (oCell.getType() <> com.sun.star.table.CellContentType.EMPTY)

        Select Case cmpCode
            Case 0   ' 以上
                matched = (oCell.getValue() >= CDbl(cmpValue))
            Case 1   ' 以下
                matched = (oCell.getValue() <= CDbl(cmpValue))
            Case 2   ' 超
                matched = (oCell.getValue() > CDbl(cmpValue))
            Case 3   ' 未満
                matched = (oCell.getValue() < CDbl(cmpValue))
            Case 4   ' 等しい（数値セルは数値比較・それ以外は文字列比較）
                If isNumericCell Then
                    matched = (oCell.getValue() = CDbl(cmpValue))
                Else
                    matched = (oCell.getString() = CStr(cmpValue))
                End If
            Case 5   ' を含む（常に文字列の部分一致）
                matched = (InStr(oCell.getString(), CStr(cmpValue)) > 0)
            Case Else
                matched = False
        End Select

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
