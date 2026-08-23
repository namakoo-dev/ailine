Option VBASupport 1
Option Explicit

' 条件に合う行のセルに背景色を付ける例（在庫が 10 未満なら薄赤）。
' ★★ 色は必ず 16 進リテラル &HRRGGBB& で指定する。RGB() は使わない。
'    Option VBASupport 1 の下では RGB(r,g,b) が VBA 互換(BGR)になり、
'    RGB(255,0,0) が赤でなく青になる（実測で確認）。&HFF0000& なら確実に赤。
' 　　例: 赤=&HFF0000&  薄赤=&HFFCCCC&  緑=&H00B050&  黄=&HFFFF00&  青=&H0000FF&
Sub Run(oDoc As Object)
    Dim oSheet As Object, i As Long, lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終データ行を探す（A 列を上から走査）
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    For i = 1 To lastRow
        If oSheet.getCellByPosition(2, i).getValue() < 10 Then
            oSheet.getCellByPosition(2, i).CellBackColor = &HFFCCCC&   ' 薄赤
        End If
    Next i
End Sub
