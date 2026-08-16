Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, oRange As Object, lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 0 Then Exit Sub
    oRange = oSheet.getCellRangeByPosition(0, 0, 0, lastRow)
    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
End Sub
