Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, oRange As Object, lastRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 2)
    If lastRow < 2 Then Exit Sub
    oRange = oSheet.getCellRangeByPosition(0, 2, 0, lastRow)
    oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
End Sub
