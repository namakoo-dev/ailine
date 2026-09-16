Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, r As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 2)
    If lastRow < 2 Then Exit Sub
    For r = 2 To lastRow
        oSheet.getCellByPosition(3, r).CellBackColor = &HFF0000&
    Next r
End Sub
