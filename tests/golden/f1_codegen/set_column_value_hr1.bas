Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, r As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 0)
    If lastRow < 1 Then Exit Sub
    For r = 1 To lastRow
        oSheet.getCellByPosition(0, r).setString("確認済み")
    Next r
End Sub
