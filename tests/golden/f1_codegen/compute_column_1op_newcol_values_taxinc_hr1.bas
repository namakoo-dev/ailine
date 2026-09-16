Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, i As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 0)
    If lastRow < 1 Then Exit Sub
    oSheet.getCellByPosition(4, 0).setString("税込金額")
    For i = 1 To lastRow
        oSheet.getCellByPosition(4, i).setValue(oSheet.getCellByPosition(3, i).getValue() * 1.1)
    Next i
End Sub
