Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, totalRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 2)
    If lastRow < 3 Then Exit Sub
    totalRow = lastRow + 1
    oSheet.getCellByPosition(0, totalRow).setString("税込合計")
    oSheet.getCellByPosition(3, totalRow).setFormula("=SUM(" & "D" & 4 & ":INDEX(" & "D" & ":" & "D" & ";ROW()-1))" & "*1.1")
    oSheet.getCellByPosition(3, totalRow).NumberFormat = oSheet.getCellByPosition(3, totalRow - 1).NumberFormat
End Sub
