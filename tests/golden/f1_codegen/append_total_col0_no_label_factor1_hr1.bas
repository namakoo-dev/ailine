Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, totalRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = TableLastRow(oSheet, 0)
    If lastRow < 1 Then Exit Sub
    totalRow = lastRow + 1
    oSheet.getCellByPosition(0, totalRow).setFormula("=SUM(" & "A" & 2 & ":INDEX(" & "A" & ":" & "A" & ";ROW()-1))" & "")
    oSheet.getCellByPosition(0, totalRow).NumberFormat = oSheet.getCellByPosition(0, totalRow - 1).NumberFormat
End Sub
