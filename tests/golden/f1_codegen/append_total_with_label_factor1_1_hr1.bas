Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, totalRow As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 1 Then Exit Sub
    totalRow = lastRow + 1
    oSheet.getCellByPosition(2, totalRow).setString("税込合計")
    oSheet.getCellByPosition(3, totalRow).setFormula("=SUM(" & "D" & 2 & ":INDEX(" & "D" & ":" & "D" & ";ROW()-1))" & "*1.1")
End Sub
