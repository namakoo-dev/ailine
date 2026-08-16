Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, r As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 3
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 2 Then Exit Sub
    For r = 2 To lastRow
        oSheet.getCellByPosition(3, r).CellBackColor = &HFF0000&
    Next r
End Sub
