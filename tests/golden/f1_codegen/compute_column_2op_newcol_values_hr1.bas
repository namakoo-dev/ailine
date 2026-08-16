Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, lastRow As Long, i As Long
    oSheet = oDoc.Sheets.getByIndex(0)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1
    If lastRow < 1 Then Exit Sub
    oSheet.getCellByPosition(4, 0).setString("数量*単価")
    For i = 1 To lastRow
        oSheet.getCellByPosition(4, i).setValue(oSheet.getCellByPosition(1, i).getValue() * oSheet.getCellByPosition(2, i).getValue())
    Next i
End Sub
