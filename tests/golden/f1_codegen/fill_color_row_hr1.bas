Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Dim oSheet As Object, c As Integer
    oSheet = oDoc.Sheets.getByIndex(0)
    For c = 0 To 3
        oSheet.getCellByPosition(c, 0).CellBackColor = &HFFFF00&
    Next c
End Sub
