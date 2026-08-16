Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    oDoc.Sheets.getByIndex(0).getCellByPosition(4, 0).setString("備考")
    Call VLookupFromTable(oDoc, 0, 0, 4, "単価表")
End Sub
