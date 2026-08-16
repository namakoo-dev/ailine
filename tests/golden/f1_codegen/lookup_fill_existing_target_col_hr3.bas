Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call VLookupFromTable(oDoc, 2, 0, 2, "単価表")
End Sub
