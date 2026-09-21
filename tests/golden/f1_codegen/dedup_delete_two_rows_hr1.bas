Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call DeleteRows(oDoc, 4, 1)
    Call DeleteRows(oDoc, 2, 1)
End Sub
