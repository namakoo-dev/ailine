Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call DeleteRows(oDoc, 1, 3)
End Sub
