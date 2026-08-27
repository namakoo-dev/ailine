Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call InsertColumnAt(oDoc, 2, "備考", 0)
End Sub
