Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SetColumnValueWhere(oDoc, 0, 2, 1, 0, 40000.0, "◎")
End Sub
