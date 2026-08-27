Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SetColumnValueWhere(oDoc, 0, 2, 0, 5, "セット", "×", "")
End Sub
