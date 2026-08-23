Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call InsertChart(oDoc, 2, 0, 3, "bar")
End Sub
