Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SortByColumn(oDoc, 0, 3, 3, True)
End Sub
