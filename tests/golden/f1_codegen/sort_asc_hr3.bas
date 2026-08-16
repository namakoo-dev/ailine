Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SortByColumn(oDoc, 2, 3, 3, True)
End Sub
