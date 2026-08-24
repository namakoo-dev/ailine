Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SplitColumn(oDoc, 0, 0, "" & Chr(10) & "", "商品_1,商品_2")
End Sub
