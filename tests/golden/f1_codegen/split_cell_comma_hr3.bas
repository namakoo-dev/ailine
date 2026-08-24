Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SplitColumn(oDoc, 2, 0, ",", "商品_1,商品_2,商品_3")
End Sub
