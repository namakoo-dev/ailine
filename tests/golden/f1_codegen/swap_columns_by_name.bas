Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SwapColumnsByName(oDoc, "商品", "金額", 0)
End Sub
