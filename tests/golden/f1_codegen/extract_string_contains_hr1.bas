Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call ExtractRows(oDoc, 0, 0, 5, "セット", "商品セットを含む", "")
End Sub
