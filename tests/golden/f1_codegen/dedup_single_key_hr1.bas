Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call DedupRows(oDoc, 0, "0", "商品の重複除去")
End Sub
