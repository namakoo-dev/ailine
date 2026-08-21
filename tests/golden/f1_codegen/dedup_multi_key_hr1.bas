Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call DedupRows(oDoc, 0, "0,2", "商品・単価の重複除去")
End Sub
