Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call ExtractColumns(oDoc, 0, "0,2", "商品・在庫だけ")
End Sub
