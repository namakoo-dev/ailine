Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call FillFormatMapSheet(oDoc, "単価表", "Sheet", "単価表_出力", 0, 1, 2, "3,4")
End Sub
