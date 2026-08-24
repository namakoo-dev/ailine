Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call FillReportSheet(oDoc, "単価表", "甲社", "Sheet", 3, 2)
    Call FillReportSheet(oDoc, "単価表", "乙社", "Sheet", 4, 2)
End Sub
