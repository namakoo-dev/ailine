Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call FillReportSheet(oDoc, "単価表", "甲社", "Sheet", 1, 0)
    Call FillReportSheet(oDoc, "単価表", "乙社", "Sheet", 2, 0)
End Sub
