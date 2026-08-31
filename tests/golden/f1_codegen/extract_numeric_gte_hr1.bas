Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call ExtractRows(oDoc, 0, 3, 0, 40000, "金額40000以上", "")
End Sub
