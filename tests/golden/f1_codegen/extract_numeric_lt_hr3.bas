Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call ExtractRows(oDoc, 2, 3, 3, 20, "金額20未満")
End Sub
