Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SetCellByName(oDoc, "りんご", 0, 1, "2000.0", "n", 0)
End Sub
