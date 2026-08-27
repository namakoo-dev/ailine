Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SetCellByName(oDoc, "りんご", 0, 0, "洋梨", "s", 0)
End Sub
