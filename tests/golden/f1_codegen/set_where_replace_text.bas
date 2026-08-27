Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SetColumnValueWhere(oDoc, 0, 0, 0, 4, "りんご", "林檎", "")
End Sub
