Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call SwapRowsByName(oDoc, "みかん", "ぶどう", 0, 0)
End Sub
