Option VBASupport 1
Option Explicit

Sub Run(oDoc As Object)
    Call ExtractRows(oDoc, 0, 0, 6, "みかん" & Chr(2) & "りんご", "商品みかん・りんごのどれか", "")
End Sub
