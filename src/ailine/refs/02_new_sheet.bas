Option VBASupport 1
Option Explicit

' 新しいシートを末尾に作り、そこに書き込む例。
' ★ insertNewByName(名前, 位置)。位置は Sheets.Count で末尾。
'    既にあれば作り直さない（hasByName で確認）。
Sub Run(oDoc As Object)
    Dim oSheets As Object, oNew As Object
    oSheets = oDoc.Sheets
    If Not oSheets.hasByName("メモ") Then
        oSheets.insertNewByName("メモ", oSheets.Count)
    End If
    oNew = oSheets.getByName("メモ")
    oNew.getCellByPosition(0, 0).setString("hello")
End Sub
