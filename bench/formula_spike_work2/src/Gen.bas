Option VBASupport 1
Option Explicit
Sub Run(oDoc As Object)
    Dim oSheet As Object, oCell As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    ' 変種1: セミコロン (LO 方言)
    oCell = oSheet.getCellByPosition(3, 4)
    oCell.setFormula("=SUM(D2:INDEX(D:D;ROW()-1))*1.1")
    ' 変種2: カンマ (Excel 方言)
    oCell = oSheet.getCellByPosition(4, 4)
    oCell.setFormula("=SUM(E2:INDEX(E:E,ROW()-1))*1.1")
End Sub
