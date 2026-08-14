Option VBASupport 1
Option Explicit
Sub Run(oDoc As Object)
    Dim oSheet As Object, oCell As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    oCell = oSheet.getCellByPosition(3, 1)
    oCell.setFormula("=B2*C2")
    oCell = oSheet.getCellByPosition(3, 2)
    oCell.setFormula("=B3*C3")
    oCell = oSheet.getCellByPosition(3, 3)
    oCell.setFormula("=B4*C4")
    oCell = oSheet.getCellByPosition(4, 1)
    oCell.setFormula("=SUM(D2:D4)")
    oCell = oSheet.getCellByPosition(5, 1)
    oCell.setFormula("=VLOOKUP(A2,単価表!A2:B4,2,0)")
    oCell = oSheet.getCellByPosition(5, 2)
    oCell.setFormula("=VLOOKUP(A3;単価表!A2:B4;2;0)")
    oCell = oSheet.getCellByPosition(6, 1)
    oCell.setFormula("=VLOOKUP(A2,単価表.A2:B4,2,0)")
    oCell = oSheet.getCellByPosition(6, 2)
    oCell.setFormula("=VLOOKUP(A3;単価表.A2:B4;2;0)")
End Sub
