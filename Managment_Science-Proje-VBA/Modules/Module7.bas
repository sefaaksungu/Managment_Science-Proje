Attribute VB_Name = "Module7"
Option Explicit

Sub KýsýtlarýnSaðTarafDeðerleri()
Attribute KýsýtlarýnSaðTarafDeðerleri.VB_ProcData.VB_Invoke_Func = " \n14"
'
' KýsýtlarýnSaðTarafDeðerleri Makro
'

'
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Sheets("Data ve Notasyon").Select
    Range("K4").Select
    Application.Goto Reference:="P."
    Range("K5").Select
    ActiveWorkbook.Names.Add Name:="D.", RefersToR1C1:= _
        "='Data ve Notasyon'!R5C11"
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    Range("X22").Select
    ActiveCell.FormulaR1C1 = "=P."
    Range("X30").Select
    ActiveCell.FormulaR1C1 = "=D."
    Range("X31").Select
    ActiveWindow.SmallScroll Down:=-22
    Range("S4").Select
    ActiveCell.FormulaR1C1 = "=a.11"
    Range("S5").Select
    ActiveCell.FormulaR1C1 = "=a.12"
    Range("S6").Select
    ActiveCell.FormulaR1C1 = "=a.13"
    Range("S7").Select
    ActiveCell.FormulaR1C1 = "=a.21"
    Range("S8").Select
    ActiveCell.FormulaR1C1 = "=a.22"
    Range("S9").Select
    ActiveCell.FormulaR1C1 = "=a.23"
    Range("S10").Select
    ActiveCell.FormulaR1C1 = "=a.31"
    Range("S11").Select
    ActiveCell.FormulaR1C1 = "=a.32"
    Range("S12").Select
    ActiveCell.FormulaR1C1 = "=a.33"
    Range("S13").Select
    ActiveWindow.SmallScroll Down:=11
    Range("S17").Select
    ActiveCell.FormulaR1C1 = "=b.1"
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "=b.2"
    Range("S19").Select
    ActiveCell.FormulaR1C1 = "=b.3"
    Range("S20").Select
    ActiveCell.FormulaR1C1 = "=b.4"
    Range("S21").Select
    ActiveCell.FormulaR1C1 = "=b.5"
    Range("S22").Select
    ActiveWindow.SmallScroll Down:=0
    Range("R26").Select
    ActiveCell.FormulaR1C1 = "=c.1"
    Range("R27").Select
    ActiveCell.FormulaR1C1 = "=c.2"
    Range("R28").Select
    ActiveCell.FormulaR1C1 = "=c.3"
    Range("R29").Select
    ActiveCell.FormulaR1C1 = "=c.4"
    Range("R30").Select
    ActiveCell.FormulaR1C1 = "=c.5"
    Range("R31").Select
    ActiveWindow.SmallScroll Down:=11
    Range("L33").Select
    ActiveCell.FormulaR1C1 = "=d.1"
    Range("M33").Select
    ActiveCell.FormulaR1C1 = "=d.2"
    Range("N33").Select
    ActiveCell.FormulaR1C1 = "=d.3"
    Range("O33").Select
    ActiveCell.FormulaR1C1 = "=d.4"
    Range("O34").Select
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
End Sub
