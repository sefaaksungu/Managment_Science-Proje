Attribute VB_Name = "Module6"
Option Explicit

Sub BirinciAsamaDengeKýsýtýSolTaraf()
Attribute BirinciAsamaDengeKýsýtýSolTaraf.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BirinciAsamaDengeKýsýtýSolTaraf Makro
'

'
    Range("L39").Select
    ActiveCell.FormulaR1C1 = "=X.111+X.211+X.311-2*Y11tY12tY13tY14tY15"
    Range("L40").Select
    ActiveWindow.SmallScroll Down:=11
    Range("M39").Select
    ActiveCell.FormulaR1C1 = "=X.121+X.221+X.321-2*Y21tY22tY23tY24tY25"
    Range("N39").Select
    ActiveCell.FormulaR1C1 = "=X.131+X.231+X.331-2*Y31tY32tY33tY34tY35"
    Range("N40").Select
    ActiveWindow.SmallScroll Down:=0
    Range("O39").Select
    ActiveCell.FormulaR1C1 = "=X.141+X.241+X.341-2*Y41tY42tY43tY44tY45"
    Range("O40").Select
    ActiveWindow.SmallScroll Down:=11
    Range("P39").Select
    ActiveCell.FormulaR1C1 = "=X.151+X.251+X.351-2*Y51tY52tY53tY54tY55"
    Range("P40").Select
    ActiveWindow.SmallScroll Down:=11
    Range("L40").Select
    ActiveCell.FormulaR1C1 = "=X.112+X.212+X.312-1*Y11tY12tY13tY14tY15"
    Range("M40").Select
    ActiveCell.FormulaR1C1 = "=X.122+X.222+X.322-1*Y21tY22tY23tY24tY25"
    Range("N40").Select
    ActiveCell.FormulaR1C1 = "=X.132+X.232+X.332-1*Y31tY32tY33tY34tY35"
    Range("O40").Select
    ActiveCell.FormulaR1C1 = "=X.142+X.242+X.342-1*Y41tY42tY43tY44tY45"
    Range("P40").Select
    ActiveCell.FormulaR1C1 = "=X.152+X.252+X.352-1*Y51tY52tY53tY54tY55"
    Range("L41").Select
    ActiveCell.FormulaR1C1 = "=X.113+X.213+X.313-3*Y11tY12tY13tY14tY15"
    Range("M41").Select
    ActiveCell.FormulaR1C1 = "=X.123+X.223+X.323-3*Y21tY22tY23tY24tY25"
    Range("N41").Select
    ActiveCell.FormulaR1C1 = "=X.133+X.233+X.333-3*Y31tY32tY33tY34tY35"
    Range("O41").Select
    ActiveCell.FormulaR1C1 = "=X.143+X.243+X.343-3*Y41tY42tY43tY44tY45"
    Range("P41").Select
    ActiveCell.FormulaR1C1 = "=X.153+X.253+X.353-3*Y51tY52tY53tY54tY55"
    Range("P42").Select
End Sub
Sub ÝkinciAsamaDengeKýsýtýSolTaraf()
Attribute ÝkinciAsamaDengeKýsýtýSolTaraf.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ÝkinciAsamaDengeKýsýtýSolTaraf Makro
'

'
    ActiveCell.FormulaR1C1 = "=Y.11+Y.21+Y.31+Y.41+Y.51-Z11tZ12tZ13tZ14"
    Range("M37").Select
    ActiveCell.FormulaR1C1 = "=Y.12+Y.22+Y.32+Y.42+Y.52-Z21tZ22tZ23tZ24"
    Range("N37").Select
    ActiveCell.FormulaR1C1 = "=Y.13+Y.23+Y.33+Y.43+Y.53-Z31tZ32tZ33tZ34"
    Range("O37").Select
    ActiveCell.FormulaR1C1 = "=Y.14+Y.24+Y.34+Y.44+Y.54-Z41tZ42tZ43tZ44"
    Range("P37").Select
    ActiveCell.FormulaR1C1 = "=Y.15+Y.25+Y.35+Y.45+Y.55-Z51tZ52tZ53tZ54"
    Range("P38").Select
End Sub
Sub ParametrelerinAdlandýrýlmasý()
Attribute ParametrelerinAdlandýrýlmasý.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ParametrelerinAdlandýrýlmasý Makro
'

'
    Range("L11:P19").Select
    ActiveWorkbook.Names.Add Name:="Cijt", RefersToR1C1:= _
        "='Data ve Notasyon'!R11C12:R19C16"
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
    Range("L24:P28").Select
    ActiveWorkbook.Names.Add Name:="Cjk", RefersToR1C1:= _
        "='Data ve Notasyon'!R24C12:R28C16"
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    Range("L33:O37").Select
    ActiveWorkbook.Names.Add Name:="Ckl", RefersToR1C1:= _
        "='Data ve Notasyon'!R33C12:R37C15"
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
    ActiveWindow.ScrollRow = 4
    Range("T17:T21").Select
    ActiveWorkbook.Names.Add Name:="Qj", RefersToR1C1:= _
        "='Data ve Notasyon'!R17C20:R21C20"
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    Range("T25:T29").Select
    ActiveWorkbook.Names.Add Name:="Sk", RefersToR1C1:= _
        "='Data ve Notasyon'!R25C20:R29C20"
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
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
    Range("L4:P12").Select
    ActiveWorkbook.Names.Add Name:="Xijt", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R4C12:R12C16"
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    Range("L17:P21").Select
    ActiveWorkbook.Names.Add Name:="Yjk", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R17C12:R21C16"
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
    Range("L26:O30").Select
    ActiveWorkbook.Names.Add Name:="Zkl", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R26C12:R30C15"
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
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    Range("V17:V21").Select
    ActiveWorkbook.Names.Add Name:="FÝj", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R17C22:R21C22"
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    Range("V25:V29").Select
    ActiveWorkbook.Names.Add Name:="DELTAk", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R25C22:R29C22"
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
End Sub
Sub AmacFonkisyonu()
Attribute AmacFonkisyonu.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AmacFonkisyonu Makro
'

'
    Range("U33").Select
    ActiveCell.FormulaR1C1 = "Amaç Fonksiyonu"
    Range("U33").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("V33").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Cijt*Xijt)+SUMPRODUCT(Cjk*Yjk)+SUMPRODUCT(Ckl*Zkl)+SUMPRODUCT(Qj*FÝj)+SUMPRODUCT(Sk*DELTAk)"
    Range("V34").Select
End Sub
