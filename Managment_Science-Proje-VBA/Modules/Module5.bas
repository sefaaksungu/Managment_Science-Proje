Attribute VB_Name = "Module5"
Sub TedarikciKapasiteKýsýtýSolTaraf()
'
' TedarikciKapasiteKýsýtýSolTaraf Makro
'

'
    ActiveCell.FormulaR1C1 = "=SUM(RC[-5]:RC[-1])"
    Range("Q4").Select
    Selection.AutoFill Destination:=Range("Q4:Q12"), Type:=xlFillDefault
    Range("Q4:Q12").Select
    Range("Q4").Select
    ActiveWorkbook.Names.Add Name:="X111tX121tX131tX141tX151", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R4C17"
    Range("Q5").Select
    ActiveWorkbook.Names.Add Name:="X112tX122tX132tX142tX152", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R5C17"
    Range("Q6").Select
    ActiveWorkbook.Names.Add Name:="X113tX123tX133tX143tX153", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R6C17"
    Range("Q7").Select
    ActiveWorkbook.Names.Add Name:="X211tX221tX231tX241tX251", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R7C17"
    Range("Q8").Select
    ActiveWorkbook.Names.Add Name:="X212tX222tX232tX242tX252", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R8C17"
    Range("Q9").Select
    ActiveWorkbook.Names.Add Name:="X213tX223tX233tX243tX253", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R9C17"
    Range("Q10").Select
    ActiveWorkbook.Names.Add Name:="X311tX321tX331tX341tX351", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R10C17"
    Range("Q11").Select
    ActiveWorkbook.Names.Add Name:="X312tX322tX332tX342tX352", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R11C17"
    Range("Q12").Select
    ActiveWorkbook.Names.Add Name:="X313tX323tX333tX343tX353", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R12C17"
End Sub

Sub TalepKisitiSolTarafDegeri()
Attribute TalepKisitiSolTarafDegeri.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TalepKisitiSolTarafDegeri Makro
'

'
    Range("L31").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    Range("L31").Select
    Selection.AutoFill Destination:=Range("L31:O31"), Type:=xlFillDefault
    Range("L31:O31").Select
    Range("L31").Select
    ActiveWorkbook.Names.Add Name:="Z11tZ21tZ31tZ41tZ51", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R31C12"
    Range("M31").Select
    ActiveWorkbook.Names.Add Name:="Z12tZ22tZ32tZ42tZ52", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R31C13"
    Range("N31").Select
    ActiveWorkbook.Names.Add Name:="Z13tZ23tZ33tZ43tZ53", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R31C14"
    Range("O31").Select
    ActiveWorkbook.Names.Add Name:="Z14tZ24tZ34tZ44tZ54", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R31C15"
    
End Sub
Sub FabrikaKapasiteSolTarafKisiti()
Attribute FabrikaKapasiteSolTarafKisiti.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FabrikaKapasiteSolTarafKisiti Makro
'

'
    ActiveCell.FormulaR1C1 = "=SUM(RC[-5]:RC[-1])"
    Range("Q17").Select
    Selection.AutoFill Destination:=Range("Q17:Q21"), Type:=xlFillDefault
    Range("Q17:Q21").Select
    Range("Q17").Select
    ActiveWorkbook.Names.Add Name:="Y11tY12tY13tY14tY15", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R17C17"
    Range("Q18").Select
    ActiveWorkbook.Names.Add Name:="Y21tY22tY23tY24tY25", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R18C17"
    Range("Q19").Select
    ActiveWorkbook.Names.Add Name:="Y31tY32tY33tY34tY35", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R19C17"
    Range("Q20").Select
    ActiveWorkbook.Names.Add Name:="Y41tY42tY43tY44tY45", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R20C17"
    Range("Q21").Select
    ActiveWorkbook.Names.Add Name:="Y51tY52tY53tY54tY55", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R21C17"
End Sub
Sub FabrikaSayýsýSolTarafKisiti()
Attribute FabrikaSayýsýSolTarafKisiti.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FabrikaSayýsýSolTarafKisiti Makro
'

'
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    Range("V22").Select
    ActiveWorkbook.Names.Add Name:="FÝ1tFÝ2tFÝ3tFÝ4tFÝ5", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R22C22"
End Sub
Sub DagitimMerkezSayisiSolTarafKisiti()
Attribute DagitimMerkezSayisiSolTarafKisiti.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DagitimMerkezSayisiSolTarafKisiti Makro
'

'
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    Range("V30").Select
    ActiveWorkbook.Names.Add Name:="DELTA1tDELTA2tDELTA3tDELTA4tDELTA5", _
        RefersToR1C1:="='Amaç F. ve Kýsýtlar'!R30C22"
    ActiveWorkbook.Save
    ActiveWorkbook.RunAutoMacros Which:=xlAutoClose
End Sub
