Attribute VB_Name = "Module4"
Sub DagitimMerkeziKapasiteSolTarafKisiti()
Attribute DagitimMerkeziKapasiteSolTarafKisiti.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DagitimMerkeziKapasiteSolTarafKisiti Makro
'

'
    ActiveCell.FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
    Range("P26").Select
    Selection.AutoFill Destination:=Range("P26:P30"), Type:=xlFillDefault
    Range("P26:P30").Select
    Range("P26").Select
    ActiveWorkbook.Names.Add Name:="Z11tZ12tZ13tZ14", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R26C16"
    Range("P27").Select
    ActiveWorkbook.Names.Add Name:="Z21tZ22tZ23tZ24", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R27C16"
    Range("P28").Select
    ActiveWorkbook.Names.Add Name:="Z31tZ32tZ33tZ34", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R28C16"
    Range("P29").Select
    ActiveWorkbook.Names.Add Name:="Z41tZ42tZ43tZ44", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R29C16"
    Range("P30").Select
    ActiveWorkbook.Names.Add Name:="Z51tZ52tZ53tZ54", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R30C16"
End Sub
