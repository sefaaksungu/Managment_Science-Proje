Attribute VB_Name = "Module3"
Sub FÝj_Adlandýrýlmasý()
Attribute FÝj_Adlandýrýlmasý.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FÝj_Adlandýrýlmasý Makro
'

'
    ActiveWorkbook.Names.Add Name:="FÝ.1", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R17C22"
    Range("V18").Select
    ActiveWorkbook.Names.Add Name:="FÝ.2", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R18C22"
    Range("V19").Select
    ActiveWorkbook.Names.Add Name:="FÝ.3", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R19C22"
    Range("V20").Select
    ActiveWorkbook.Names.Add Name:="FÝ.4", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R20C22"
    Range("V21").Select
    ActiveWorkbook.Names.Add Name:="FÝ.5", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R21C22"
End Sub
Sub DELTAk_Adlandýrýlmasý()
Attribute DELTAk_Adlandýrýlmasý.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DELTAk_Adlandýrýlmasý Makro
'

'
    ActiveWorkbook.Names.Add Name:="DELTA.1", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R25C22"
    Range("V26").Select
    ActiveWorkbook.Names.Add Name:="DELTA.2", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R26C22"
    Range("V27").Select
    ActiveWorkbook.Names.Add Name:="DELTA.3", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R27C22"
    Range("V28").Select
    ActiveWorkbook.Names.Add Name:="DELTA.4", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R28C22"
    Range("V29").Select
    ActiveWorkbook.Names.Add Name:="DELTA.5", RefersToR1C1:= _
        "='Amaç F. ve Kýsýtlar'!R29C22"
End Sub
