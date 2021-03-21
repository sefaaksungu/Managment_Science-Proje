VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub Kontrol_Gizle3()
Worksheets("Tedarik Zinciri Yönetimi").Protect
 Sheets("Amaç F. ve Kýsýtlar").Visible = False
 Sheets("Karar Destek Sistemi").Visible = False
 ActiveWindow.DisplayVerticalScrollBar = False
 ActiveWindow.DisplayHorizontalScrollBar = False
 Application.DisplayFormulaBar = False
 ActiveWindow.DisplayHeadings = False
 ActiveWindow.DisplayGridlines = False
 ActiveWindow.DisplayWorkbookTabs = False
 Application.DisplayStatusBar = False
End Sub
Private Sub CommandButton1_Click()
    Sheets("Amaç F. ve Kýsýtlar").Select
    SolverReset
    SolverOk SetCell:="$V$33", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$V$17:$V$21,$L$4:$P$12,$V$25:$V$29,$L$17:$P$21,$L$26:$O$30", Engine:=2, _
        EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$Q$4", Relation:=1, FormulaText:="$S$4"
    SolverAdd CellRef:="$Q$5", Relation:=1, FormulaText:="$S$5"
    SolverAdd CellRef:="$Q$6", Relation:=1, FormulaText:="$S$6"
    SolverAdd CellRef:="$Q$7", Relation:=1, FormulaText:="$S$7"
    SolverAdd CellRef:="$Q$8", Relation:=1, FormulaText:="$S$8"
    SolverAdd CellRef:="$Q$9", Relation:=1, FormulaText:="$S$9"
    SolverAdd CellRef:="$Q$10", Relation:=1, FormulaText:="$S$10"
    SolverAdd CellRef:="$Q$11", Relation:=1, FormulaText:="$S$11"
    SolverAdd CellRef:="$Q$12", Relation:=1, FormulaText:="$S$12"
    SolverAdd CellRef:="$Q$17", Relation:=1, FormulaText:="$S$17*$V$17"
    SolverAdd CellRef:="$Q$18", Relation:=1, FormulaText:="$S$18*$V$18"
    SolverAdd CellRef:="$Q$19", Relation:=1, FormulaText:="$S$19*$V$19"
    SolverAdd CellRef:="$Q$20", Relation:=1, FormulaText:="$S$20*$V$20"
    SolverAdd CellRef:="$Q$21", Relation:=1, FormulaText:="$S$21*$V$21"
    SolverAdd CellRef:="$P$26", Relation:=1, FormulaText:="$R$26*$V$25"
    SolverAdd CellRef:="$P$27", Relation:=1, FormulaText:="$R$27*$V$26"
    SolverAdd CellRef:="$P$28", Relation:=1, FormulaText:="$R$28*$V$27"
    SolverAdd CellRef:="$P$29", Relation:=1, FormulaText:="$R$29*$V$28"
    SolverAdd CellRef:="$P$30", Relation:=1, FormulaText:="$R$30*$V$29"
    SolverAdd CellRef:="$L$31", Relation:=3, FormulaText:="$L$33"
    SolverAdd CellRef:="$M$31", Relation:=3, FormulaText:="$M$33"
    SolverAdd CellRef:="$N$31", Relation:=3, FormulaText:="$N$33"
    SolverAdd CellRef:="$O$31", Relation:=3, FormulaText:="$O$33"
    SolverAdd CellRef:="$V$17:$V$21", Relation:=5, FormulaText:="ikili düzen"
    SolverAdd CellRef:="$V$22", Relation:=1, FormulaText:="$X$22"
    SolverAdd CellRef:="$V$25:$V$29", Relation:=5, FormulaText:="ikili düzen"
    SolverAdd CellRef:="$V$30", Relation:=1, FormulaText:="$X$30"
    SolverAdd CellRef:="$L$4:$P$12", Relation:=4, FormulaText:="tamsayý"
    SolverAdd CellRef:="$L$17:$P$21", Relation:=4, FormulaText:="tamsayý"
    SolverAdd CellRef:="$L$26:$O$30", Relation:=4, FormulaText:="tamsayý"
    SolverAdd CellRef:="$L$37:$P$37", Relation:=2, FormulaText:="0"
    SolverAdd CellRef:="$L$39:$P$41", Relation:=2, FormulaText:="0"
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
    SolverDelete CellRef:="$Q$4", Relation:=1, FormulaText:="$S$4"
    SolverDelete CellRef:="$Q$5", Relation:=1, FormulaText:="$S$5"
    SolverDelete CellRef:="$Q$6", Relation:=1, FormulaText:="$S$6"
    SolverDelete CellRef:="$Q$7", Relation:=1, FormulaText:="$S$7"
    SolverDelete CellRef:="$Q$8", Relation:=1, FormulaText:="$S$8"
    SolverDelete CellRef:="$Q$9", Relation:=1, FormulaText:="$S$9"
    SolverDelete CellRef:="$Q$10", Relation:=1, FormulaText:="$S$10"
    SolverDelete CellRef:="$Q$11", Relation:=1, FormulaText:="$S$11"
    SolverDelete CellRef:="$Q$12", Relation:=1, FormulaText:="$S$12"
    SolverDelete CellRef:="$Q$17", Relation:=1, FormulaText:="$S$17*$V$17"
    SolverDelete CellRef:="$Q$18", Relation:=1, FormulaText:="$S$18*$V$18"
    SolverDelete CellRef:="$Q$19", Relation:=1, FormulaText:="$S$19*$V$19"
    SolverDelete CellRef:="$Q$20", Relation:=1, FormulaText:="$S$20*$V$20"
    SolverDelete CellRef:="$Q$21", Relation:=1, FormulaText:="$S$21*$V$21"
    SolverDelete CellRef:="$P$26", Relation:=1, FormulaText:="$R$26*$V$25"
    SolverDelete CellRef:="$P$27", Relation:=1, FormulaText:="$R$27*$V$26"
    SolverDelete CellRef:="$P$28", Relation:=1, FormulaText:="$R$28*$V$27"
    SolverDelete CellRef:="$P$29", Relation:=1, FormulaText:="$R$29*$V$28"
    SolverDelete CellRef:="$P$30", Relation:=1, FormulaText:="$R$30*$V$29"
    SolverDelete CellRef:="$L$31", Relation:=3, FormulaText:="$L$33"
    SolverDelete CellRef:="$M$31", Relation:=3, FormulaText:="$M$33"
    SolverDelete CellRef:="$N$31", Relation:=3, FormulaText:="$N$33"
    SolverDelete CellRef:="$O$31", Relation:=3, FormulaText:="$O$33"
    SolverDelete CellRef:="$V$17:$V$21", Relation:=5, FormulaText:="ikili düzen"
    SolverDelete CellRef:="$V$22", Relation:=1, FormulaText:="$X$22"
    SolverDelete CellRef:="$V$25:$V$29", Relation:=5, FormulaText:="ikili düzen"
    SolverDelete CellRef:="$V$30", Relation:=1, FormulaText:="$X$30"
    SolverDelete CellRef:="$L$4:$P$12", Relation:=4, FormulaText:="tamsayý"
    SolverDelete CellRef:="$L$17:$P$21", Relation:=4, FormulaText:="tamsayý"
    SolverDelete CellRef:="$L$26:$O$30", Relation:=4, FormulaText:="tamsayý"
    SolverDelete CellRef:="$L$37:$P$37", Relation:=2, FormulaText:="0"
    SolverDelete CellRef:="$L$39:$P$41", Relation:=2, FormulaText:="0"
    Sheets("Karar Destek Sistemi").Select
    
End Sub

Private Sub Label1_Click()
Label1.Caption = "1. Müþteri Talebi"
End Sub

Private Sub Label10_Click()
Label10.Caption = "TL"
End Sub

Private Sub Label2_Click()
Label2.Caption = "2. Müþteri Talebi"
End Sub

Private Sub Label3_Click()
Label3.Caption = "3. Müþteri Talebi"
End Sub

Private Sub Label4_Click()
Label4.Caption = "4. Müþteri Talebi"
End Sub


Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Worksheets("Amaç F. ve Kýsýtlar").Cells(33, 22)
End Sub

Private Sub Label6_Click()
Label6.Caption = "Adet"
End Sub

Private Sub Label7_Click()
Label7.Caption = "Adet"
End Sub

Private Sub Label8_Click()
Label8.Caption = "Adet"
End Sub

Private Sub Label9_Click()
Label9.Caption = "Adet"
End Sub

Private Sub MusteriTalebi1_Change()
Worksheets("Data ve Notasyon").Cells(33, 19) = MusteriTalebi1.Value
End Sub

Private Sub MusteriTalebi2_Change()
Worksheets("Data ve Notasyon").Cells(34, 19) = MusteriTalebi2.Value
End Sub

Private Sub MusteriTalebi3_Change()
Worksheets("Data ve Notasyon").Cells(35, 19) = MusteriTalebi3.Value
End Sub

Private Sub MusteriTalebi4_Change()
Worksheets("Data ve Notasyon").Cells(36, 19) = MusteriTalebi4.Value
End Sub

Private Sub ToggleButton1_Click()
Sheets("Tedarik Zinciri Yönetimi").Select
Worksheets("Tedarik Zinciri Yönetimi").Protect
End Sub
