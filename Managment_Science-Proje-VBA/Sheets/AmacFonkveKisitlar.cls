VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub Kontrol_Gizle()
Worksheets("Tedarik Zinciri Y�netimi").Protect
 Sheets("Ama� F. ve K�s�tlar").Visible = False
 Sheets("Karar Destek Sistemi").Visible = False
 ActiveWindow.DisplayVerticalScrollBar = False
 ActiveWindow.DisplayHorizontalScrollBar = False
 Application.DisplayFormulaBar = False
 ActiveWindow.DisplayHeadings = False
 ActiveWindow.DisplayGridlines = False
 ActiveWindow.DisplayWorkbookTabs = False
 Application.DisplayStatusBar = False

End Sub
 
Private Sub ToggleButton1_Click()
Sheets("Tedarik Zinciri Y�netimi").Select
Worksheets("Tedarik Zinciri Y�netimi").Protect
End Sub