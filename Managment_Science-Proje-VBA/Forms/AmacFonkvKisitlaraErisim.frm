VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Tedarik�i Zinciri Y�netimi"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775.001
   OleObjectBlob   =   "AmacFonkvKisitlaraErisim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
Dim parola As String
Dim ID As String
        ID = UserForm3.TextBox1.Value
        parola = UserForm3.TextBox2.Value
     
        If ID = "deu" And parola = "333" Then
            Call MsgBox("Sistemimize Ho�geldiniz. :)", , "Tedarik Zinciri Y�ntemi")
            Worksheets("Tedarik Zinciri Y�netimi").Unprotect
            Sheets("Data ve Notasyon").Visible = True
            Sheets("Ama� F. ve K�s�tlar").Visible = True
            Sheets("Karar Destek Sistemi").Visible = True
            Sheets("Ama� F. ve K�s�tlar").Select
            UserForm3.Hide
            UserForm3.TextBox1.Value = ""
            UserForm3.TextBox2.Value = ""

        Else
            Call MsgBox("Kullan�c� Ad� veya Parolan�z Hatal�d�r.L�tfen Tekrar Deneyiniz.", , "Tedarik Zinciri Y�ntemi")
            Sheets("Data ve Notasyon").Visible = False
            Sheets("Ama� F. ve K�s�tlar").Visible = False
            Sheets("Karar Destek Sistemi").Visible = False
            Worksheets("Tedarik Zinciri Y�netimi").Protect
            UserForm3.TextBox1.Value = ""
            UserForm3.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            'Bu�al��maKitab�.Close
            Sheets("Tedarik Zinciri Y�netimi").Select
            UserForm3.Hide
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
If closemode = vbFormControlMenu Then
    MsgBox "�zg�n�z.. 'Giri�' veya 'Geri' Butonlar�n� Kullan�n�z!"
    Cancel = True
End If
End Sub

