VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Tedarik�i Zinciri Y�netimi"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "KararDestekSistemineErisim.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim parola As String
Dim ID As String
        ID = UserForm1.TextBox1.Value
        parola = UserForm1.TextBox2.Value
     
        If ID = "deu" And parola = "111" Then
            Call MsgBox("Sistemimize Ho�geldiniz. :)", , "Tedarik Zinciri Y�ntemi")
            Worksheets("Tedarik Zinciri Y�netimi").Unprotect
            Sheets("Ama� F. ve K�s�tlar").Visible = True
            Sheets("Karar Destek Sistemi").Visible = True
            Sheets("Karar Destek Sistemi").Select
            UserForm1.Hide
            UserForm1.TextBox1.Value = ""
            UserForm1.TextBox2.Value = ""

        Else
            Call MsgBox("Kullan�c� Ad� veya Parolan�z Hatal�d�r.L�tfen Tekrar Deneyiniz.", , "Tedarik Zinciri Y�ntemi")
            Sheets("Ama� F. ve K�s�tlar").Visible = False
            Sheets("Karar Destek Sistemi").Visible = False
            Worksheets("Tedarik Zinciri Y�netimi").Protect
            UserForm1.TextBox1.Value = ""
            UserForm1.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            'Bu�al��maKitab�.Close
            Sheets("Tedarik Zinciri Y�netimi").Select
            UserForm1.Hide
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
If closemode = vbFormControlMenu Then
    MsgBox "�zg�n�z.. 'Giri�' veya 'Geri' Butonlar�n� Kullan�n�z!"
    Cancel = True
End If
End Sub

