VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Tedarik�i Zinciri Y�netimi"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150.001
   OleObjectBlob   =   "DataNotasyonErisim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim parola As String
Dim ID As String
        ID = UserForm2.TextBox1.Value
        parola = UserForm2.TextBox2.Value
     
        If ID = "deu" And parola = "222" Then
            Call MsgBox("Sistemimize Ho�geldiniz. :)", , "Tedarik Zinciri Y�ntemi")
            Worksheets("Tedarik Zinciri Y�netimi").Unprotect
            Sheets("Data ve Notasyon").Visible = True
            Sheets("Ama� F. ve K�s�tlar").Visible = True
            Sheets("Karar Destek Sistemi").Visible = True
            Sheets("Data ve Notasyon").Select
            UserForm2.Hide
            UserForm2.TextBox1.Value = ""
            UserForm2.TextBox2.Value = ""
        Else
            Call MsgBox("Kullan�c� Ad� veya Parolan�z Hatal�d�r.L�tfen Tekrar Deneyiniz.", , "Tedarik Zinciri Y�ntemi")
            Sheets("Data ve Notasyon").Visible = False
            Sheets("Ama� F. ve K�s�tlar").Visible = False
            Sheets("Karar Destek Sistemi").Visible = False
            Worksheets("Tedarik Zinciri Y�netimi").Protect
            UserForm2.TextBox1.Value = ""
            UserForm2.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            'Bu�al��maKitab�.Close
            Sheets("Tedarik Zinciri Y�netimi").Select
            UserForm2.Hide
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
If closemode = vbFormControlMenu Then
    MsgBox "�zg�n�z.. 'Giri�' veya 'Geri' Butonlar�n� Kullan�n�z!"
    Cancel = True
End If
End Sub

