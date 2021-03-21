VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Tedarikçi Zinciri Yönetimi"
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
            Call MsgBox("Sistemimize Hoþgeldiniz. :)", , "Tedarik Zinciri Yöntemi")
            Worksheets("Tedarik Zinciri Yönetimi").Unprotect
            Sheets("Data ve Notasyon").Visible = True
            Sheets("Amaç F. ve Kýsýtlar").Visible = True
            Sheets("Karar Destek Sistemi").Visible = True
            Sheets("Data ve Notasyon").Select
            UserForm2.Hide
            UserForm2.TextBox1.Value = ""
            UserForm2.TextBox2.Value = ""
        Else
            Call MsgBox("Kullanýcý Adý veya Parolanýz Hatalýdýr.Lütfen Tekrar Deneyiniz.", , "Tedarik Zinciri Yöntemi")
            Sheets("Data ve Notasyon").Visible = False
            Sheets("Amaç F. ve Kýsýtlar").Visible = False
            Sheets("Karar Destek Sistemi").Visible = False
            Worksheets("Tedarik Zinciri Yönetimi").Protect
            UserForm2.TextBox1.Value = ""
            UserForm2.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            'BuÇalýþmaKitabý.Close
            Sheets("Tedarik Zinciri Yönetimi").Select
            UserForm2.Hide
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
If closemode = vbFormControlMenu Then
    MsgBox "Üzgünüz.. 'Giriþ' veya 'Geri' Butonlarýný Kullanýnýz!"
    Cancel = True
End If
End Sub

