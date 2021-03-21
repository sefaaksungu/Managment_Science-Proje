VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Tedarikçi Zinciri Yönetimi"
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
            Call MsgBox("Sistemimize Hoþgeldiniz. :)", , "Tedarik Zinciri Yöntemi")
            Worksheets("Tedarik Zinciri Yönetimi").Unprotect
            Sheets("Amaç F. ve Kýsýtlar").Visible = True
            Sheets("Karar Destek Sistemi").Visible = True
            Sheets("Karar Destek Sistemi").Select
            UserForm1.Hide
            UserForm1.TextBox1.Value = ""
            UserForm1.TextBox2.Value = ""

        Else
            Call MsgBox("Kullanýcý Adý veya Parolanýz Hatalýdýr.Lütfen Tekrar Deneyiniz.", , "Tedarik Zinciri Yöntemi")
            Sheets("Amaç F. ve Kýsýtlar").Visible = False
            Sheets("Karar Destek Sistemi").Visible = False
            Worksheets("Tedarik Zinciri Yönetimi").Protect
            UserForm1.TextBox1.Value = ""
            UserForm1.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            'BuÇalýþmaKitabý.Close
            Sheets("Tedarik Zinciri Yönetimi").Select
            UserForm1.Hide
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
If closemode = vbFormControlMenu Then
    MsgBox "Üzgünüz.. 'Giriþ' veya 'Geri' Butonlarýný Kullanýnýz!"
    Cancel = True
End If
End Sub

