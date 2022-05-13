VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "¬вод"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3990
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox_Margin_Promo_Click()
If CheckBox_Margin_Promo = True Then
    UserForm4.TextBox1.Enabled = False
Else
    UserForm4.TextBox1.Enabled = True
End If
   
End Sub

Private Sub CommandButton1_Click()
    UserForm4.Hide
End Sub

