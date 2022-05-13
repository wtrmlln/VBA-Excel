VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Salemaker 
   Caption         =   "Акция"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
   OleObjectBlob   =   "UserForm_Salemaker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Salemaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
UserForm_Salemaker.Hide
End Sub

Private Sub CommandButton2_Click()
UserForm_Salemaker.CommandButton2.Cancel = True
UserForm_Salemaker.Hide
End Sub

Private Sub UserForm_Terminate()
'Unload UserForm_Salemaker
End Sub
