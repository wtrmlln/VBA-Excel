VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Ввод"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4170
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox3_Click()

    If UserForm3.CheckBox3 = True Then
        UserForm3.Label1.Caption = "Укажите число которое необходимо добавить к наценке"
        UserForm3.CheckBox4.Enabled = False
        UserForm3.CheckBox5.Enabled = False
    Else
        UserForm3.Label1.Caption = vbNullString
        UserForm3.CheckBox4.Enabled = True
        UserForm3.CheckBox5.Enabled = True
    End If

End Sub

Private Sub CheckBox4_Click()

    If UserForm3.CheckBox4 = True Then
        UserForm3.Label1.Caption = "Укажите число которое необходимо вычесть из наценки"
        UserForm3.CheckBox3.Enabled = False
        UserForm3.CheckBox5.Enabled = False
    Else
        UserForm3.Label1.Caption = vbNullString
        UserForm3.CheckBox3.Enabled = True
        UserForm3.CheckBox5.Enabled = True
    End If

End Sub

Private Sub CheckBox5_Click()
    
    If UserForm3.CheckBox5 = True Then
        UserForm3.Label1.Caption = "Наценки формата X,XX;X,XX;X,XX;X,XX"
        UserForm3.CheckBox3.Enabled = False
        UserForm3.CheckBox4.Enabled = False
    Else
        UserForm3.Label1.Caption = vbNullString
        UserForm3.CheckBox3.Enabled = True
        UserForm3.CheckBox4.Enabled = True
    End If
    
End Sub

Private Sub CheckBox6_Click()
    
    If UserForm3.CheckBox6 = True Then
        UserForm3.TextBox2.Enabled = True
        UserForm3.TextBox2.BackColor = &HFFFFFF
    ElseIf UserForm3.CheckBox6 = False Then
        UserForm3.TextBox2.Enabled = False
        UserForm3.TextBox2.BackColor = &H8000000F
        
    End If
    
End Sub

Private Sub CommandButton1_Click()
    UserForm3.Hide
End Sub

Private Sub Cancel_Click()
    Unload UserForm3
    End
End Sub

