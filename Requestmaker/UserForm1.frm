VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Запрос прайсов"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6780
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox1_Click()

    If UserForm1.CheckBox1 = True Then
        UserForm1.CheckBox2.Enabled = False
        UserForm1.CheckBox3.Enabled = False
        UserForm1.CheckBox4.Enabled = False
        UserForm1.CheckBox5.Enabled = False
    Else
        UserForm1.CheckBox2.Enabled = True
        UserForm1.CheckBox3.Enabled = True
        UserForm1.CheckBox4.Enabled = True
        UserForm1.CheckBox5.Enabled = True
    End If

End Sub

Private Sub CheckBox2_Click()

    If UserForm1.CheckBox2 = True Then
        UserForm1.CheckBox1.Enabled = False
        UserForm1.CheckBox3.Enabled = False
        UserForm1.CheckBox4.Enabled = False
        UserForm1.CheckBox5.Enabled = False
    Else
        UserForm1.CheckBox1.Enabled = True
        UserForm1.CheckBox3.Enabled = True
        UserForm1.CheckBox4.Enabled = True
        UserForm1.CheckBox5.Enabled = True
    End If

End Sub

Private Sub CheckBox3_Click()

    If UserForm1.CheckBox3 = True Then
        UserForm1.CheckBox1.Enabled = False
        UserForm1.CheckBox2.Enabled = False
        UserForm1.CheckBox4.Enabled = False
        UserForm1.CheckBox5.Enabled = False
    Else
        UserForm1.CheckBox1.Enabled = True
        UserForm1.CheckBox2.Enabled = True
        UserForm1.CheckBox4.Enabled = True
        UserForm1.CheckBox5.Enabled = True
    End If

End Sub

Private Sub CheckBox4_Click()

    If UserForm1.CheckBox4 = True Then
        UserForm1.CheckBox1.Enabled = False
        UserForm1.CheckBox2.Enabled = False
        UserForm1.CheckBox3.Enabled = False
        UserForm1.CheckBox5.Enabled = False
    Else
        UserForm1.CheckBox1.Enabled = True
        UserForm1.CheckBox2.Enabled = True
        UserForm1.CheckBox3.Enabled = True
        UserForm1.CheckBox5.Enabled = True
    End If
    
End Sub

Private Sub CheckBox5_Click()

    If UserForm1.CheckBox5 = True Then
        UserForm1.CheckBox1.Enabled = False
        UserForm1.CheckBox2.Enabled = False
        UserForm1.CheckBox3.Enabled = False
        UserForm1.CheckBox4.Enabled = False
    Else
        UserForm1.CheckBox1.Enabled = True
        UserForm1.CheckBox2.Enabled = True
        UserForm1.CheckBox3.Enabled = True
        UserForm1.CheckBox4.Enabled = True
    End If

End Sub

Private Sub CommandButton1_Click()
    UserForm1.Hide
End Sub

Private Sub CommandButton2_Click()
    IsCancelled = True
    UserForm1.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    IsCancelled = True
    UserForm1.Hide
End Sub
