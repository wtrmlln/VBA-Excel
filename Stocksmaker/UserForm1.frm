VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Остатки"
   ClientHeight    =   7800
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12825
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox111_Click()

End Sub

Private Sub Autosave_Click()
    If UserForm1.Autosave = True Then
        UserForm1.Concatenation.Enabled = True
    ElseIf UserForm1.Autosave = False Then
        UserForm1.Concatenation.Enabled = False
    End If
End Sub

Private Sub CommandButton1_Click()

UserForm1.Hide
End Sub

Private Sub CommandButton2_Click()
Unload UserForm1
End Sub

Private Sub CommandButton3_Click()
UserForm2.Show
End Sub

Private Sub Concatenation_Click()

End Sub
