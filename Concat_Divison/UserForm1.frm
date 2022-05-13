VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Список файлов"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
If UserForm1.ListBox1.ListCount = 0 Then
MsgBox "Выберите файлы"
Exit Sub
End If
ws = 1
UserForm1.Hide
End Sub

Private Sub CommandButton2_Click()
Dim a As Variant
Dim lngCounter As Variant
StartingDir = Environ("OSTATKI")

ChDir StartingDir
a = Application.GetOpenFilename(MultiSelect:=True)
Application.DefaultFilePath = sStarDir
If IsArray(a) = True Then
    For Each lngCounter In a
        UserForm1.ListBox1.AddItem lngCounter
    Next lngCounter
Else:
    MsgBox "Файл не выбран"
Exit Sub
End If

End Sub

Private Sub CommandButton3_Click()
If UserForm1.ListBox1.ListCount = 0 Then
MsgBox "Выберите файлы"
Exit Sub
End If
ws = 2
UserForm1.Hide
End Sub

Private Sub CommandButton4_Click()
vyhod = 1
UserForm1.Hide
End Sub

Private Sub CommandButton5_Click()

End Sub
