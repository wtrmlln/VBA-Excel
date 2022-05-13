VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   6180
   ClientLeft      =   240
   ClientTop       =   945
   ClientWidth     =   5640
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctrl As Control
Dim MainElementName As String
Private Sub Autosave_Change()

    If UserForm1.������_������� = True Or UserForm1.���������_���������� = True Then
        UserForm1.Autosave = False
    End If
    
End Sub

Private Sub Cancel_Click()

    UserForm1.Cancel = True
    Unload Me
    
End Sub

Private Sub CheckBox1_Change()

    If UserForm1.������_������� = True Or UserForm1.���������_���������� = True Then
        UserForm1.CheckBox1 = False
    End If
    
End Sub

Private Sub CheckBox2_Change()

    If UserForm1.������_������� = True Or UserForm1.���������_���������� = True Then
        UserForm1.CheckBox2 = False
    End If
    
End Sub


Private Sub ok_Click()

    For Each Ctrl In UserForm1.Controls
        Ctrl.Enabled = True
    Next
    UserForm1.OptionButton1.Caption = "Default"
    UserForm1.Hide
    
End Sub

Private Sub OptionButton1_Change()
    UserForm1.TextBox3.Enabled = False
End Sub


Private Sub OptionButton2_Change()
    UserForm1.TextBox3.Enabled = True
End Sub


Private Sub OptionButton3_Change()

    UserForm1.TextBox3.Enabled = True
End Sub


Private Sub UserForm_Initialize()

If ActiveWorkbook.name = "������� ��� B2B.xlsb" Then

    For Each Ctrl In UserForm1.Controls
        Ctrl.Enabled = False
    Next
    UserForm1.TextBox1.Enabled = True
    UserForm1.Autosave.Enabled = True
    UserForm1.Autosave.Value = True
    UserForm1.ok.Enabled = True
    UserForm1.OptionButton1.Enabled = True
    
    UserForm1.OptionButton1.Caption = "��������� B2B"
End If

End Sub

Private Sub ������_�����_�_�����_Change()
    
    With Me
        MainElementName = Me.ActiveControl.name
        If Me.ActiveControl = True Then
            Disable_MainControls (MainElementName)
        ElseIf Me.ActiveControl = False Then
            Enable_MainControls (MainElementName)
        End If
    End With
    
End Sub

Private Sub ���������_����������_Click()
    If ���������_���������� = True Then
        CheckBox1.Enabled = False
        Autosave.Enabled = False
        Autosave.Value = False
        ������_�������.Enabled = False
        ��������_���������.Enabled = False
        �������_���������.Enabled = False
        OptionButton3.Enabled = False
        OptionButton4.Enabled = False
        TextBox3.Enabled = False
    ElseIf ���������_���������� = False Then
        CheckBox1.Enabled = True
        Autosave.Enabled = True
        ������_�������.Enabled = True
        ��������_���������.Enabled = True
        �������_���������.Enabled = True
        OptionButton3.Enabled = True
        OptionButton4.Enabled = True
        TextBox3.Enabled = True
    End If
End Sub

Private Sub ������_�������_Change()
    
    With Me
        MainElementName = Me.ActiveControl.name
        If Me.ActiveControl = True Then
            Disable_MainControls (MainElementName)
        ElseIf Me.ActiveControl = False Then
            Enable_MainControls (MainElementName)
        End If
    End With
    
End Sub

Private Sub ���������_���������_Change()

    With Me
        MainElementName = Me.ActiveControl.name
        If Me.ActiveControl = True Then
            Disable_MainControls (MainElementName)
        ElseIf Me.ActiveControl = False Then
            Enable_MainControls (MainElementName)
        End If
    End With
    
End Sub

Private Sub ��������_���������_Click()
    
    With Me
        MainElementName = Me.ActiveControl.name
        If Me.ActiveControl = True Then
            Disable_MainControls (MainElementName)
        ElseIf Me.ActiveControl = False Then
            Enable_MainControls (MainElementName)
        End If
    End With
    
End Sub

Private Sub �������_���������_Click()
    
    With Me
        MainElementName = Me.ActiveControl.name
        If Me.ActiveControl = True Then
            Disable_MainControls (MainElementName)
        ElseIf Me.ActiveControl = False Then
            Enable_MainControls (MainElementName)
        End If
    End With
    
End Sub

Private Function Disable_MainControls(MainElementName)
    
    For Each Ctrl In UserForm1.Controls
        If Ctrl.name = MainElementName Or Ctrl.name = "TextBox1" Or Ctrl.name = "ok" _
        Or Ctrl.name = "Cancel" Or TypeName(Ctrl) = "Label" Then
            Ctrl.Enabled = True
        Else
            Ctrl.Enabled = False
            If TypeName(Ctrl) = "CheckBox" Then
                Ctrl.Value = False
            End If
        End If
    Next
    
End Function

Private Function Enable_MainControls(MainElementName)
    
    For Each Ctrl In UserForm1.Controls
        If Ctrl.name <> MainElementName Or Ctrl.name <> "TextBox1" Or Ctrl.name <> "ok" _
        Or Ctrl.name <> "Cancel" Or TypeName(Ctrl) = "Label" Then
            Ctrl.Enabled = True
        End If
    Next
    
End Function

Private Function GetActiveControlName()
    
    With Me
        GetActiveControlName = Me.ActiveControl.name
    End With

End Function
