VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4350
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
UserForm2.Hide
    
    Dim InputArray() As String
    Dim UFKods_Array() As String
    Dim AddedWb As Object
    Dim AddedWs As Object
    Dim i As Long
    Dim Sklad_Value As Variant
    Dim Ostatki_Value As Variant
    Dim Sklad_Array() As Variant
    Dim Ostatki_Array() As Variant
    Dim nof As Long
    Dim sEmpty As Variant
    
    InputArray() = Split(UserForm2.TextBox1, vbCrLf)
    If UserForm2.TextBox2 = sEmpty Then
        Sklad_Value = ""
    Else
        Sklad_Value = UserForm2.TextBox2
    End If
    
    If UserForm2.TextBox3 = sEmpty Then
        Ostatki_Value = ""
    Else
        Ostatki_Value = UserForm2.TextBox3
    End If
    'Sklad_Value = UserForm2.TextBox2
    'Ostatki_Value = UserForm2.TextBox3
    
    Unload UserForm2
    
    If (UBound(InputArray)) = -1 Then
        Exit Sub
    End If
    
    ReDim UFKods_Array(UBound(InputArray))
    
    For i = 0 To UBound(InputArray)
        InputArray(i) = Application.Trim(InputArray(i))
        If InputArray(i) = vbNullString Then
            ReDim Preserve UFKods_Array(UBound(UFKods_Array) - 1)
        Else
            UFKods_Array(i - (UBound(InputArray) - UBound(UFKods_Array))) = InputArray(i)
        End If
    Next
    
    ReDim Sklad_Array(UBound(UFKods_Array))
    ReDim Ostatki_Array(UBound(UFKods_Array))
    
    For i = 0 To UBound(UFKods_Array)
        Sklad_Array(i) = Sklad_Value
        Ostatki_Array(i) = Ostatki_Value
    Next
    
    Application.ScreenUpdating = False
    
    ''''''''Склад''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Workbooks.Add
    Set AddedWs = ActiveWorkbook.ActiveSheet
    AddedWs.Cells(1, 1) = "Sklad"
    AddedWs.Cells(1, 2) = "Kod1C"
    For i = 0 To UBound(UFKods_Array)
        AddedWs.Cells(i + 2, 1) = Sklad_Array(i)
        AddedWs.Cells(i + 2, 2) = UFKods_Array(i)
    Next i
    Set AddedWs = Nothing
    Set AddedWb = ActiveWorkbook

    Do While Dir(Environ("OSTATKI") & "Склад Постоянный " & Date & " " & nof & ".csv") <> sEmpty
        nof = nof + 1
    Loop
    AddedWb.SaveAs Filename:=Environ("OSTATKI") & "Склад Постоянный " & Date & " " & nof & ".csv", FileFormat:=xlCSV

    
    ''''''''Склад ПМК''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Workbooks.Add
    Set AddedWs = ActiveWorkbook.ActiveSheet
    AddedWs.Cells(1, 1) = "Kod1C"
    AddedWs.Cells(1, 2) = "Sklad"
    For i = 0 To UBound(UFKods_Array)
        AddedWs.Cells(i + 2, 1) = UFKods_Array(i)
        AddedWs.Cells(i + 2, 2) = Sklad_Array(i)
    Next i

    Set AddedWs = Nothing
    Set AddedWb = ActiveWorkbook

    Do While Dir(Environ("OSTATKI") & "Склад ПМК Постоянный " & Date & " " & nof & ".csv") <> sEmpty
        nof = nof + 1
    Loop
    AddedWb.SaveAs Filename:=Environ("OSTATKI") & "Склад ПМК Постоянный " & Date & " " & nof & ".csv", FileFormat:=xlCSV
    
    
    ''''''''Остатки''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Workbooks.Add
    Set AddedWs = ActiveWorkbook.ActiveSheet
    AddedWs.Cells(1, 1) = "Kod1C"
    AddedWs.Cells(1, 2) = "Ostatki"
    AddedWs.Range("A2:A" & (UBound(UFKods_Array) + 2)).Value = UFKods_Array()
    For i = 0 To UBound(UFKods_Array)
        AddedWs.Cells(i + 2, 1) = UFKods_Array(i)
        AddedWs.Cells(i + 2, 2) = Ostatki_Array(i)
    Next i
    Set AddedWs = Nothing
    Set AddedWb = ActiveWorkbook
    
    Do While Dir(Environ("OSTATKI") & "Остатки Постоянные " & Date & " " & nof & ".csv") <> sEmpty
        nof = nof + 1
    Loop
    AddedWb.SaveAs Filename:=Environ("OSTATKI") & "Остатки Постоянные " & Date & " " & nof & ".csv", FileFormat:=xlCSV
    
    
    ''''''''Остатки КК''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Workbooks.Add
    Set AddedWs = ActiveWorkbook.ActiveSheet
    AddedWs.Cells(1, 1) = "Category"
    AddedWs.Cells(1, 2) = "Kod1C"
    AddedWs.Cells(1, 3) = "Ostatki"
    AddedWs.Range("B2:B" & (UBound(UFKods_Array) + 2)).Value = UFKods_Array()
    For i = 0 To UBound(UFKods_Array)
        AddedWs.Cells(i + 2, 2) = UFKods_Array(i)
        AddedWs.Cells(i + 2, 3) = Ostatki_Array(i)
    Next i

    Set AddedWs = Nothing
    Set AddedWb = ActiveWorkbook

    Do While Dir(Environ("OSTATKI") & "Остатки КК Постоянные " & Date & " " & nof & ".csv") <> sEmpty
        nof = nof + 1
    Loop
    AddedWb.SaveAs Filename:=Environ("OSTATKI") & "Остатки КК Постоянные " & Date & " " & nof & ".csv", FileFormat:=xlCSV


    Application.ScreenUpdating = True
    
    Unload UserForm1

End Sub
