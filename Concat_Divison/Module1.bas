Attribute VB_Name = "Module1"
Public a, vyhod, mconc As Variant, shapka As Variant, abook, m(), delitel, ws, countrow, nof1

Public Sub division_concat()

    UserForm1.Caption = Left(ThisWorkbook.Name, 50)
    UserForm1.Show

    If vyhod = 1 Then GoTo ended

    If ws = 1 Then
        Call concat
    ElseIf ws = 2 Then
        Call division
    End If

ended:
    Unload UserForm1

End Sub


Private Sub concat()

    ReDim m(UserForm1.ListBox1.ListCount)
    i1 = 0
    Workbooks.Add
    abook = ActiveWorkbook.Name

    For i = 0 To UserForm1.ListBox1.ListCount - 1

        For Each wbk In Application.Workbooks
            If wbk.FullName = UserForm1.ListBox1.List(i) Then GoTo start
        Next wbk
    
        Workbooks.Open Filename:=UserForm1.ListBox1.List(i), local:=True
        m(i1) = ActiveWorkbook.Name
        i1 = i1 + 1
    
        If Workbooks(abook).ActiveSheet.Cells(1, 1).Value = "" Then
            mconc = ActiveWorkbook.ActiveSheet. _
            Range("A1:OH" & ActiveWorkbook.ActiveSheet. _
            Cells.SpecialCells(xlCellTypeLastCell).Row).Value
            Workbooks(abook).ActiveSheet.Range("A1:OH" & UBound(mconc, 1)).Value = mconc
        Else
            mconc = ActiveWorkbook.ActiveSheet. _
            Range("A2:OH" & ActiveWorkbook.ActiveSheet. _
            Cells.SpecialCells(xlCellTypeLastCell).Row).Value
            Workbooks(abook).ActiveSheet.Range("A" & Workbooks(abook).ActiveSheet. _
            Cells.SpecialCells(xlCellTypeLastCell).Row + 1 & ":OH" & Workbooks(abook).ActiveSheet. _
            Cells.SpecialCells(xlCellTypeLastCell).Row + UBound(mconc, 1)).Value = mconc
        End If
start:
    Next i


    If UserForm1.CheckBox1 = True Then
        If InStr(m(0), "Остат") = 0 Then
            Do While Dir(Environ("PROGRUZKA") & m(0) & "+concat " & Date & " " & nof1 & ".csv") <> ""
                nof1 = nof1 + 1
            Loop
            Workbooks(abook).SaveAs Filename:=Environ("PROGRUZKA") & m(0) & "+concat " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
        Else
            Do While Dir(Environ("OSTATKI") & m(0) & "+concat " & Date & " " & nof1 & ".csv") <> ""
                nof1 = nof1 + 1
            Loop
            Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & m(0) & "+concat " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
        End If
    End If


    For i = 0 To i1 - 1
        Workbooks(m(i)).Close False
    Next i

End Sub

Private Sub division()

    UserForm2.Show

    If IsNumeric(delitel) = False Then
        MsgBox "Введите корректное значение"
        Call division
    End If

    ReDim m(UserForm1.ListBox1.ListCount)
    i1 = 0

    For i = 0 To UserForm1.ListBox1.ListCount - 1

        For Each wbk In Application.Workbooks
            If wbk.FullName = UserForm1.ListBox1.List(i) Then GoTo start
        Next wbk

        Workbooks.Open Filename:=UserForm1.ListBox1.List(i), local:=True
    
        m(i1) = ActiveWorkbook.Name
        i1 = i1 + 1

        shapka = ActiveWorkbook.ActiveSheet.Range("A1:OH1").Value
        If ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row / delitel > 1 Then

            countrow = ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
            Rows("1:1").Select
            Selection.Delete Shift:=xlUp

            For i2 = 1 To WorksheetFunction.RoundUp((countrow + (WorksheetFunction.RoundUp((countrow) / delitel, 0) - 1)) / delitel, 0)

                Workbooks.Add
                a = ActiveWorkbook.Name

                mconc = Workbooks(m(i1 - 1)).ActiveSheet.Range("A" & (delitel - 1) * (i2 - 1) + 1 & ":OH" & (delitel - 1) * i2).Value
                Workbooks(a).ActiveSheet.Range("A1:OH" & delitel - 1).Value = mconc
                Rows("1:1").Select
                Selection.Insert Shift:=xlDown
                Workbooks(a).ActiveSheet.Range("A1:OH1").Value = shapka

                If UserForm1.CheckBox1 = True Then
                    If InStr(m(i1 - 1), "Остат") = 0 Then
                        Do While Dir(Environ("PROGRUZKA") & m(i1 - 1) & " " & i2 & " " & Date & " " & nof1 & ".csv") <> ""
                            nof1 = nof1 + 1
                        Loop
                        Workbooks(a).SaveAs Filename:=Environ("PROGRUZKA") & m(i1 - 1) & " " & i2 & " " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
                    Else
                        Do While Dir(Environ("OSTATKI") & m(i1 - 1) & " " & i2 & " " & Date & " " & nof1 & ".csv") <> ""
                            nof1 = nof1 + 1
                        Loop
                        Workbooks(a).SaveAs Filename:=Environ("OSTATKI") & m(i1 - 1) & " " & i2 & " " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
                    End If
                End If
            Next i2
        Else

            MsgBox "Для " & m(i1 - 1) & " разделение не требуется Кол. строк <= Делитель"

        End If
start:
    Next i

    For i = 0 To i1 - 1
        Workbooks(m(i)).Close False
    Next i

End Sub
