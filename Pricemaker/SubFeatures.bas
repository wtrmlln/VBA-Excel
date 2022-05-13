Attribute VB_Name = "SubFeatures"
Option Private Module
Option Explicit

''Окно "Выгрузить категории"
'Public Sub Выгрузить_категории()
'
'    Dim d
'    Dim a() As String
'    Dim mvvod()
'    Dim mvvodin() As String
'    Dim NRowsArray() As Long
'    Dim str As String
'    Dim TotalRows As Long
'    Dim i1 As Long
'    Dim r1 As Long
'    Dim i As Long
'
'    mvvodin = Split(UserForm1.TextBox1, vbCrLf)
'    If (UBound(mvvodin)) = -1 Then
'        GoTo ended
'    End If
'
'    ReDim mvvod(UBound(mvvodin))
'
'    For i = 0 To UBound(mvvodin)
'        mvvodin(i) = Application.Trim(mvvodin(i))
'        If Len(mvvodin(i)) = 0 Or mvvodin(i) = vbNullString Then
'            ReDim Preserve mvvod(UBound(mvvod) - 1)
'        Else
'            mvvod(i - (UBound(mvvodin) - UBound(mvvod))) = mvvodin(i)
'        End If
'    Next
'
'    If DublicateFind(mvvod) = True Then
'        GoTo ended
'    End If
'
'    ReDim NRowsArray(UBound(mvvod))
'
'    Workbooks.Add
'
'    ActiveWorkbook.ActiveSheet.Cells(1, 1).Value = "Код 1С"
'    ActiveWorkbook.ActiveSheet.Cells(1, 2).Value = "Категория"
'
'    For d = 0 To UBound(mvvod)
'        r1 = 1
'        str = mvvod(LBound(mvvod) + d)
'        Do While Trim(StrConv(ProschetArray(r1, 17), vbUpperCase)) _
'        <> Trim(StrConv(str, vbUpperCase))
'            r1 = r1 + 1
'        Loop
'        NRowsArray(d) = r1
'        a = Split(ProschetArray(NRowsArray(d), 7), ";")
'
'        ReDim mvvodin(UBound(a))
'        For i = 0 To UBound(a)
'            If a(i) = vbNullString Then
'                ReDim Preserve mvvodin(UBound(mvvodin) - 1)
'            Else
'                mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
'            End If
'        Next
'
'        If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
'            MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(d), 17)
'            GoTo ended
'        End If
'
'        TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows
'
'optimization1:
'        If d < UBound(mvvod) Then
'            If Trim(StrConv(ProschetArray(r1 + 1, 17), vbUpperCase)) = mvvod(LBound(mvvod) + d + 1) Then
'                r1 = r1 + 1
'                d = d + 1
'                NRowsArray(d) = r1
'
'                a = Split(ProschetArray(NRowsArray(d), 7), ";")
'
'                ReDim mvvodin(UBound(a))
'                For i = 0 To UBound(a)
'                    If a(i) = vbNullString Then
'                        ReDim Preserve mvvodin(UBound(mvvodin) - 1)
'                    Else
'                        mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
'                    End If
'                Next
'
'                If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
'                    MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(d), 17)
'                    GoTo ended
'                End If
'
'
'                GoTo optimization1
'            End If
'        End If
'
'    Next d
'
'    For i = 0 To UBound(NRowsArray, 1)
'        a = Split(ProschetArray(NRowsArray(i), 7), ";")
'
'        For i1 = 4 To UBound(a) Step 5
'            ActiveWorkbook.ActiveSheet.Cells(ActiveWorkbook.ActiveSheet.Cells. _
'            SpecialCells(xlCellTypeLastCell).Row + 1, 1).Value = ProschetArray(NRowsArray(i), 17)
'            ActiveWorkbook.ActiveSheet.Cells(ActiveWorkbook.ActiveSheet.Cells. _
'            SpecialCells(xlCellTypeLastCell).Row, 2).Value = Replace(a(i1), "П", vbNullString)
'        Next i1
'    Next i
'
'ended:
'    Unload UserForm1
'
'End Sub

Public Sub Добавить_удалить_компонент()
    
    Const TotalMargins As Integer = 4
    Const TotalMarginsKomponent As Integer = 5
    
    Dim TempUFKodsArray() As String
    Dim MarginCount As Long
    Dim i As Long
    Dim j As Long
    Dim Kod1C As String
    Dim MarginOldArray() As String
    Dim KomponentTextBox As String
    Dim KomponentsCount As Long
    Dim NewMarginArray() As String
    Dim MarginArray() As String
    Dim NewTorudaKomponent As Boolean
    Dim ChangeKomponentValue As Long
    Dim UFCount As Long
    Dim NewMargin As String
    Dim FillStep As Long
    Dim DublicateKomponent As Boolean
    
    TempUFKodsArray = Split(UserForm1.TextBox1, vbCrLf)

    MarginCount = -1
    If (UBound(TempUFKodsArray)) = -1 Then
        GoTo ended
    Else
        ReDim UFKodsArray(UBound(TempUFKodsArray))
        For i = 0 To UBound(TempUFKodsArray)
            If TempUFKodsArray(i) <> vbNullString Or Len(TempUFKodsArray(i)) <> 0 Then
                MarginCount = MarginCount + 1
                UFKodsArray(MarginCount) = TempUFKodsArray(i)
            End If
        Next i
        ReDim Preserve UFKodsArray(MarginCount)
    End If
    
    If DublicateFind(UFKodsArray) = True Then
        GoTo ended
    End If
    
    ReDim NRowsArray(UBound(UFKodsArray))
    Dim ErrorCount As Long
    For i = 0 To UBound(UFKodsArray)
        
        Kod1C = Trim(UFKodsArray(i))
        NRowsArray(i) = Proschet_Dict.Item(Kod1C)
        MarginOldArray = Split(ProschetArray(NRowsArray(i), 7), ";")
        MarginCount = -1
        
        ReDim MarginClearedArray(UBound(MarginOldArray))
        For j = 0 To UBound(MarginOldArray)
            If MarginOldArray(j) <> vbNullString Or Len(MarginOldArray(j)) <> 0 Then
                MarginCount = MarginCount + 1
                MarginClearedArray(MarginCount) = MarginOldArray(j)
            Else
                ErrorCount = ErrorCount + 1
            End If
        Next j
        ReDim Preserve MarginClearedArray(MarginCount)
        
        If ErrorCount > 0 Then
            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = Join(MarginClearedArray, ";")
            ProschetArray(NRowsArray(i), 7) = Join(MarginClearedArray, ";")
        End If
         
        If (UBound(MarginClearedArray) + 1) Mod TotalMarginsKomponent <> 0 Then
            MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(i), 17)
            GoTo ended
        End If

    Next i
    
    UserForm5.Show
    If UserForm1.Добавить_компонент = True Then
        KomponentTextBox = UserForm5.TextBox1.Value
        If KomponentTextBox <> vbNullString Then
    
            For i = 0 To UBound(NRowsArray, 1)
    
                MarginArray = Split(ProschetArray(NRowsArray(i), 7), ";")
                ReDim Preserve MarginArray(UBound(MarginArray) + TotalMarginsKomponent)
                KomponentsCount = (UBound(MarginArray) \ TotalMarginsKomponent) + 1
    
                If InStr(KomponentTextBox, "t") <> 0 Then
                    NewTorudaKomponent = True
                Else
                    NewTorudaKomponent = False
                End If
    
                ChangeKomponentValue = 0
                DublicateKomponent = False
                NewMargin = vbNullString
                For j = 0 To KomponentsCount
                    'Если новый компонент - Торуда, берется наценка из первого другого компонента Торуда
                    If NewTorudaKomponent = True Then
                        If InStr(MarginArray(ChangeKomponentValue), "t") <> 0 Then
    
                            If MarginArray(ChangeKomponentValue) = KomponentTextBox Then
                                DublicateKomponent = True
                                Exit For
                            End If
                            
                            If j <= 0 Then
                                j = TotalMargins
    
                            Else
                                j = ChangeKomponentValue - TotalMargins
                            End If
    
                            UFCount = 0
                            FillStep = UBound(MarginArray) + 1 - TotalMarginsKomponent
                            Do While (j + 1) Mod TotalMarginsKomponent <> 0
    
                                MarginArray(FillStep) = MarginArray(j)
                                FillStep = FillStep + 1
                                j = j + 1
                            Loop
    
                            MarginArray(UBound(MarginArray)) = KomponentTextBox
                            NewMargin = Join(MarginArray, ";")
                            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
                            Exit For
    
                        Else
                            If ChangeKomponentValue = 0 Then
                                ChangeKomponentValue = ChangeKomponentValue + TotalMargins
                            Else
                                ChangeKomponentValue = ChangeKomponentValue + TotalMarginsKomponent
                            End If
                        End If
                    
                    'Если новый компонент не Торуда - берется наценка из любого первого компонента не Торуда
                    ElseIf NewTorudaKomponent = False Then
                        If InStr(MarginArray(ChangeKomponentValue), "t") = 0 Then
                            
                            If MarginArray(ChangeKomponentValue) = KomponentTextBox Then
                                DublicateKomponent = True
                                Exit For
                            End If
                            
                            If j <= 0 Then
                                j = 0
                            Else
                                j = ChangeKomponentValue - TotalMargins
                            End If
    
                            UFCount = 0
                            FillStep = UBound(MarginArray) + 1 - TotalMarginsKomponent
                            Do While (j + 1) Mod TotalMarginsKomponent <> 0
    
                                MarginArray(FillStep) = MarginArray(j)
                                FillStep = FillStep + 1
                                j = j + 1
                            Loop
    
                            MarginArray(UBound(MarginArray)) = KomponentTextBox
                            NewMargin = Join(MarginArray, ";")
                            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
                            Exit For
    
                        Else
                            If ChangeKomponentValue = 0 Then
                                ChangeKomponentValue = ChangeKomponentValue + TotalMargins
                            Else
                                ChangeKomponentValue = ChangeKomponentValue + TotalMarginsKomponent
                            End If
                        End If
                    End If
                Next
                
                'Если новый компонент Торуда, и в наценке нет других компонентов Торуда - вычитается 0.05 из первого любого компонента
                If NewMargin = vbNullString And DublicateKomponent = False Then
                    
                    ChangeKomponentValue = 0

                    For j = 0 To KomponentsCount

                        If InStr(MarginArray(ChangeKomponentValue), "t") = 0 Then

                            If j <= 0 Then
                                j = 0
                            Else
                                j = ChangeKomponentValue - TotalMargins
                            End If

                            UFCount = 0
                            FillStep = UBound(MarginArray) + 1 - TotalMarginsKomponent
                            Do While (j + 1) Mod TotalMarginsKomponent <> 0

                                MarginArray(FillStep) = MarginArray(j) - 0.05
                                FillStep = FillStep + 1
                                j = j + 1
                            Loop

                            MarginArray(UBound(MarginArray)) = KomponentTextBox
                            NewMargin = Join(MarginArray, ";")
                            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
                            Exit For

                        Else
                            If ChangeKomponentValue = 0 Then
                                ChangeKomponentValue = ChangeKomponentValue + TotalMargins
                            Else
                                ChangeKomponentValue = ChangeKomponentValue + TotalMarginsKomponent
                            End If
                        End If
                    Next j
                End If
            Next
        End If
    ElseIf UserForm1.Удалить_компонент = True Then
        KomponentTextBox = UserForm5.TextBox1.Value
        For i = 0 To UBound(NRowsArray, 1)
            MarginArray = Split(ProschetArray(NRowsArray(i), 7), ";")
            KomponentsCount = (UBound(MarginArray) \ TotalMarginsKomponent) + 1
            ChangeKomponentValue = 0
            For j = 0 To KomponentsCount
                If MarginArray(ChangeKomponentValue) = KomponentTextBox Then

                    If j <= 0 Then
                        j = TotalMargins

                    Else
                        j = ChangeKomponentValue - TotalMargins
                    End If

                    UFCount = 0
                    FillStep = 0
                    
                    If UBound(MarginArray) - TotalMarginsKomponent < 0 Then
                        ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = vbNullString
                        Exit For
                    Else
                        ReDim Preserve NewMarginArray(UBound(MarginArray) - TotalMarginsKomponent)
                    End If
                    
                    Dim x As Long
                    For x = 0 To UBound(MarginArray)
                        If x = j Then
                            x = x + TotalMargins
                        Else
                            NewMarginArray(FillStep) = MarginArray(x)
                            FillStep = FillStep + 1
                        End If
                            
                    Next x
                    
                    NewMargin = Join(NewMarginArray, ";")
                    ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
                    Exit For

                Else
                    If ChangeKomponentValue = 0 Then
                        ChangeKomponentValue = ChangeKomponentValue + TotalMargins
                    Else
                        ChangeKomponentValue = ChangeKomponentValue + TotalMarginsKomponent
                    End If
                End If
            Next
        Next
    End If

ended:
    Unload UserForm1
    Unload UserForm5
End Sub

'Окно "Замена наценок"
Public Sub Замена_наценок()
    
    Const TotalMargins As Integer = 4
    Const TotalMarginsKomponent As Integer = 5
    Dim NewMargin As String
    Dim UFKodsArray() As String
    Dim UFMarginArray() As String
    Dim UFValue As String
    Dim TempUFKodsArray() As String
    Dim MarginOldArray() As String
    Dim MarginClearedArray() As String
    Dim MarginArray() As String
    Dim NRowsArray() As Long
    Dim Kod1C As String
    Dim MarginCount As Long
    Dim e As Long
    Dim i As Long
    Dim j As Long
    Dim ChangeKomponentValue As Long
    Dim UFCount As Long
    
    TempUFKodsArray = Split(UserForm1.TextBox1, vbCrLf)
    
    MarginCount = -1
    If (UBound(TempUFKodsArray)) = -1 Then
        GoTo ended
    Else
        ReDim UFKodsArray(UBound(TempUFKodsArray))
        For i = 0 To UBound(TempUFKodsArray)
            If TempUFKodsArray(i) <> vbNullString Or Len(TempUFKodsArray(i)) <> 0 Then
                MarginCount = MarginCount + 1
                UFKodsArray(MarginCount) = TempUFKodsArray(i)
            End If
        Next i
        ReDim Preserve UFKodsArray(MarginCount)
    End If
    
    If DublicateFind(UFKodsArray) = True Then
        GoTo ended
    End If

    ReDim NRowsArray(UBound(UFKodsArray))
    Dim ErrorCount As Long
    For i = 0 To UBound(UFKodsArray)
        
        Kod1C = Trim(UFKodsArray(i))
        NRowsArray(i) = Proschet_Dict.Item(Kod1C)
        MarginOldArray = Split(ProschetArray(NRowsArray(i), 7), ";")
        MarginCount = -1
        
        ReDim MarginClearedArray(UBound(MarginOldArray))
        For j = 0 To UBound(MarginOldArray)
            If MarginOldArray(j) <> vbNullString Or Len(MarginOldArray(j)) <> 0 Then
                MarginCount = MarginCount + 1
                MarginClearedArray(MarginCount) = MarginOldArray(j)
            Else
                ErrorCount = ErrorCount + 1
            End If
        Next j
        ReDim Preserve MarginClearedArray(MarginCount)
        
        If ErrorCount > 0 Then
            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = Join(MarginClearedArray, ";")
            ProschetArray(NRowsArray(i), 7) = Join(MarginClearedArray, ";")
        End If
         
        If (UBound(MarginClearedArray) + 1) Mod TotalMarginsKomponent <> 0 Then
            MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(i), 17)
            GoTo ended
        End If

    Next i
 
start_uf:
    UserForm3.Show
    
    'Сложить - вычесть
    If UserForm3.CheckBox3 = True Or UserForm3.CheckBox4 = True Then
        GoTo start
    'Заменить
    ElseIf UserForm3.CheckBox5 = True Then
        UFMarginArray = Split(UserForm3.TextBox1, ";")
    Else
        GoTo start_uf
    End If
    
    MarginCount = -1
    For i = 0 To UBound(UFMarginArray)
        If UFMarginArray(i) <> vbNullString Or Len(UFMarginArray(i)) <> 0 Then
            MarginCount = MarginCount + 1
            UFMarginArray(MarginCount) = UFMarginArray(i)
        End If
    Next i
    ReDim Preserve UFMarginArray(MarginCount)

    If UBound(UFMarginArray) <> 3 Then
        MsgBox "Некорректная наценка/категория"
        GoTo ended
    End If

start:
    
    Dim KomponentTextBox As String
    Dim KomponentsCount As Long
    KomponentTextBox = UserForm3.TextBox2.Value
    
    If UserForm3.CheckBox3 = True Or UserForm3.CheckBox4 = True Then
        UFValue = UserForm3.TextBox1.Value
        
        If InStr(UFValue, ".") <> 0 Then
            UFValue = Replace(UFValue, ".", ",")
        End If
    End If
    
    If KomponentTextBox <> vbNullString Then
        
        For i = 0 To UBound(NRowsArray, 1)
            MarginArray = Split(ProschetArray(NRowsArray(i), 7), ";")
            KomponentsCount = (UBound(MarginArray) \ TotalMarginsKomponent) + 1
            ChangeKomponentValue = 0
            For j = 0 To KomponentsCount
                If MarginArray(ChangeKomponentValue) = KomponentTextBox Then
                    j = ChangeKomponentValue - TotalMargins
                    UFCount = 0
                    Do While (j + 1) Mod TotalMarginsKomponent <> 0
                        
                        If UserForm3.CheckBox3 = True Then
                            MarginArray(j) = CDbl(MarginArray(j)) + CDbl(UFValue)
                        ElseIf UserForm3.CheckBox4 = True Then
                            MarginArray(j) = CDbl(MarginArray(j)) - CDbl(UFValue)
                        ElseIf UserForm3.CheckBox5 = True Then
                            MarginArray(j) = UFMarginArray(UFCount)
                            UFCount = UFCount + 1
                        End If
                        
                        j = j + 1
                        
                    Loop
                    NewMargin = Join(MarginArray, ";")
                    ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
                    Exit For
                Else
                    If ChangeKomponentValue = 0 Then
                        ChangeKomponentValue = ChangeKomponentValue + TotalMargins
                    Else
                        ChangeKomponentValue = ChangeKomponentValue + TotalMarginsKomponent
                    End If
                End If
            Next
        Next
    Else
        
        
        For i = 0 To UBound(NRowsArray, 1)
            MarginArray = Split(ProschetArray(NRowsArray(i), 7), ";")
            e = 0
            j = 0
            Do While j <> UBound(MarginArray)
                If (j + 1) Mod TotalMarginsKomponent <> 0 Then
                    
                    If UserForm3.CheckBox3 = True Then
                        MarginArray(j) = CDbl(MarginArray(j)) + CDbl(UFValue)
                    ElseIf UserForm3.CheckBox4 = True Then
                        MarginArray(j) = CDbl(MarginArray(j)) - CDbl(UFValue)
                    ElseIf UserForm3.CheckBox5 = True Then
                        MarginArray(j) = UFMarginArray(j - (UBound(UFMarginArray) * e + e * 2))
                    End If
                    j = j + 1
                Else
                    j = j + 1
                    e = e + 1
                End If
            Loop
            NewMargin = Join(MarginArray, ";")
            ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = NewMargin
        Next
    End If

ended:
    Unload UserForm1
    Unload UserForm3
End Sub

'Окно "Выгрузить компоненты"
Public Sub Выгрузить_компоненты()

    Dim OnlyPromo As Boolean
    Dim KomponentsStr As String
    Dim NRowsArray() As Long
    Dim KomponentsArray() As String
    Dim i As Long
    Dim j As Long
    Dim CountRows As Long
    
    Dim ResultArray() As String
    ReDim NRowsArray(0)
    
    If UserForm1.OptionButton2 = True Then
        OnlyPromo = True
    ElseIf UserForm1.OptionButton1 = True Then
        OnlyPromo = False
    End If
    
    If OnlyPromo = True Then
        For i = 1 To UBound(ProschetArray, 1)
            KomponentsStr = ProschetArray(i, 7)
            If InStr(KomponentsStr, "П") <> 0 Then
                NRowsArray(UBound(NRowsArray, 1)) = i
                ReDim Preserve NRowsArray(UBound(NRowsArray, 1) + 1)
            End If
        Next i
        
        ReDim ResultArray(1, 0)

        For i = 0 To UBound(NRowsArray, 1) - 1
            KomponentsStr = ProschetArray(NRowsArray(i), 7)
            KomponentsArray = Split(KomponentsStr, ";")
            For j = 0 To UBound(KomponentsArray)
                If InStr(KomponentsArray(j), "П") <> 0 Then
                    ResultArray(0, UBound(ResultArray, 2)) = ProschetArray(NRowsArray(i), 17)
                    ResultArray(1, UBound(ResultArray, 2)) = KomponentsArray(j)
                    ReDim Preserve ResultArray(1, UBound(ResultArray, 2) + 1)
                End If
            Next j
        Next i
    ElseIf OnlyPromo = False Then
           
        For i = 1 To UBound(ProschetArray, 1)
            KomponentsStr = ProschetArray(i, 7)
            NRowsArray(UBound(NRowsArray, 1)) = i
            ReDim Preserve NRowsArray(UBound(NRowsArray, 1) + 1)
        Next i

        ReDim ResultArray(1, 0)

        For i = 0 To UBound(NRowsArray, 1) - 1
            KomponentsStr = ProschetArray(NRowsArray(i), 7)
            KomponentsArray = Split(KomponentsStr, ";")
            For j = 4 To UBound(KomponentsArray) Step 5
                ResultArray(0, UBound(ResultArray, 2)) = ProschetArray(NRowsArray(i), 17)
                ResultArray(1, UBound(ResultArray, 2)) = KomponentsArray(j)
                ReDim Preserve ResultArray(1, UBound(ResultArray, 2) + 1)
            Next j
        Next i
    End If
        
    Workbooks.Add
    ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Код 1С"
    ActiveWorkbook.ActiveSheet.Cells(1, 2) = "Категория"
    For i = 0 To UBound(ResultArray, 2)
        ActiveWorkbook.ActiveSheet.Cells(i + 2, 1) = ResultArray(0, i)
        ActiveWorkbook.ActiveSheet.Cells(i + 2, 2) = ResultArray(1, i)
    Next i

    Unload UserForm1

End Sub

'Создает файл с изменениями через Reference
'Public Sub Изменения_Тепложар(ByRef mteplotoxls, ByRef mteplo, ByRef TriggerTeplo)
'
'    ReDim mteplotoxls(2, 0)
'
'    Dim referencedict, mreference As Variant
'    Dim nof2                                     'Счетчик для наименования в шапке файла "Изменения"
'
'    mreference = Workbooks(ProschetWorkbookName).Worksheets("Reference").Range("A1:NColumnProgruzka" & Workbooks(ProschetWorkbookName).Worksheets("Reference"). _
'    Cells.SpecialCells(xlCellTypeLastCell).Row)
'
'    Set referencedict = CreateObject("Scripting.Dictionary")
'
'    For i = 1 To UBound(mreference, 1)
'        If Trim(mreference(i, 1)) <> vbNullString Then
'            If referencedict.exists(Trim(mreference(i, 1))) = True Then
'                MsgBox "Дубль в Reference! " & mreference(i, 1)
'                Exit Sub
'            End If
'            referencedict.Add Trim(mreference(i, 1)), i
'        End If
'    Next i
'
'    For i = 0 To UBound(mteplo, 2) - 1
'
'
'        If referencedict.exists(mteplo(0, i)) = False Then GoTo exitif
'
'        If IsError(mreference(referencedict(mteplo(0, i)), 3)) = True Then
'            mreference(referencedict(mteplo(0, i)), 3) = 0
'        End If
'
'        If IsError(mteplo(1, i)) = True Then
'            mteplo(1, i) = 0
'        End If
'
'
'        If mreference(referencedict(mteplo(0, i)), 3) = 0 And mteplo(1, i) = 0 Or _
'        IsNumeric(mreference(referencedict(mteplo(0, i)), 3)) = False And mteplo(1, i) = 0 Or _
'        IsNumeric(mreference(referencedict(mteplo(0, i)), 3)) = False And IsNumeric(mteplo(1, i)) = False Or _
'        mreference(referencedict(mteplo(0, i)), 3) = 0 And IsNumeric(mteplo(1, i)) = False Then
'            GoTo exitif
'        Else
'            If mteplo(1, i) = 0 Or IsNumeric(mteplo(1, i)) = False Then
'                Workbooks(ProschetWorkbookName).Worksheets("Reference").Cells(referencedict(mteplo(0, i)), 3).Value = 0
'                TriggerTeplo = 2
'                If UBound(mteplotoxls, 2) <> 0 Then
'                    mteplotoxls(0, UBound(mteplotoxls, 2)) = mteplo(0, i)
'                    mteplotoxls(1, UBound(mteplotoxls, 2)) = "-"
'                    mteplotoxls(2, UBound(mteplotoxls, 2)) = mteplo(2, i)
'                    ReDim Preserve mteplotoxls(2, UBound(mteplotoxls, 2) + 1)
'                Else
'                    mteplotoxls(0, UBound(mteplotoxls, 2)) = mteplo(0, i)
'                    mteplotoxls(1, UBound(mteplotoxls, 2)) = "-"
'                    mteplotoxls(2, UBound(mteplotoxls, 2)) = mteplo(2, i)
'                    ReDim Preserve mteplotoxls(2, UBound(mteplotoxls, 2) + 1)
'                End If
'
'            ElseIf mreference(referencedict(mteplo(0, i)), 3) = 0 Or IsNumeric(mreference(referencedict(mteplo(0, i)), 3)) = False Then
'                Workbooks(ProschetWorkbookName).Worksheets("Reference").Cells(referencedict(mteplo(0, i)), 3).Value = mteplo(1, i)
'                TriggerTeplo = 2
'                If UBound(mteplotoxls, 2) <> 0 Then
'                    mteplotoxls(0, UBound(mteplotoxls, 2)) = mteplo(0, i)
'                    mteplotoxls(1, UBound(mteplotoxls, 2)) = mteplo(1, i)
'                    mteplotoxls(2, UBound(mteplotoxls, 2)) = mteplo(2, i)
'                    ReDim Preserve mteplotoxls(2, UBound(mteplotoxls, 2) + 1)
'                Else
'                    mteplotoxls(0, UBound(mteplotoxls, 2)) = mteplo(0, i)
'                    mteplotoxls(1, UBound(mteplotoxls, 2)) = mteplo(1, i)
'                    mteplotoxls(2, UBound(mteplotoxls, 2)) = mteplo(2, i)
'                    ReDim Preserve mteplotoxls(2, UBound(mteplotoxls, 2) + 1)
'                End If
'            End If
'        End If
'exitif:
'
'    Next i
'
'
'    If TriggerTeplo = 2 Then
'        Workbooks.Add
'
'        ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Наименование"
'        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "Код 1С"
'        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "Цена"
'
'        For i = 2 To UBound(mteplotoxls, 2) + 1
'            ActiveWorkbook.ActiveSheet.Cells(i, 1) = ProschetArray(mteplotoxls(2, i - 2), 2)
'            ActiveWorkbook.ActiveSheet.Cells(i, 2) = mteplotoxls(0, i - 2)
'            ActiveWorkbook.ActiveSheet.Cells(i, 3) = mteplotoxls(1, i - 2)
'
'            If i >= 3 Then
'                If ProschetArray(mteplotoxls(2, i - 2), 1) <> ProschetArray(mteplotoxls(2, i - 3), 1) Then
'                    nof2 = nof2 & " " & ProschetArray(mteplotoxls(2, i - 2), 1)
'                End If
'            Else
'                nof2 = ProschetArray(mteplotoxls(2, i - 2), 1)
'            End If
'
'        Next i
'        If UserForm1.Autosave = True Then
'           Number = Number + 1
'           Do While Dir(Environ("TEPLOMAN") & "Изменения " & nof2 & " " & Date & " " & Number & ".csv") <> vbNullString
'               Number = Number + 1
'           Loop
'        End If
'        ActiveWorkbook.SaveAs Filename:=Environ("TEPLOMAN") & "Изменения " & nof2 & " " & Date & " " & Number & ".csv", FileFormat:=xlCSV, CreateBackup:=False
'
'    End If
'
'End Sub

'Поставить на акцию
Sub Salemaker()
    Dim i As Long
    Dim AkciaArray As Variant
    Dim mvvod() As String
    Dim DateEndAkc As Date
    Dim NRowAKcia As Long
    Dim kody() As String
    Dim dict
    Dim zakup() As String
    Dim RetailPrice() As String
    Dim ProschetArray As Variant
    Dim lLastRow As Long
    Dim provdict As Object
    Dim DictAkcia As Object
    Dim WsProschetAkcia As Object

    For Each wks In Application.Worksheets
        If wks.name = "Акция" Then
            GoTo start
        End If
    Next wks

    Call CreateSheet("Акция")

start:

    AkciaArray = Application.Worksheets("Акция").Range("A1:F" & _
    Application.Worksheets("Акция").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.ScreenUpdating = False
    For i = UBound(AkciaArray, 1) To 1 Step -1
        If Application.Worksheets("Акция").Cells(i, 1) = vbNullString Then
            Application.Worksheets("Акция").Rows(i).Delete
        End If
    Next i
    Application.ScreenUpdating = True

    lLastRow = Application.Worksheets("Акция").UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    AkciaArray = Application.Worksheets("Акция").Range("A1:F" & lLastRow)


    'Ключ - Код 1С из листа "Акция", элемент - номер строки
    Set DictAkcia = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(AkciaArray, 1)

        If Trim(AkciaArray(i, 1)) <> "" Then
            DictAkcia.Add Trim(AkciaArray(i, 1)), i
        End If

    Next i

    UserForm_Salemaker.Show

    If UserForm_Salemaker.CommandButton2.Cancel = True Then
        Unload UserForm_Salemaker
        Exit Sub
    End If

    If UserForm1.TextBox1 = "" Then
        MsgBox "Введите коды 1С"
        GoTo start
    ElseIf UserForm_Salemaker.TextBox3 = "" Then
        MsgBox "Введите цены"
        GoTo start
    ElseIf IsDate(UserForm_Salemaker.TextBox4) = False Or Not UserForm_Salemaker.TextBox4 Like "##.##.####" Then
        MsgBox "Некорректная дата отключения ДД.ММ.ГГГГ"
        GoTo start
    End If

    'Записывает коды 1С в массив
    If InStr(UserForm1.TextBox1, vbCrLf) <> 0 Then
        mvvod = Split(UserForm1.TextBox1, vbCrLf)
    Else
        ReDim mvvod(0)
        mvvod(0) = UserForm1.TextBox1
    End If

    ReDim kody(UBound(mvvod))

    'Удаляет последний лишний разрыв строки
    For i = 0 To UBound(mvvod)
        If Trim(mvvod(i)) = "" Then
            ReDim Preserve kody(UBound(kody) - 1)
        Else
            kody(i - (UBound(mvvod) - UBound(kody))) = Trim(mvvod(i))
        End If
    Next

    'Проверка на повторяющиеся коды 1С в юзерформе
    Set provdict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(kody, 1)
        If provdict.exists(kody(i)) = False Then
            provdict.Add kody(i), i
        Else
            MsgBox kody(i) & " Дубль!"
            GoTo start
        End If
    Next

'    Проверка на повторяющиеся коды 1С в листе "Акция" и юзерформе
    For i = 0 To UBound(kody, 1)
        If DictAkcia.exists(kody(i)) = True Then
            MsgBox kody(i) & " Уже в таблице!"
            GoTo start
        End If
    Next

    'Записывает РРЦ из юзерформы в массив
    If Trim(UserForm_Salemaker.TextBox2) <> "" Then
        If InStr(UserForm_Salemaker.TextBox2, vbCrLf) <> 0 Then
            mvvod = Split(UserForm_Salemaker.TextBox2, vbCrLf)
        Else
            ReDim mvvod(0)
            mvvod(0) = UserForm_Salemaker.TextBox2
        End If

        ReDim RetailPrice(UBound(mvvod))

        For i = 0 To UBound(mvvod)
            If Trim(mvvod(i)) = "" Then
                ReDim Preserve RetailPrice(UBound(RetailPrice) - 1)
            Else
                RetailPrice(i - (UBound(mvvod) - UBound(RetailPrice))) = mvvod(i)
            End If
        Next
    End If

    'Записывает ОПТ из юзерформы в массив
    If InStr(UserForm_Salemaker.TextBox3, vbCrLf) <> 0 Then
        mvvod = Split(UserForm_Salemaker.TextBox3, vbCrLf)
    Else
        ReDim mvvod(0)
        mvvod(0) = UserForm_Salemaker.TextBox3
    End If

    ReDim zakup(UBound(mvvod))

    For i = 0 To UBound(mvvod)
        If Trim(mvvod(i)) = "" Then
            ReDim Preserve zakup(UBound(zakup) - 1)
        Else
            zakup(i - (UBound(mvvod) - UBound(zakup))) = mvvod(i)
        End If
    Next

    If UBound(kody, 1) <> UBound(zakup, 1) Then
        MsgBox "Количество цен - не равно количеству кодов 1С"
        GoTo start
    End If

    ProschetArray = ActiveWorkbook.Worksheets("Просчет цен").Range("Q1:S" & _
    ActiveWorkbook.Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    'Ключ - Код 1С, элемент - номер строки
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(ProschetArray, 1)
        If Trim(ProschetArray(i, 1)) <> "" Then
            dict.Add Trim(ProschetArray(i, 1)), i
        End If
    Next i

    For i = 0 To UBound(kody, 1)

        If dict.exists(kody(i)) = False Then
            MsgBox kody(i) & " Не найден!"
            GoTo start
        End If

    Next

    If Trim(UserForm_Salemaker.TextBox2) <> "" Then
        If UBound(kody, 1) <> UBound(RetailPrice, 1) Then
            MsgBox "Количество РРЦ - не равно количеству кодов 1С"
            GoTo start
        End If
    End If

    For i = 0 To UBound(kody, 1)
        Do While InStr(kody(i), Chr(9)) <> 0
            kody(i) = Replace(kody(i), Chr(9), "")
        Loop
    Next i

    DateEndAkc = CDate(UserForm_Salemaker.TextBox4)
    Set WsProschetAkcia = ActiveWorkbook.Worksheets("Акция")

    For i = 0 To UBound(kody, 1)
    
        'lLastRow =
        'NRowAkcia - номер строки для записи в лист "Акция"
        If (WsProschetAkcia.UsedRange.Row + WsProschetAkcia.UsedRange.Rows.Count - 1) = 1 And _
        WsProschetAkcia.Cells(1, 1) = "" Then
            NRowAKcia = 1
        Else
            NRowAKcia = WsProschetAkcia.UsedRange.Row + WsProschetAkcia.UsedRange.Rows.Count
        End If
    
        'Запись кода 1С в лист "Акция"
        WsProschetAkcia.Cells(NRowAKcia, 1).Value = kody(i)
    
        'Запись текущей РРЦ из просчета в лист "Акция"
        If IsError(ProschetArray(dict(kody(i)), 2)) = True Then
            WsProschetAkcia.Cells(NRowAKcia, 2).Value = 0
        Else
            WsProschetAkcia.Cells(NRowAKcia, 2).Value = ProschetArray(dict(kody(i)), 2)
        End If
        
        'Запись текущего ОПТ из просчета в лист "Акция"
        If IsError(ProschetArray(dict(kody(i)), 3)) = True Then
            WsProschetAkcia.Cells(NRowAKcia, 3).Value = 0
        Else
            WsProschetAkcia.Cells(NRowAKcia, 3).Value = ProschetArray(dict(kody(i)), 3)
        End If
    
        'Запись даты окончания акции в лист "Акция"
        WsProschetAkcia.Cells(NRowAKcia, 4).Value = DateEndAkc
    
        'Запись акционной РРЦ в лист "Акция"
        If Trim(UserForm_Salemaker.TextBox2) <> "" Then
            WsProschetAkcia.Cells(NRowAKcia, 5).Value = RetailPrice(i)
        End If
    
        'Запись акционного ОПТ в лист "Акция"
        WsProschetAkcia.Cells(NRowAKcia, 6).Value = zakup(i)
    
        'Запись акционного РРЦ и ОПТ в просчет
        If Trim(UserForm_Salemaker.TextBox2) <> "" Then
            ActiveWorkbook.Worksheets("Просчет цен").Cells(dict(kody(i)), 18).Value = RetailPrice(i)
        End If
        ActiveWorkbook.Worksheets("Просчет цен").Cells(dict(kody(i)), 19).Value = zakup(i)

    Next i

    Unload UserForm_Salemaker

End Sub


