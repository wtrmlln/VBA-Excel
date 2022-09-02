Attribute VB_Name = "B2B"
Option Private Module
Option Explicit

Public Sub Прогрузка_Б2Б()
    Dim i As Long
    Dim j As Long
    Dim tempArray() As String
    Dim ClearedTempArray() As String
    Dim Kod1C As String
    Dim WbProschet As Object
    Dim UFKods_Array() As String
    
    Dim NRowArray() As Long
    Dim KURS_TENGE As Double
    Dim KURS_BELRUB As Double
    Dim KURS_SOM As Double
    Dim KURS_DRAM As Double
    
    Dim TorudaArray() As String
    Dim KABArray() As String
    Dim PMKArray() As String
    Dim LedosvetArray() As String

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Proschet() As Variant                    'Активный лист от А1:T До последней заполненной ячейки + 1
    Dim ProgruzkaPMK() As Variant
    Dim ProgruzkaPMK_2() As Variant
    Dim ProgruzkaToruda() As Variant
    Dim ProgruzkaKab() As Variant
    Dim ProgruzkaLedosvet() As Variant
    
    Dim CityForMarginOutput As String
    Dim MarginOutput As String
    
    Set WbProschet = ActiveWorkbook
    
    'Proschet = Активная книга. Активный лист. От A1:T до последней заполненной ячейки + 1
    Proschet = WbProschet.ActiveSheet.Range("A1:T" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
    
    Set Proschet_Dict = CreateObject("Scripting.Dictionary")
    
    'Проверка кодов 1С на дубли в просчете
    'И добавление в словарь Proschet_Dict: Ключ - Код 1С из листа "Просчет цен", Элемент - № строки
    For i = 1 To UBound(Proschet, 1)
        Proschet(i, 17) = Trim(Proschet(i, 17))
        If Len(Proschet(i, 17)) <> 0 Then
            If Proschet_Dict.exists(Proschet(i, 17)) = True Then
                MsgBox "Дубль в просчете! " & Proschet(i, 17)
                Exit Sub
            End If
            Proschet_Dict.Add Proschet(i, 17), i
        End If
    Next i
    
    'Константы из просчета:
    SNG_MARGIN = Proschet(7, 16)
    KURS_TENGE = Proschet(1, 16)
    KURS_BELRUB = Proschet(2, 16)
    KURS_SOM = Proschet(3, 16)
    KURS_DRAM = Proschet(4, 16)
    
    UserForm1.Caption = Left(ThisWorkbook.name, 19)
    UserForm1.Show
    
    If UserForm1.Вывести_наценку = True Then
        'UserForm1.Hide
        UserForm6.Show
        CityForMarginOutput = UserForm6.TextBox1.Text
    End If
    
    Dim InputArray() As String

    InputArray = Split(UserForm1.TextBox1, vbCrLf)
    
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
    
    If (Not UFKods_Array) = -1 Then
        MsgBox "Коды отсутствуют в юзерформе"
        Exit Sub
    End If
    
    If DublicateFind(UFKods_Array) = True Then
        Exit Sub
    End If
     
    ReDim NRowArray(UBound(UFKods_Array))
    
    Dim TotalRowsTORUDA As Long
    Dim TotalRowsKAB As Long
    Dim TotalRowsPMK As Long
    Dim TotalRowsLEDOSVET As Long
    
    For i = 0 To UBound(UFKods_Array)
    
        Kod1C = Trim(UFKods_Array(LBound(UFKods_Array) + i))
        NRowArray(i) = Proschet_Dict.Item(Kod1C)
        
        tempArray = Split(Proschet(NRowArray(i), 7), ";")
        
        ReDim ClearedTempArray(UBound(tempArray))
        
        For j = 0 To UBound(tempArray)
            tempArray(j) = Trim(tempArray(j))
            If tempArray(j) = vbNullString Then
                ReDim Preserve ClearedTempArray(UBound(ClearedTempArray) - 1)
            Else
                ClearedTempArray(j - (UBound(tempArray) - UBound(ClearedTempArray))) = tempArray(j)
            End If
        Next
        
        If (UBound(ClearedTempArray) + 1) Mod 4 = 0 Then
            For j = 3 To UBound(ClearedTempArray) Step 4
                
                'Поиск сайта по компоненту - Торуда, Кабинетоф, ПМК +
                'Запись компонентов и наценок по сайтам в разные массивы.
                If InStr(ClearedTempArray(j), "t") <> 0 Then
                    TotalRowsTORUDA = TotalRowsTORUDA + 1
                    
                    ReDim Preserve TorudaArray(4, TotalRowsTORUDA)
                    
                    TorudaArray(4, TotalRowsTORUDA) = NRowArray(i)
                    TorudaArray(3, TotalRowsTORUDA) = ClearedTempArray(j)
                    TorudaArray(2, TotalRowsTORUDA) = ClearedTempArray(j - 1)
                    TorudaArray(1, TotalRowsTORUDA) = ClearedTempArray(j - 2)
                    TorudaArray(0, TotalRowsTORUDA) = ClearedTempArray(j - 3)
                    
                ElseIf InStr(ClearedTempArray(j), "Kab") <> 0 Then
                    TotalRowsKAB = TotalRowsKAB + 1
                    
                    ReDim Preserve KABArray(4, TotalRowsKAB)
                    
                    KABArray(4, TotalRowsKAB) = NRowArray(i)
                    KABArray(3, TotalRowsKAB) = ClearedTempArray(j)
                    KABArray(2, TotalRowsKAB) = ClearedTempArray(j - 1)
                    KABArray(1, TotalRowsKAB) = ClearedTempArray(j - 2)
                    KABArray(0, TotalRowsKAB) = ClearedTempArray(j - 3)
                    
                ElseIf InStr(ClearedTempArray(j), "611") <> 0 Then
                    TotalRowsLEDOSVET = TotalRowsLEDOSVET + 1
                    
                    ReDim Preserve LedosvetArray(4, TotalRowsLEDOSVET)
                    
                    LedosvetArray(4, TotalRowsLEDOSVET) = NRowArray(i)
                    LedosvetArray(3, TotalRowsLEDOSVET) = ClearedTempArray(j)
                    LedosvetArray(2, TotalRowsLEDOSVET) = ClearedTempArray(j - 1)
                    LedosvetArray(1, TotalRowsLEDOSVET) = ClearedTempArray(j - 2)
                    LedosvetArray(0, TotalRowsLEDOSVET) = ClearedTempArray(j - 3)
                    
                Else
                    TotalRowsPMK = TotalRowsPMK + 1
                    
                    ReDim Preserve PMKArray(4, TotalRowsPMK)
                    
                    PMKArray(4, TotalRowsPMK) = NRowArray(i)
                    PMKArray(3, TotalRowsPMK) = ClearedTempArray(j)
                    PMKArray(2, TotalRowsPMK) = ClearedTempArray(j - 1)
                    PMKArray(1, TotalRowsPMK) = ClearedTempArray(j - 2)
                    PMKArray(0, TotalRowsPMK) = ClearedTempArray(j - 3)
                End If
            Next
        Else
            MsgBox "Некорректная наценка/категория " & Proschet(NRowArray(i), 17)
            Exit Sub
        End If
               
    Next
    Erase NRowArray
    
    If TotalRowsPMK > 0 Then
        For i = 0 To UBound(PMKArray, 2)
            If PMKArray(3, i) = "830" Or PMKArray(3, i) = "917" Then
                Dim TotalRowsPMK_2 As Variant
                TotalRowsPMK_2 = TotalRowsPMK_2 + 1
            End If
        Next
    End If
    
    If UserForm1.Cancel.Cancel = True Then
        Unload UserForm1
        Exit Sub
    End If
    
    If UserForm1.OptionButton1 = True Then
        ReDim ProgruzkaPMK_2(1 To (TotalRowsPMK_2 + 1), 1 To 102)
        ReDim ProgruzkaPMK(1 To (TotalRowsPMK + 1), 1 To 102)
        ReDim ProgruzkaToruda(1 To (TotalRowsTORUDA + 1), 1 To 147)
        ReDim ProgruzkaKab(1 To (TotalRowsKAB + 1), 1 To 110)
        ReDim ProgruzkaLedosvet(1 To (TotalRowsLEDOSVET + 1), 1 To 127)
    End If

    'Создание шапки прогрузки и массива с наименованиями городов
    'и перенос значений из листов с наценками в массивы
    If TotalRowsPMK <> 0 Then
        Dim CitiesArrayPMK() As String
        Dim WsMarginsPMK() As Variant
        ReDim CitiesArrayPMK(1 To 98)
        ProgruzkaPMK = CreateHeader(ProgruzkaPMK, "Б2Б ПМК")
        ProgruzkaPMK_2 = CreateHeader(ProgruzkaPMK_2, "Б2Б ПМК")
        CitiesArrayPMK = CreateCities(CitiesArrayPMK, "Б2Б ПМК")
        WsMarginsPMK = WbProschet.Worksheets("ПМК").Range("A2:CX" & _
        WbProschet.Worksheets("ПМК").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
    End If
    
    If TotalRowsTORUDA <> 0 Then
        Dim CitiesArrayToruda() As String
        Dim WsMarginsToruda() As Variant
        ReDim CitiesArrayToruda(1 To 143)
        ProgruzkaToruda = CreateHeader(ProgruzkaToruda, "Б2Б Торуда")
        CitiesArrayToruda = CreateCities(CitiesArrayToruda, "Б2Б Торуда")
        WsMarginsToruda = WbProschet.Worksheets("Торуда").Range("A2:EQ" & _
        WbProschet.Worksheets("Торуда").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
    End If
    
    If TotalRowsKAB <> 0 Then
        Dim CitiesArrayKab() As String
        Dim WsMarginsKab() As Variant
        ReDim CitiesArrayKab(1 To 106)
        ProgruzkaKab = CreateHeader(ProgruzkaKab, "Б2Б Кабинетоф")
        CitiesArrayKab = CreateCities(CitiesArrayKab, "Б2Б Кабинетоф")
        WsMarginsKab = WbProschet.Worksheets("Кабинетоф").Range("A2:DF" & _
        WbProschet.Worksheets("Кабинетоф").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
    End If
    
    If TotalRowsLEDOSVET <> 0 Then
        Dim CitiesArrayLedosvet() As String
        Dim WsMarginsLedosvet() As Variant
        ReDim CitiesArrayLedosvet(1 To 123)
        ProgruzkaLedosvet = CreateHeader(ProgruzkaLedosvet, "Б2Б ПМК Ледосвет")
        CitiesArrayLedosvet = CreateCities(CitiesArrayLedosvet, "Б2Б ПМК Ледосвет")
        WsMarginsLedosvet = WbProschet.Worksheets("Ледосвет").Range("A2:DW" & _
        WbProschet.Worksheets("Ледосвет").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
    End If
        
    'Перенос значений границ в массивы
    Dim WsLimits() As Variant
    WsLimits = WbProschet.Worksheets("Границы").Range("A2:E" & _
    WbProschet.Worksheets("Границы").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
   
    Dim NRow As Long, Nrow_PriceLimit As Long, NRowProschet As Long, Nrow_Margin As Long
    Dim Nrow_Progruzka As Long, Nrow_Progruzka_2 As Long, NCol_Progruzka As Long, WbProgruzka As String
    Dim NumberOfMargin As String, PriceLimit As Long
    Dim Komponent As String, Postavshik As String, Artikul As String
    Dim OldPrice As Long, RetailPrice As Long
    Dim MarginsArray() As Double
    Dim PriceArray() As Long
    Dim Transitive As Boolean
    
    
    '''ПМК'''
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If TotalRowsPMK <> 0 Then
        
        Nrow_Progruzka = 1
        Nrow_Progruzka_2 = 1
        ReDim PriceArray(1 To UBound(CitiesArrayPMK))
        For NRow = 1 To TotalRowsPMK
            
            'Внесение переменных
            Nrow_PriceLimit = CLng(PMKArray(2, NRow))
            Komponent = PMKArray(3, NRow)
            NRowProschet = PMKArray(4, NRow)
            
            If NRow <> 1 And Transitive = False Then
                If Postavshik <> Proschet(NRowProschet, 1) Then
                    Transitive = True
                Else
                    Transitive = False
                End If
            End If
            
            Postavshik = Proschet(NRowProschet, 1)
            
            Artikul = Proschet(NRowProschet, 2)
            Kod1C = Proschet(NRowProschet, 17)
            
            RetailPrice = SetRRC(Proschet(NRowProschet, 18))
            OldPrice = SetOldPrice(Proschet(NRowProschet, 19))
            
            'Выбор границы по опту
            If PMKArray(1, NRow) = 0 And PMKArray(2, NRow) = 0 Then
                PriceLimit = 0
            Else
                PriceLimit = FindPriceLimit(WsLimits, Komponent, Nrow_PriceLimit)
            End If
            
            'Выбор строки наценки по границе
            If PriceLimit <> 0 Then
                If OldPrice > PriceLimit Then
                    NumberOfMargin = CStr(PMKArray(1, NRow))
                Else
                    NumberOfMargin = CStr(PMKArray(0, NRow))
                End If
            Else
                NumberOfMargin = CStr(PMKArray(0, NRow))
            End If
            
            
            For i = 1 To UBound(WsMarginsPMK, 1)
                
                WsMarginsPMK(i, 2) = CStr(WsMarginsPMK(i, 2))
                WsMarginsPMK(i, 4) = CStr(WsMarginsPMK(i, 4))
                
                If (Komponent & CStr(NumberOfMargin)) = (WsMarginsPMK(i, 2) & WsMarginsPMK(i, 4)) Then
                    
                    Nrow_Margin = i
                    
                    ReDim MarginsArray(1 To UBound(CitiesArrayPMK))
                    For j = 1 To UBound(MarginsArray)
                    
                        If WsMarginsPMK(Nrow_Margin, j + 4) = "-" Then
                            WsMarginsPMK(Nrow_Margin, j + 4) = 0
                        End If
                        
                        MarginsArray(j) = WsMarginsPMK(Nrow_Margin, j + 4)
                    
                    Next
                    Exit For
                
                End If
                
            Next
            
            'Основной расчет цены
            ReDim PriceArray(1 To UBound(CitiesArrayPMK))
            For i = 1 To UBound(CitiesArrayPMK)
                
                If CityForMarginOutput <> vbNullString And CitiesArrayPMK(i) = CityForMarginOutput _
                Or CityForMarginOutput <> "" And CitiesArrayPMK(i) = CityForMarginOutput Then
                    MarginOutput = MarginOutput & vbCrLf & CStr(MarginsArray(i)) & " " & CStr(Komponent)
                End If
                
                If MarginsArray(i) = 0 Then
                    PriceArray(i) = 0
                Else
                    
                    If i < 88 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    ElseIf i >= 88 And i < 95 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_TENGE
                    ElseIf i = 95 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    ElseIf i = 96 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_BELRUB
                    ElseIf i = 97 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_SOM
                    ElseIf i = 98 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_DRAM
                    End If
                End If
                
                If PriceArray(i) > 7500 Then
                    PriceArray(i) = Round_Up(PriceArray(i), -1)
                ElseIf PriceArray(i) > 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                ElseIf PriceArray(i) < 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                End If
                
                If PriceArray(i) <> 0 Then
                    If PriceArray(i) <= RetailPrice Then
                        If i < 88 Then
                            PriceArray(i) = RetailPrice
                        ElseIf i >= 88 And i < 95 Then
                            PriceArray(i) = RetailPrice * KURS_TENGE
                        ElseIf i = 95 Then
                            PriceArray(i) = RetailPrice
                        ElseIf i = 96 Then
                            PriceArray(i) = RetailPrice * KURS_BELRUB
                        ElseIf i = 97 Then
                            PriceArray(i) = RetailPrice * KURS_SOM
                        ElseIf i = 98 Then
                            PriceArray(i) = RetailPrice * KURS_DRAM
                        End If
                        
                        If PriceArray(i) > 7500 Then
                            PriceArray(i) = Round_Up(PriceArray(i), -1)
                        ElseIf PriceArray(i) > 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        ElseIf PriceArray(i) < 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        End If
                        
                    End If
                End If
                
            Next
            
            'Разделение прогрузки ПМК на две, если есть компоненты 830/917
            If Komponent = "830" Or Komponent = "917" Then
                
                'Запись массива в прогрузку
                ProgruzkaPMK_2(Nrow_Progruzka_2 + 1, 1) = Komponent
                ProgruzkaPMK_2(Nrow_Progruzka_2 + 1, 2) = Artikul
                ProgruzkaPMK_2(Nrow_Progruzka_2 + 1, 3) = Kod1C
                ProgruzkaPMK_2(Nrow_Progruzka_2 + 1, 4) = OldPrice

                Nrow_Progruzka_2 = Nrow_Progruzka_2 + 1
                NCol_Progruzka = 5

                For i = 1 To (UBound(PriceArray) - 1)
                    ProgruzkaPMK_2(Nrow_Progruzka_2, NCol_Progruzka) = PriceArray(i)
                    
                    NCol_Progruzka = NCol_Progruzka + 1
                Next
                
                'Уникализация
                If OldPrice <> 0 Then
                    Call Unify(ProgruzkaPMK_2, Nrow_Progruzka_2, "Б2Б ПМК", "Row")
                End If
                
                'Уникализация в рамках одного кода 1С разных компонентов на одном городе:
                'Запускается при условии что текущая строка прогрузки не пуста и предыдущая строка прогрузки не равна шапке таблицы
                If Len(ProgruzkaPMK_2(Nrow_Progruzka_2, 3)) <> 0 And ProgruzkaPMK_2(Nrow_Progruzka_2 - 1, 3) <> "Код 1С" Then
                    'Запускается при условии что код 1С в текущей строке прогрузки не равен предыдущему коду 1С
                    If ProgruzkaPMK_2(Nrow_Progruzka_2, 3) <> ProgruzkaPMK_2(Nrow_Progruzka_2 - 1, 3) Or NRow = UBound(PMKArray, 2) Then
                        Call Unify(ProgruzkaPMK_2, Nrow_Progruzka_2, "Б2Б ПМК", "Column")
                    End If
                End If
                
            Else
            
                'Запись массива в прогрузку
                ProgruzkaPMK(Nrow_Progruzka + 1, 1) = Komponent
                ProgruzkaPMK(Nrow_Progruzka + 1, 2) = Artikul
                ProgruzkaPMK(Nrow_Progruzka + 1, 3) = Kod1C
                ProgruzkaPMK(Nrow_Progruzka + 1, 4) = OldPrice
                
                NCol_Progruzka = 5
                Nrow_Progruzka = Nrow_Progruzka + 1
                For i = 1 To UBound(PriceArray)
                    ProgruzkaPMK(Nrow_Progruzka, NCol_Progruzka) = PriceArray(i)
                    NCol_Progruzka = NCol_Progruzka + 1
                Next
                
                If UserForm1.CheckBox1 = False Then
                    'Уникализация в рамках одного кода 1С одного компонента на разных городах
                    If OldPrice <> 0 Then
                        Call Unify(ProgruzkaPMK, Nrow_Progruzka, "Б2Б ПМК", "Row")
                    End If
                    'Уникализация в рамках одного кода 1С разных компонентов на одном городе:
                    'Запускается при условии что текущая строка прогрузки не пуста и предыдущая строка прогрузки не равна шапке таблицы
                    If Len(ProgruzkaPMK(Nrow_Progruzka, 3)) <> 0 And ProgruzkaPMK(Nrow_Progruzka - 1, 3) <> "Код 1С" Then
                        'Запускается при условии что код 1С в текущей строке прогрузки не равен предыдущему коду 1С
                        If ProgruzkaPMK(Nrow_Progruzka, 3) <> ProgruzkaPMK(Nrow_Progruzka - 1, 3) Or NRow = UBound(PMKArray, 2) Then
                            Call Unify(ProgruzkaPMK, Nrow_Progruzka - 1, "Б2Б ПМК", "Column")
                        End If
                    End If
                End If
                    
            End If
            
        Next NRow

        If TotalRowsPMK <> 0 Then
                       
            Application.ScreenUpdating = False
            Workbooks.Add
            WbProgruzka = ActiveWorkbook.name
            Workbooks(WbProgruzka).ActiveSheet.Range("A1:CX" & UBound(ProgruzkaPMK, 1)).Value = ProgruzkaPMK
            
            If UserForm1.Autosave = True Then
                If Transitive = True Then
                    Postavshik = vbNullString
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК")
                Else
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК")
                End If
            End If
            
            If Nrow_Progruzka_2 > 2 Then
                ProgruzkaPMK = CreateHeader(ProgruzkaPMK, "Б2Б ПМК")
                Workbooks.Add
                WbProgruzka = ActiveWorkbook.name
                Workbooks(WbProgruzka).ActiveSheet.Range("A1:CW" & UBound(ProgruzkaPMK_2, 1)).Value = ProgruzkaPMK_2
                
                If UserForm1.Autosave = True Then
                    If Transitive = True Then
                        Postavshik = vbNullString
                        Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК 830/917")
                    Else
                        Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК 830/917")
                    End If
                End If
            End If
            Application.ScreenUpdating = True
        End If
        
        Erase PMKArray, PriceArray, MarginsArray, WsMarginsPMK, CitiesArrayPMK, ProgruzkaPMK
    End If
    
    
    '''ТОРУДА'''
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    If TotalRowsTORUDA <> 0 Then
        
        ReDim PriceArray(1 To UBound(CitiesArrayToruda))
        For NRow = 1 To TotalRowsTORUDA
            Nrow_PriceLimit = CLng(TorudaArray(2, NRow))
            Komponent = TorudaArray(3, NRow)
            NRowProschet = TorudaArray(4, NRow)
            
            If NRow <> 1 And Transitive = False Then
                If Postavshik <> Proschet(NRowProschet, 1) Then
                    Transitive = True
                Else
                    Transitive = False
                End If
            End If
            
            Postavshik = Proschet(NRowProschet, 1)
            
            Artikul = Proschet(NRowProschet, 2)
            Kod1C = Proschet(NRowProschet, 17)
            
            RetailPrice = SetRRC(Proschet(NRowProschet, 18))
            OldPrice = SetOldPrice(Proschet(NRowProschet, 19))
            
            If TorudaArray(1, NRow) = 0 And TorudaArray(2, NRow) = 0 Then
                PriceLimit = 0
            Else
                PriceLimit = FindPriceLimit(WsLimits, Komponent, Nrow_PriceLimit)
            End If
        
            If PriceLimit <> 0 Then
                If OldPrice > PriceLimit Then
                    NumberOfMargin = CStr(TorudaArray(1, NRow))
                Else
                    NumberOfMargin = CStr(TorudaArray(0, NRow))
                End If
            Else
                NumberOfMargin = CStr(TorudaArray(0, NRow))
            End If

            For i = 1 To UBound(WsMarginsToruda)
                WsMarginsToruda(i, 2) = CStr(WsMarginsToruda(i, 2))
                WsMarginsToruda(i, 4) = CStr(WsMarginsToruda(i, 4))
                If (Komponent & NumberOfMargin) = (WsMarginsToruda(i, 2) & WsMarginsToruda(i, 4)) Then
                
                    Nrow_Margin = i
                    ReDim MarginsArray(1 To UBound(CitiesArrayToruda))
                    For j = 1 To UBound(MarginsArray)
                        If WsMarginsToruda(Nrow_Margin, j + 4) = "-" Then
                            WsMarginsToruda(Nrow_Margin, j + 4) = 0
                        End If
                        
                        MarginsArray(j) = WsMarginsToruda(Nrow_Margin, j + 4)
                    Next
                    Exit For
                End If
            Next
            
            For i = 1 To UBound(CitiesArrayToruda)
                
                If CityForMarginOutput <> vbNullString And CitiesArrayToruda(i) = CityForMarginOutput _
                Or CityForMarginOutput <> "" And CitiesArrayToruda(i) = CityForMarginOutput Then
                    MarginOutput = MarginOutput & vbCrLf & CStr(MarginsArray(i)) & " " & CStr(Komponent)
                End If
                
                If MarginsArray(i) = 0 Then
                    PriceArray(i) = 0
                Else
                    If i < 124 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    Else
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_TENGE
                    End If
                End If
                
                If PriceArray(i) > 7500 Then
                    PriceArray(i) = Round_Up(PriceArray(i), -1)
                ElseIf PriceArray(i) > 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                ElseIf PriceArray(i) < 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                End If
                
                If PriceArray(i) <> 0 Then
                    If PriceArray(i) <= RetailPrice Then
                        If i < 124 Then
                            PriceArray(i) = OldPrice * MarginsArray(i)
                        Else
                            PriceArray(i) = OldPrice * MarginsArray(i) * KURS_TENGE
                        End If
                        
                        If PriceArray(i) > 7500 Then
                            PriceArray(i) = Round_Up(PriceArray(i), -1)
                        ElseIf PriceArray(i) > 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        ElseIf PriceArray(i) < 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        End If
                        
                    End If
                End If
            Next
            
            'Запись массива в прогрузку
            ProgruzkaToruda(NRow + 1, 1) = Artikul
            ProgruzkaToruda(NRow + 1, 2) = Komponent
            ProgruzkaToruda(NRow + 1, 3) = Kod1C
            ProgruzkaToruda(NRow + 1, 4) = OldPrice
            
            NCol_Progruzka = 5
            Nrow_Progruzka = NRow + 1
            
            For i = 1 To UBound(PriceArray)
                ProgruzkaToruda(Nrow_Progruzka, NCol_Progruzka) = PriceArray(i)
                NCol_Progruzka = NCol_Progruzka + 1
            Next
            
            
            'Уникализация:
            If UserForm1.CheckBox1 = False Then
                'Уникализация в рамках одного кода одного компонента на разных городах
                If OldPrice <> 0 Then
                    
                    Call Unify(ProgruzkaToruda, Nrow_Progruzka, "Б2Б Торуда", "Row")
                End If
                
                'Уникализация в рамках одного кода 1С разных компонентов на одном городе:
                'Запускается при условии что текущая строка прогрузки не пуста и предыдущая строка прогрузки не равна шапке таблицы
                If Len(ProgruzkaToruda(Nrow_Progruzka, 3)) <> 0 And ProgruzkaToruda(Nrow_Progruzka - 1, 3) <> "Код 1С" Then
                    'Запускается при условии что код 1С в текущей строке прогрузки не равен предыдущему коду 1С
                    If ProgruzkaToruda(Nrow_Progruzka, 3) <> ProgruzkaToruda(Nrow_Progruzka - 1, 3) Or NRow = UBound(TorudaArray, 2) Then
                        
                        Call Unify(ProgruzkaToruda, Nrow_Progruzka, "Б2Б Торуда", "Column")
                    End If
                End If
            End If
            
        
        Next NRow
        
        If TotalRowsTORUDA <> 0 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            WbProgruzka = ActiveWorkbook.name
            Workbooks(WbProgruzka).ActiveSheet.Range("A1:EQ" & UBound(ProgruzkaToruda, 1)).Value = ProgruzkaToruda
            
            If UserForm1.Autosave = True Then
                If Transitive = True Then
                    Postavshik = vbNullString
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б Торуда")
                Else
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б Торуда")
                End If
            End If
            
            Application.ScreenUpdating = True
        End If
        
        Erase TorudaArray, PriceArray, MarginsArray, WsMarginsToruda, CitiesArrayToruda, ProgruzkaToruda
    End If
        
    
    '''КАБИНЕТОФ'''
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    If TotalRowsKAB <> 0 Then
        
        ReDim PriceArray(1 To UBound(CitiesArrayKab))
        For NRow = 1 To TotalRowsKAB
        
            Nrow_PriceLimit = CLng(KABArray(2, NRow))
            Komponent = KABArray(3, NRow)
            NRowProschet = KABArray(4, NRow)
            
            Kod1C = Proschet(NRowProschet, 17)
            
            If NRow <> 1 And Transitive = False Then
                If Postavshik <> Proschet(NRowProschet, 1) Then
                    Transitive = True
                Else
                    Transitive = False
                End If
            End If
            
            Postavshik = Proschet(NRowProschet, 1)
            
            Artikul = Proschet(NRowProschet, 2)
            
            RetailPrice = SetRRC(Proschet(NRowProschet, 18))
            OldPrice = SetOldPrice(Proschet(NRowProschet, 19))
            
            'Поиск и запись границы в переменную
            If KABArray(1, NRow) = 0 And KABArray(2, NRow) = 0 Then
                PriceLimit = 0
            Else
                PriceLimit = FindPriceLimit(WsLimits, Komponent, Nrow_PriceLimit)
            End If
            
            If PriceLimit <> 0 Then
                If OldPrice > PriceLimit Then
                    NumberOfMargin = CStr(KABArray(1, NRow))
                Else
                    NumberOfMargin = CStr(KABArray(0, NRow))
                End If
            Else
                NumberOfMargin = CStr(KABArray(0, NRow))
            End If
            
            For i = 1 To UBound(WsMarginsKab)
                WsMarginsKab(i, 2) = CStr(WsMarginsKab(i, 2))
                WsMarginsKab(i, 4) = CStr(WsMarginsKab(i, 4))
                
                If (Komponent & NumberOfMargin) = (WsMarginsKab(i, 2) & WsMarginsKab(i, 4)) Then
                
                    Nrow_Margin = i
                    ReDim MarginsArray(1 To UBound(CitiesArrayKab))
                    For j = 1 To UBound(MarginsArray)
                        If WsMarginsKab(Nrow_Margin, j + 4) = "-" Then
                            WsMarginsKab(Nrow_Margin, j + 4) = 0
                        End If
                        
                        MarginsArray(j) = WsMarginsKab(Nrow_Margin, j + 4)
                    Next
                    Exit For
                End If
                
            Next
            
            For i = 1 To UBound(CitiesArrayKab)
                
                                
                If CityForMarginOutput <> vbNullString And CitiesArrayKab(i) = CityForMarginOutput _
                Or CityForMarginOutput <> "" And CitiesArrayKab(i) = CityForMarginOutput Then
                    MarginOutput = MarginOutput & vbCrLf & CStr(MarginsArray(i)) & " " & CStr(Komponent)
                End If
                
                If MarginsArray(i) = 0 Then
                    PriceArray(i) = 0
                Else
                    If i < 92 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    ElseIf i >= 92 And i < 101 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_TENGE
                    Else
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_BELRUB
                    End If
                End If
                
                If PriceArray(i) > 7500 Then
                    PriceArray(i) = Round_Up(PriceArray(i), -1)
                ElseIf PriceArray(i) > 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                ElseIf PriceArray(i) < 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                End If

                If PriceArray(i) <> 0 Then
                    If PriceArray(i) <= RetailPrice Then
                        If i < 92 Then
                            PriceArray(i) = RetailPrice
                        ElseIf i >= 92 And i < 101 Then
                            PriceArray(i) = RetailPrice * KURS_TENGE
                        Else
                            PriceArray(i) = RetailPrice * KURS_BELRUB
                        End If
                        
                        If PriceArray(i) > 7500 Then
                            PriceArray(i) = Round_Up(PriceArray(i), -1)
                        ElseIf PriceArray(i) > 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        ElseIf PriceArray(i) < 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        End If
                        
                    End If
                End If
            Next
            
            'Запись массива в прогрузку
            ProgruzkaKab(NRow + 1, 2) = Artikul
            ProgruzkaKab(NRow + 1, 3) = Kod1C
            ProgruzkaKab(NRow + 1, 4) = OldPrice
            
            NCol_Progruzka = 5
            Nrow_Progruzka = NRow + 1
            
            For i = 1 To UBound(PriceArray)
                ProgruzkaKab(Nrow_Progruzka, NCol_Progruzka) = PriceArray(i)
                NCol_Progruzka = NCol_Progruzka + 1
            Next
            
            'Уникализация
            If UserForm1.CheckBox1 = False Then
                
                'Уникализация в рамках одного кода одного компонента на разных городах
                If OldPrice <> 0 Then
                    
                    Call Unify(ProgruzkaKab, Nrow_Progruzka, "Б2Б Кабинетоф", "Row")
                End If
                
                'Уникализация в рамках одного кода 1С разных компонентов на одном городе:
                'Запускается при условии что текущая строка прогрузки не пуста и предыдущая строка прогрузки не равна шапке таблицы
                If Len(ProgruzkaKab(Nrow_Progruzka, 3)) <> 0 And ProgruzkaKab(Nrow_Progruzka - 1, 3) <> "Код 1С" Then
                    'Запускается при условии что код 1С в текущей строке прогрузки не равен предыдущему коду 1С
                    If ProgruzkaKab(Nrow_Progruzka, 3) <> ProgruzkaKab(Nrow_Progruzka - 1, 3) Or NRow = UBound(KABArray, 2) Then
                        
                        Call Unify(ProgruzkaKab, Nrow_Progruzka, "Б2Б Кабинетоф", "Column")
                    End If
                End If
            End If
            
            
        Next NRow
        
        If TotalRowsKAB <> 0 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            WbProgruzka = ActiveWorkbook.name
            Workbooks(WbProgruzka).ActiveSheet.Range("A1:DF" & UBound(ProgruzkaKab, 1)).Value = ProgruzkaKab
            
            If UserForm1.Autosave = True Then
                If Transitive = True Then
                    Postavshik = vbNullString
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б Кабинетоф")
                Else
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б Кабинетоф")
                End If
            End If
            Application.ScreenUpdating = True
        End If
        
        Erase KABArray, PriceArray, MarginsArray, WsMarginsKab, CitiesArrayKab, ProgruzkaKab
        
    End If
    
    '''ЛЕДОСВЕТ'''
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If TotalRowsLEDOSVET <> 0 Then
        ReDim PriceArray(1 To UBound(CitiesArrayLedosvet))
        For NRow = 1 To TotalRowsLEDOSVET
        
            Nrow_PriceLimit = CLng(LedosvetArray(2, NRow))
            Komponent = LedosvetArray(3, NRow)
            NRowProschet = LedosvetArray(4, NRow)
            
            Kod1C = Proschet(NRowProschet, 17)
            
            If NRow <> 1 And Transitive = False Then
                If Postavshik <> Proschet(NRowProschet, 1) Then
                    Transitive = True
                Else
                    Transitive = False
                End If
            End If
            
            Postavshik = Proschet(NRowProschet, 1)
            
            Artikul = Proschet(NRowProschet, 2)
            
            RetailPrice = SetRRC(Proschet(NRowProschet, 18))
            OldPrice = SetOldPrice(Proschet(NRowProschet, 19))
            
            If LedosvetArray(1, NRow) = 0 And LedosvetArray(2, NRow) = 0 Then
                PriceLimit = 0
            Else
                PriceLimit = FindPriceLimit(WsLimits, Komponent, Nrow_PriceLimit)
            End If
            
            If PriceLimit <> 0 Then
                If OldPrice > PriceLimit Then
                    NumberOfMargin = CStr(LedosvetArray(1, NRow))
                Else
                    NumberOfMargin = CStr(LedosvetArray(0, NRow))
                End If
            Else
                NumberOfMargin = CStr(LedosvetArray(0, NRow))
            End If
            
            For i = 1 To UBound(WsMarginsLedosvet)
                WsMarginsLedosvet(i, 2) = CStr(WsMarginsLedosvet(i, 2))
                WsMarginsLedosvet(i, 4) = CStr(WsMarginsLedosvet(i, 4))
                
                If (Komponent & NumberOfMargin) = (WsMarginsLedosvet(i, 2) & WsMarginsLedosvet(i, 4)) Then
                
                    Nrow_Margin = i
                    ReDim MarginsArray(1 To UBound(CitiesArrayLedosvet))
                    For j = 1 To UBound(MarginsArray)
                        If WsMarginsLedosvet(Nrow_Margin, j + 4) = "-" Then
                            WsMarginsLedosvet(Nrow_Margin, j + 4) = 0
                        End If
                        
                        MarginsArray(j) = WsMarginsLedosvet(Nrow_Margin, j + 4)
                    Next
                    Exit For
                End If
                
            Next
            
            For i = 1 To UBound(CitiesArrayLedosvet)
                
                If CityForMarginOutput <> vbNullString And CitiesArrayKab(i) = CityForMarginOutput _
                Or CityForMarginOutput <> "" And CitiesArrayKab(i) = CityForMarginOutput Then
                    MarginOutput = MarginOutput & vbCrLf & CStr(MarginsArray(i)) & " " & CStr(Komponent)
                End If
                
                If MarginsArray(i) = 0 Then
                    PriceArray(i) = 0
                Else
                    If i < 104 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    ElseIf i >= 104 And i < 111 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_TENGE
                    ElseIf i = 111 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_BELRUB
                    ElseIf i > 111 And i < 122 Then
                        PriceArray(i) = OldPrice * MarginsArray(i)
                    ElseIf i = 122 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_SOM
                    ElseIf i = 123 Then
                        PriceArray(i) = OldPrice * MarginsArray(i) * KURS_DRAM
                    End If

                End If
                
                If PriceArray(i) > 7500 Then
                    PriceArray(i) = Round_Up(PriceArray(i), -1)
                ElseIf PriceArray(i) > 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                ElseIf PriceArray(i) < 1000 Then
                    PriceArray(i) = Round_Up(PriceArray(i), 0)
                End If

                If PriceArray(i) <> 0 Then
                    If PriceArray(i) <= RetailPrice Then
                        If i < 104 Then
                            PriceArray(i) = RetailPrice
                        ElseIf i >= 104 And i < 111 Then
                            PriceArray(i) = RetailPrice * KURS_TENGE
                        ElseIf i = 111 Then
                            PriceArray(i) = RetailPrice * KURS_BELRUB
                        ElseIf i > 111 And i < 122 Then
                            PriceArray(i) = RetailPrice
                        ElseIf i = 122 Then
                            PriceArray(i) = RetailPrice * KURS_SOM
                        ElseIf i = 123 Then
                            PriceArray(i) = RetailPrice * KURS_DRAM
                        End If
                        
                        If PriceArray(i) > 7500 Then
                            PriceArray(i) = Round_Up(PriceArray(i), -1)
                        ElseIf PriceArray(i) > 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        ElseIf PriceArray(i) < 1000 Then
                            PriceArray(i) = Round_Up(PriceArray(i), 0)
                        End If
                        
                    End If
                End If
            Next
                
            'Запись массива в прогрузку
            ProgruzkaLedosvet(NRow + 1, 1) = Komponent
            ProgruzkaLedosvet(NRow + 1, 2) = Artikul
            ProgruzkaLedosvet(NRow + 1, 3) = Kod1C
            ProgruzkaLedosvet(NRow + 1, 4) = OldPrice
            
            NCol_Progruzka = 5
            Nrow_Progruzka = NRow + 1
            
            For i = 1 To UBound(PriceArray)
                ProgruzkaLedosvet(Nrow_Progruzka, NCol_Progruzka) = PriceArray(i)
                NCol_Progruzka = NCol_Progruzka + 1
            Next
            
            If UserForm1.CheckBox1 = False Then
                
                'Уникализация в рамках одного кода одного компонента на разных городах
                If OldPrice <> 0 Then
                    
                    Call Unify(ProgruzkaLedosvet, Nrow_Progruzka, "Б2Б ПМК Ледосвет", "Row")
                End If
                
                'Уникализация в рамках одного кода 1С разных компонентов на одном городе:
                'Запускается при условии что текущая строка прогрузки не пуста и предыдущая строка прогрузки не равна шапке таблицы
                If Len(ProgruzkaLedosvet(Nrow_Progruzka, 3)) <> 0 And ProgruzkaLedosvet(Nrow_Progruzka - 1, 3) <> "Код 1С" Then
                    'Запускается при условии что код 1С в текущей строке прогрузки не равен предыдущему коду 1С
                    If ProgruzkaLedosvet(Nrow_Progruzka, 3) <> ProgruzkaLedosvet(Nrow_Progruzka - 1, 3) Or NRow = UBound(LedosvetArray, 2) Then
                        
                        Call Unify(ProgruzkaLedosvet, Nrow_Progruzka, "Б2Б ПМК Ледосвет", "Column")
                    End If
                End If
            End If
                
        Next NRow
            
        If TotalRowsLEDOSVET <> 0 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            WbProgruzka = ActiveWorkbook.name
            Workbooks(WbProgruzka).ActiveSheet.Range("A1:DW" & UBound(ProgruzkaLedosvet, 1)).Value = ProgruzkaLedosvet
            
            If UserForm1.Autosave = True Then
                If Transitive = True Then
                    Postavshik = vbNullString
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК Ледосвет")
                Else
                    
                    Call Autosave(Postavshik, WbProgruzka, "Б2Б ПМК Ледосвет")
                End If
            End If
            Application.ScreenUpdating = True
        End If
        
        Erase LedosvetArray, PriceArray, MarginsArray, WsMarginsLedosvet, CitiesArrayLedosvet, ProgruzkaLedosvet
    End If
    
    If CityForMarginOutput <> vbNullString Or CityForMarginOutput <> "" Then
        MsgBox MarginOutput
    End If
    
    Unload UserForm1

End Sub


