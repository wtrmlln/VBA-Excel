Attribute VB_Name = "B2C_СОТ_Торуда"
Option Private Module
Option Explicit

'Прогрузка SOT и Toruda
Public Sub Прогрузка_СОТ_Торуда()

    Dim OptPrice As Long
    Dim PriceStep As Long
    
    Dim FullDostavkaPrice As Double                                 'Рассчитывается в fDostavka
    Dim TerminalTerminalPrice As Double                                       'Межтерминалка: либо мин. знач. тарифа ДПД, либо PlatniiVes * Тариф ДПД
    Dim PlatniiVes As Double                                 'Равен Объем * 250, или Физический вес из просчета, или Объемный вес из просчета
    
    Dim Tarif As Double
    Dim ProgruzkaSOTArray() As Variant
    Dim TeplozharPromoArray() As Long
    Dim KompkreslaPromoArray() As Long
    Dim KiskaPromoArray() As Long
    Dim PMKPromoArray() As Variant
    Dim TotalRowsTeplozharPromo As Long                      '620П
    Dim TotalRowsKompkreslaPromo As Long                     'НП
    Dim TotalRowsKiskaPromo As Long                    '1196П
    Dim TotalRowsPMK As Long                        '796
    Dim TotalRowsSOTPromo As Long                 '?
    Dim TotalRowsPMKPromo As Long
    Dim TriggerTeplo As Long
    Dim i1 As Long
    Dim b1 As Long
    Dim g1 As Long
    Dim pmk1 As Long
    Dim NRowProschet As Long
    Dim r3 As Long
    Dim i As Long
    Dim j As Long
    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim f As Long
    Dim f1 As Long
    Dim MinTerminalTerminalPrice As Double
    
    Dim Transitive As Boolean
    Transitive = True
    
    'mvvodin = Коды 1С из юзерформы разделенные Enter
    Dim mvvodin() As String
    mvvodin = Split(UserForm1.TextBox1, vbCrLf)
    If (UBound(mvvodin)) = -1 Then
        GoTo ended
    End If

    'Изменяет границу mvvod на количество Кодов 1С из юзерформы + Пустая строка
    Dim mvvod() As String
    ReDim mvvod(UBound(mvvodin))

    'Если в mvvodin есть пустое значение, тогда mvvod его пропускает
    'Если пустого значения нет, тогда код 1С вносится в mvvod
    For i = 0 To UBound(mvvodin)
        mvvodin(i) = Application.Trim(mvvodin(i))
        If Len(mvvodin(i)) = 0 Or mvvodin(i) = vbNullString Then
            ReDim Preserve mvvod(UBound(mvvod) - 1)
        Else
            mvvod(i - (UBound(mvvodin) - UBound(mvvod))) = mvvodin(i)
        End If
    Next

    'Проверяет коды 1С из mvvod на дубликаты или отсутствие в просчете, если дубликат есть - макрос закрывается
    If DublicateFind(mvvod) = True Then
        GoTo ended
    End If

    'Проверяет коды 1С из mvvod на акцию
    If SaleFind(mvvod) = True Then
        GoTo ended
    End If
    
    Dim NcolumnProgruzka As Long
    Dim NRowProgruzka As Long
    Dim ErrorCount As Long
    NcolumnProgruzka = 5                                  'Счетчик столбцов в прогрузке
    NRowProgruzka = 2                                     'номер строки для заполнения прогрузки, начинается с 2
    ErrorCount = 0
    
    Dim smkat As Long
    
    Dim NRowsArray() As Long
    ReDim NRowsArray(UBound(mvvod))
    
    Dim str As String
    Dim r1 As Long
    Dim a() As String
    Dim TotalRows As Long
    
    'Обычная прогрузка
    If UserForm1.OptionButton1 = True Or UserForm1.OptionButton4 = True Then
        '(1).   От 0 до кол-ва кодов в юзерформе
        For i = 0 To UBound(mvvod)
    
            'Ищет код 1С из юзерформы в просчете, r1 - счетчик строки
            'str = Код в юзерформе(Нижняя граница + d=0)
            str = Trim(mvvod(LBound(mvvod) + i))
            r1 = Proschet_Dict.Item(str)
            'NRowsArray(i) = Номер строки в просчете кода 1с из юзерформы по порядку
            NRowsArray(i) = r1
        
            'a = Массив из наценок и компонентов в просчете по коду 1С из юзерформы без точек с запятой
            a = Split(ProschetArray(NRowsArray(i), 7), ";")
        
            'mvvodin = Массив из наценок и компонентов в просчете по коду 1C
            'из юзерформы без точек с запятой и пустых значений
            'Изменяет границу mvvodin на количество наценок и компонентов
            'Если в "a" есть пустое значение,
            'Тогда в mvvodin оно не заносится и граница сокращается на 1
            'Если в "a" пустого значения нет,
            'Тогда оно вносится в mvvodin
            ReDim mvvodin(UBound(a))
            For j = 0 To UBound(a)
                If Len(a(j)) = 0 Then
                    ErrorCount = ErrorCount + 1
                    ReDim Preserve mvvodin(UBound(mvvodin) - 1)
                Else
                    mvvodin(j - (UBound(a) - UBound(mvvodin))) = a(j)
                End If
            Next
            
            If ErrorCount > 0 Then
                ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = Join(mvvodin, ";")
                ProschetArray(NRowsArray(i), 7) = Join(mvvodin, ";")
            End If
        
            'Если количество наценок и компонентов + 1 не делится без остатка на 5, тогда
            'Выводится ошибка и код 1С
            If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
                MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(i), 17)
                GoTo ended
            End If
        
            'Через mvvodin считается общее количество компонентов, складывается через цикл
            'и получается кол-во строк для прогрузки
            TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows
    
optimization1:
            'Если счетчик цикла < Кол-во кодов в юзерформы без пустой строки
            'Если код в просчете = Код в юзерформе, тогда
            'Повторяются все те же действия из цикла (1).
            If i < UBound(mvvod) Then
                If Trim(StrConv(ProschetArray(r1 + 1, 17), vbUpperCase)) = mvvod(LBound(mvvod) + i + 1) Then
                
                    r1 = r1 + 1
                    i = i + 1
                    NRowsArray(i) = r1
    
                    a = Split(ProschetArray(NRowsArray(i), 7), ";")
    
                    ReDim mvvodin(UBound(a))
                    For j = 0 To UBound(a)
                        If Len(a(j)) = 0 Then
                            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
                            ErrorCount = ErrorCount + 1
                        Else
                            mvvodin(j - (UBound(a) - UBound(mvvodin))) = a(j)
                        End If
                    Next
            
                    If ErrorCount > 0 Then
                        ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = Join(mvvodin, ";")
                        ProschetArray(NRowsArray(i), 7) = Join(mvvodin, ";")
                    End If
                    
                    If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
                        MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(i), 17)
                        GoTo ended
                    End If
    
                    TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows
    
                    GoTo optimization1
                End If
            End If
        Next
    'Снять с промо
    ElseIf UserForm1.OptionButton3 = True Then
        For i = 0 To UBound(mvvod)
    
            str = Trim(mvvod(LBound(mvvod) + i))
            r1 = Proschet_Dict.Item(str)
            NRowsArray(i) = r1
    
            a = Split(ProschetArray(NRowsArray(i), 7), ";")
    
            ReDim mvvodin(UBound(a))
            For j = 0 To UBound(a)
                If Len(a(j)) = 0 Then
                    ReDim Preserve mvvodin(UBound(mvvodin) - 1)
                Else
                    mvvodin(j - (UBound(a) - UBound(mvvodin))) = a(j)
                End If
            Next
    
            If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
                MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(i), 17)
                GoTo ended
            End If
    
            TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows
    
            Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(NRowsArray(i), 7).Value = Replace(Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(r1, 7).Value, Trim(kat) & "П", kat)
            ProschetArray(NRowsArray(i), 7) = Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(NRowsArray(i), 7).Value
            
            If InStr(Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(NRowsArray(i), 7).Value, "П") = 0 Then
                Windows(ProschetWorkbookName).Activate
                Range(Cells(NRowsArray(i), 1), Cells(NRowsArray(i), 19)).Select
                Application.ScreenUpdating = False
                With Selection.Interior
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                Application.ScreenUpdating = True
            End If
        Next
    End If

    Dim d As Long
    Dim TotalRowsKompkresla As Long, TotalRowsTeplozhar As Long, TotalRowsPromo As Long, TotalRowsKiska As Long
    d = 0
    TotalRowsKompkresla = 0
    TotalRowsPMK = 0
    TotalRowsTeplozhar = 0
    TotalRowsPromo = 0
    TotalRowsKiska = 0

    'Общая прогрузка для СОТ и Торуда
    Dim ProgruzkaTorudaArray() As Variant
    ReDim ProgruzkaMainArray(1 To TotalRows + 1, 1 To 168)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С СОТ/Торуда")
    
    Dim CitiesArray() As String
    ReDim CitiesArray(20 To 170)
    CitiesArray = CreateCities(CitiesArray, "Б2С СОТ/Торуда")

    Dim TeplozharArray() As Long
    ReDim TeplozharArray(3, UBound(mvvod))
    
    Dim KompkreslaArray() As Long
    ReDim KompkreslaArray(3, UBound(mvvod))
    
    Dim TotalPromoArray() As Variant
    ReDim TotalPromoArray(4, 100, UBound(mvvod))
    
    Dim KiskaArray() As Long
    ReDim KiskaArray(3, UBound(mvvod))
    
    Dim mteplo() As Variant
    ReDim mteplo(2, 0)
    
    Dim PMKArray() As Long
    ReDim PMKArray(3, UBound(mvvod))

    'При GoTo Kod1C_Change, переходит к следующему коду 1С из юзерформы
Kod1C_Change:
 
    NRowProschet = NRowsArray(d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
        
    'a = Массив из наценок и компонентов в просчете по коду 1С из юзерформы по порядку без точек с запятой
    a = Split(ProschetArray(NRowProschet, 7), ";")
    
    ReDim mvvodin(UBound(a))
    'Если в "a" есть пустое значение, тогда
    'В mvvodin оно не вносится, и граница mvvodin сокращается на 1
    'В ином случае
    'В mvvodin заносятся наценки и компоненты из просчета по коду 1С из юзерформы
    For i = 0 To UBound(a)
        If Len(a(i)) = 0 Then
            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
        Else
            mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
        End If
    Next
    
    'Замена ProschetArray(NRowProschet, "") на переменные:
    Dim Artikul As String, GorodOtpravki As String, TypeZaborProschet As String, Kod1CProschet As String, Komplektuyeshee As String
    Dim Quantity As Long, RetailPrice As Long, OptPriceProschet As Long, RetailPriceInfo As Long
    Dim Volume As Double, VolumetricWeight As Double, RazgruzkaPrice As Double, ZaborPriceProschet As Double, JU As Double, Weight As Double
    Artikul = ProschetArray(NRowProschet, 2)
    Quantity = ProschetArray(NRowProschet, 5)
    GorodOtpravki = ProschetArray(NRowProschet, 6)
    TypeZaborProschet = ProschetArray(NRowProschet, 8)
    Volume = ProschetArray(NRowProschet, 9)
    Weight = ProschetArray(NRowProschet, 10)
    VolumetricWeight = ProschetArray(NRowProschet, 11)
    RazgruzkaPrice = ProschetArray(NRowProschet, 12)
    ZaborPriceProschet = ProschetArray(NRowProschet, 13)
    JU = ProschetArray(NRowProschet, 14)
    Komplektuyeshee = ProschetArray(NRowProschet, 16)
    Kod1CProschet = ProschetArray(NRowProschet, 17)
    RetailPrice = ProschetArray(NRowProschet, 18)
    OptPriceProschet = ProschetArray(NRowProschet, 19)
    RetailPriceInfo = ProschetArray(NRowProschet, 20)

    'Обрешетка:
    'Если (Объем * 1440) < 2280, тогда ObreshetkaPrice = 950
    'В ином случае ObreshetkaPrice = Объем * 1440
    Dim ObreshetkaPrice As Double
    '2280 -> 2280
    If Volume * 1440 < 950 Then
        '2280 -> 2280
        ObreshetkaPrice = 950
    Else
        '1440 -> 1440
        ObreshetkaPrice = Volume * 1440
    End If

    'Если Вес > 100 и Кол-во > 1 тогда,
    'tps = Вес * 3 / Количество
    'В ином случае, если вес > 100 тогда,
    'tps = Вес * 3
    'В ином случае
    'tps = 0
    Dim tps As Double
    tps = fTPS(Weight, Quantity)

    'Если город льготный, и забор не равен нулю, тогда
    'Забор = 200
    'LgotniiGorod = True
    'В ином случае
    'Забор = Из столбца "Забор" просчета
    'LgotniiGorod = False
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    If DictLgotniiGorod.exists(GorodOtpravki) = True And ZaborPriceProschet <> 0 Then
        ZaborPrice = 200
        LgotniiGorod = True
    Else
        ZaborPrice = ZaborPriceProschet
        LgotniiGorod = False
    End If

    'Если d<>0 и Поставщик из просчета = Postavshik
    'Переходит к Komponent_Change
    'В ином случае если d <> 0, тогда
    'transitive = False?
    
    Dim Postavshik As String
    If d <> 0 And ProschetArray(NRowProschet, 1) = Postavshik Then
        GoTo Komponent_Change
    ElseIf d <> 0 Then
        Transitive = False
    End If

    'Postavshik = Поставщик из первого столбца просчета
    Postavshik = ProschetArray(NRowProschet, 1)

    'Переключается на другой компонент кода 1С
    'str = Компонент кода 1С по наценке
    'Cycle = Komponent_Change
Komponent_Change:

    str = a(4 + smkat)

    Dim TorudaKomponent As Boolean
    TorudaKomponent = False

    'Если в компоненте есть t,
    'К триггеру плюсуется 1
    Dim TriggerToruda As Long
    Dim TriggerTorudaPromo As Long
    If InStr(str, "t") <> 0 Then
        TriggerToruda = TriggerToruda + 1
        TorudaKomponent = True
        
        'Если в компоненте есть tП,
        'К триггеру плюсуется 1
        If InStr(str, "П") Then
            TriggerTorudaPromo = TriggerTorudaPromo + 1
        End If
    End If

    'Определяет - является ли код подходящий под условия для исключения торуда (опт <= 1000)
    Dim NeiskluchenieToruda As Boolean
    NeiskluchenieToruda = False
    If TorudaKomponent = True Then
        If OptPriceProschet * Quantity <= 1000 Then
            If DictIsklucheniyaToruda.exists(Kod1CProschet) = False Then
                NeiskluchenieToruda = True
            End If
        End If
    End If

    'TypeZabor = Тип забора для инфостолбца по FillInfocolumnTypeZabor
    Dim TypeZabor As String
    TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended

    'KomponentNumber = Числовое значение компонента кода 1С
    Dim KomponentNumber As String
    KomponentNumber = CStr(Val(a(4 + smkat)))

    'Если числовое значение компонента <> компоненту, и в компоненте нет "Н","П","t", тогда
    'Выводится окно с кодом 1С и компонентом
    If KomponentNumber <> str And str <> "Н" And InStr(str, "П") = 0 And TorudaKomponent = False Then
        MsgBox "Некорректная наценка/категория " & Kod1CProschet & " " & str
        GoTo ended
    End If

    'Для Teploman:
    'Если компонент кода 1С подходит под одно из этих условий, тогда
    'Если в mteplo добавляются элементы не впервые, тогда
    'Если предыдущий код 1С в mteplo <> Текущему коду 1С по d из юзерформы, тогда
    'В mteplo(0,1,2; Верхняя граница) вносятся:
    'В 0 - Код 1С из просчета
    'В 1 - ОПТ из просчета
    'В 2 - № строки кода 1С по просчету
    'В ином случае всё аналогично, с добавлением триггера + 1
    
    If KomponentNumber = 562 Or KomponentNumber = 618 Or KomponentNumber = 620 _
    Or KomponentNumber = 680 Or KomponentNumber = 681 Or KomponentNumber = 682 _
    Or KomponentNumber = 683 Or KomponentNumber = 708 Or KomponentNumber = 768 _
    Or KomponentNumber = 775 Or KomponentNumber = 777 Or KomponentNumber = 786 _
    Or KomponentNumber = 793 Or KomponentNumber = 821 Or KomponentNumber = 834 _
    Or KomponentNumber = 849 Or KomponentNumber = 997 Then
        If UBound(mteplo, 2) <> 0 Then
            If mteplo(0, UBound(mteplo, 2) - 1) <> Kod1CProschet Then
                mteplo(0, UBound(mteplo, 2)) = Trim(Kod1CProschet)
                mteplo(1, UBound(mteplo, 2)) = OptPriceProschet
                mteplo(2, UBound(mteplo, 2)) = NRowProschet
                ReDim Preserve mteplo(2, UBound(mteplo, 2) + 1)
            End If
        Else
            mteplo(0, UBound(mteplo, 2)) = Trim(Kod1CProschet)
            mteplo(1, UBound(mteplo, 2)) = OptPriceProschet
            mteplo(2, UBound(mteplo, 2)) = NRowProschet
            TriggerTeplo = 1
            ReDim Preserve mteplo(2, UBound(mteplo, 2) + 1)
        End If
    End If

    'Если компонент равен 796, 620, 1116, Н тогда:
    'В PMKArray(0,1,2,3; d)
    'В KompkreslaArray(0,1,2,3; d)
    'В TeplozharArray(0,1,2,3; d)
    'В KiskaArray(0,1,2,3; d)
    'В PromoArray(0,1,2,3,4; d)
    'В 0 вносится порядковый № компонента (отсчет с 0, разделитель ";")
    'В 1 вносится кол-во наценок и компонентов в наценке из просчета
    'В 2 вносится № строки в просчете у кода 1С из юзерформы
    'В 3 вносится значение для смены компонента у одного кода 1С
    'В 4 (только для PromoArray) вносится сам компонент с "П"
    'И плюсуются триггеры
    Dim TriggerPMK As Long, TriggerPromo As Long, TriggerKompkresla As Long, TriggerTeplozhar As Long, TriggerKiska As Long
    If str = "796" Or str = "1034" Then
        ReDim Preserve PMKArray(3, TotalRowsPMK)
        PMKArray(0, TotalRowsPMK) = NRowProgruzka - smkat / 5 + 2
        PMKArray(1, TotalRowsPMK) = UBound(a)
        PMKArray(2, TotalRowsPMK) = NRowProschet
        PMKArray(3, TotalRowsPMK) = smkat
        TriggerPMK = TriggerPMK + 1
        TotalRowsPMK = TotalRowsPMK + 1
        GoTo Change_Row_Progruzka
    End If
   
    If InStr(str, "П") <> 0 Then
        TotalPromoArray(0, TriggerPromo, d) = NRowProgruzka - smkat / 5 + 2
        TotalPromoArray(1, TriggerPromo, d) = UBound(mvvodin) + 1
        TotalPromoArray(2, TriggerPromo, d) = NRowProschet
        TotalPromoArray(3, TriggerPromo, d) = smkat
        TotalPromoArray(4, TriggerPromo, d) = a(4 + smkat)
        TriggerPromo = TriggerPromo + 1
        TotalRowsPromo = TotalRowsPromo + 1
        GoTo Change_Row_Progruzka
    End If

    If str = "Н" Then
        KompkreslaArray(0, TotalRowsKompkresla) = NRowProgruzka - smkat / 5 + 2
        KompkreslaArray(1, TotalRowsKompkresla) = UBound(a)
        KompkreslaArray(2, TotalRowsKompkresla) = NRowProschet
        KompkreslaArray(3, TotalRowsKompkresla) = smkat
        TriggerKompkresla = 1
        TotalRowsKompkresla = 1 + TotalRowsKompkresla
        GoTo Change_Row_Progruzka
    End If

    If str = "620" Then
        TeplozharArray(0, TotalRowsTeplozhar) = NRowProgruzka - smkat / 5 + 2
        TeplozharArray(1, TotalRowsTeplozhar) = UBound(a)
        TeplozharArray(2, TotalRowsTeplozhar) = NRowProschet
        TeplozharArray(3, TotalRowsTeplozhar) = smkat
        TriggerTeplozhar = 1
        TotalRowsTeplozhar = 1 + TotalRowsTeplozhar
        GoTo Change_Row_Progruzka
    End If

    If str = "1116" Then
        KiskaArray(0, TotalRowsKiska) = NRowProgruzka - smkat / 5 + 2
        KiskaArray(1, TotalRowsKiska) = UBound(a)
        KiskaArray(2, TotalRowsKiska) = NRowProschet
        KiskaArray(3, TotalRowsKiska) = smkat
        TriggerKiska = 1
        TotalRowsKiska = 1 + TotalRowsKiska
        GoTo Change_Row_Progruzka
    End If
    
    'fmargin:
    'Если промо,
    'ОПТ < 10 тыс. - первая наценка
    'ОПТ < 20 тыс. - вторая наценка
    'ОПТ >= 200 тыс. - третья наценка
    'Если кол-во штук > 1:
    '(ОПТ * Кол-во) < 10 тыс. - первая наценка
    '(ОПТ * Кол-во) < 20 тыс. - вторая наценка
    '(ОПТ * Кол-во) >= 20 тыс. - третья наценка
    'Если не промо,
    'ОПТ < 15 тыс. - первая наценка
    'ОПТ < 50 тыс. - вторая наценка
    'ОПТ < 200 тыс. - третья наценка
    'ОПТ >= 200 тыс. - четвертая наценка
    'Если кол-во > 1: аналогично, умножая на кол-во штук
    Dim Margin As Double
    Margin = fMargin(OptPriceProschet, Quantity, smkat, 0, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = str
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 157) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 158) = FillInfocolumnMargin(Komplektuyeshee, Margin, NeiskluchenieToruda, False)
    ProgruzkaMainArray(NRowProgruzka, 159) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 160) = Volume
    ProgruzkaMainArray(NRowProgruzka, 161) = Weight
    ProgruzkaMainArray(NRowProgruzka, 162) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 163) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 164) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 165) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 166) = Postavshik
    ProgruzkaMainArray(NRowProgruzka, 167) = Artikul

    'Если в компоненте есть t, и если ОПТ <= 1000
    'И если такого кода 1С нет в листе "Исключения торуда(<=1000), тогда
    'В прогрузке в столбце JU указывается "нет"
    If NeiskluchenieToruda = True Then
        ProgruzkaMainArray(NRowProgruzka, 163) = "нет"
    End If
        
    'Если в ОПТ просчета ошибка, тогда
    'Переходит к Change_Row_Progruzka
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    'Если в ОПТ просчета не числовое значение или 0, тогда
    'Переходит к Change_Row_Progruzka
    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    'Выбор конечного количества городов:
    Dim TotalCitiesCount As Long
    If TorudaKomponent = True Then
        TotalCitiesCount = 170
    Else
        TotalCitiesCount = 142
    End If
        
    TerminalTerminalPrice = 0
    FullDostavkaPrice = 0

    'С этого цикла происходит запись цен в массив прогрузки в одной строке
    Dim TotalPrice As Long
    Dim Price As Long
    Dim NCity As Long
    For NCity = 20 To TotalCitiesCount
    
        'Расчет цены с наценкой на комплектующие и Исключения(торуда <=1000):

        'Если в столбце "Формула" стоит "К"/"K1"/"K2",
        'Или Код 1С с оптом <= 1000 отсутствует на листе "Исключения(Торуда <= 1000)
        'Тогда цена умножается на 1.5/1.7/2 и округляется:

        If Komplektuyeshee = "Т" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = OptPriceProschet * Margin
                ElseIf NCity >= 123 And NCity < 135 Then
                    TotalPrice = OptPriceProschet * Margin * KURS_TENGE
                ElseIf NCity >= 135 And NCity < 163 Then
                    TotalPrice = OptPriceProschet * Margin
                Else
                    TotalPrice = OptPriceProschet * Margin * KURS_TENGE
                End If
            ElseIf NCity < 123 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            ElseIf NCity = 142 Then
                TotalPrice = OptPriceProschet * Margin * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К2" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = OptPriceProschet * 2
                ElseIf NCity >= 123 And NCity < 135 Then
                    TotalPrice = OptPriceProschet * 2 * KURS_TENGE
                ElseIf NCity >= 135 And NCity < 163 Then
                    TotalPrice = OptPriceProschet * 2
                Else
                    TotalPrice = OptPriceProschet * 2 * KURS_TENGE
                End If
            ElseIf NCity < 123 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 2 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
    
        ElseIf NeiskluchenieToruda = True Then

            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 1.5
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * (1.5 + SNG_MARGIN) * KURS_TENGE
            ElseIf NCity < 143 Then
                TotalPrice = OptPriceProschet * (1.5 + SNG_MARGIN)
            ElseIf NCity < 163 Then
                TotalPrice = OptPriceProschet * 1.5
            Else
                TotalPrice = OptPriceProschet * (1.5 + SNG_MARGIN) * KURS_TENGE
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
    
        ElseIf Komplektuyeshee = "К" Then

            If TorudaKomponent = True Then

                If NCity < 123 Then
                    TotalPrice = OptPriceProschet * 2
                ElseIf NCity < 135 Then
                    TotalPrice = OptPriceProschet * 2 * KURS_TENGE
                ElseIf NCity >= 135 And NCity < 163 Then
                    TotalPrice = OptPriceProschet * 2
                Else
                    TotalPrice = OptPriceProschet * 2 * KURS_TENGE
                End If

            ElseIf NCity < 123 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 2 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis

        ElseIf Komplektuyeshee = "К1" Then

            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = OptPriceProschet * 1.7
                ElseIf NCity >= 123 And NCity < 135 Then
                    TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
                ElseIf NCity >= 135 And NCity < 163 Then
                    TotalPrice = OptPriceProschet * 1.7
                Else
                    TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
                End If

            ElseIf NCity < 123 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 1.7 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        End If
          
        'Умножение опта на наценку:
        'NCity = № города
        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity >= 123 And NCity < 143 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 163 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        Else
            Price = OptPriceProschet * Margin
        End If
    
        'Расчет тарифа ДПД и минималки:

        'Описание fTarif:
        'Если прогрузка не для компкресел или DDI тогда
        'Если объемный вес = 0, тогда
        'fTarif = из листа "Тариф(omaxT)" по ГородОтправления&ГородПолучения по минимальному значению
        'В ином случае
        'Если объемный вес <= 100,
        'fTarif = до 100
        'Если объемный вес <= 200,
        'fTarif = до 200
        'Если объемный вес <= 800,
        'fTarif = до 800
        'Если объемный вес <= 1500,
        'fTarif = до 1500
        'Если объемный вес <= 3000,
        'fTarif = до 3000
        'Если объемный вес > 3000,
        'fTarif = более 3000
        'Если прогрузка для компкресел или DDI, тогда
        'Условия аналогичны, fTarif берется из листа omaxД
        
        'Если прогрузка не для компкресел или DDI,
        'Тариф рассчитывается по объемному весу из просчета по листу omaxT
        'Если прогрузка для компкресел или DDI,
        'Тариф рассчитывается по объемному весу из просчета по листу omaxД
        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        'MinTerminalTerminalPrice = fTarif без платного веса, т.е. всегда по минимальному значению
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))
  
        'Расчет платного веса:
        If NCity >= 123 And NCity < 143 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        ElseIf NCity >= 163 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        Else
            PlatniiVes = VolumetricWeight
        End If
        
        'Межтерминалка:
        'Если PlatniiVes * Tarif < минимального значения тарифа ДПД, тогда
        'TerminalTerminalPrice = Минимальное знач. тарифа ДПД
        'В ином случае
        'TerminalTerminalPrice = PlatniiVes * Tarif
        If PlatniiVes * Tarif < MinTerminalTerminalPrice Then
            TerminalTerminalPrice = MinTerminalTerminalPrice
        Else
            TerminalTerminalPrice = PlatniiVes * Tarif
        End If
    
        '+500 рублей к межтерминалке для Калининграда. 20.10.2021
        If NCity = 44 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If

        'Расчет полных транспортных:
        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps

        TotalPrice = Price + FullDostavkaPrice
    
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
    
        'Округление TotalPrice:
        'ju = Числовое значение в столбце ЖУ (0 или 0,01)
    
        If TorudaKomponent = True Then
            If NCity < 123 Or NCity >= 135 And NCity < 163 Then
                If OptPrice > 7500 Then
                    TotalPrice = Round_Up((TotalPrice + JU), -1)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU), 0)
                End If
            Else
                TotalPrice = Round_Up(((TotalPrice + JU) * KURS_TENGE), -1)
            End If

        ElseIf TorudaKomponent = False Then
            If NCity < 123 Or NCity >= 143 Then
                If OptPrice > 7500 Then
                    TotalPrice = Round_Up((TotalPrice + JU), -1)
                ElseIf OptPrice >= 1000 Then
                    TotalPrice = Round_Up((TotalPrice + JU), 0)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU + 200), 0)
                End If
        
            ElseIf NCity >= 123 And NCity < 135 Then
                If OptPrice >= 1000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_TENGE, -1)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_TENGE, -1)
                End If
            
            ElseIf NCity >= 135 And NCity < 141 Then
                If OptPrice > 150000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, -1)
                ElseIf OptPrice >= 1000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, 0)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_BELRUB, 0)
                End If
            
            ElseIf NCity = 141 Then
                If OptPrice > 5000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_SOM, -1)
                ElseIf OptPrice >= 1000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_SOM, 0)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_SOM, 0)
                End If
        
            ElseIf NCity = 142 Then
                If OptPrice >= 1000 Then
                    TotalPrice = Round_Up((TotalPrice + JU) * KURS_DRAM, -1)
                Else
                    TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_DRAM, -1)
                End If
            End If
        End If

        'Заполнение инфостолбца "GorodInfo" в прогрузке, если не стоит галочка "Без РРЦ"
zapis:

        'Если в юзерформе стоит галочка на "Без РРЦ"
        'Переходит к unrrc
        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
        
        'Приведение TotalPrice к РРЦ, если подходит под условия:
    
        'Основное условие: если TotalPrice < РРЦ
    
        '1. Компонент торуда, не СНГ, с оптом > 7500 - приводится к РРЦ до -1
        '2. Компонент торуда, не СНГ, с оптом < 7500 - приводится к РРЦ до 0
        '3. Компонент торуда, не СНГ - приводится к (РРЦ * Курс тенге) до -1
    
        '4. Компонент не торуда, не СНГ: аналогично (п.1 и п.2)
        '5. Казахстан - аналогично п.3
        '6. Беларусь - если РРЦ > 150 тыс., приводится к РРЦ до -1
        '7. Беларусь - если РРЦ < 150 тыс., приводится к РРЦ до 0
        '8. Киргизия - если РРЦ > 5 тыс., приводится к РРЦ до -1
        '9. Киргизия - если РРЦ < 5 тыс., приводится к РРЦ до 0
        '10. Армения - приводится к РРЦ до -1
    
        'GorodInfo по № городу = Предыдущий GorodInfo & № города :: TotalPrice :: Межтерминалка :: Полные транспортные :: DPD
        ProgruzkaMainArray(NRowProgruzka, 168) = ProgruzkaMainArray(NRowProgruzka, 168) & ProgruzkaMainArray(1, NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
    
        If TorudaKomponent = True Then
            If TotalPrice < RetailPrice Then
                If NCity < 123 Or NCity >= 135 And NCity < 163 Then
                    If OptPriceProschet > 7500 Then
                        TotalPrice = Round_Up(RetailPrice, -1)
                    Else
                        TotalPrice = Round_Up(RetailPrice, 0)
                    End If
                Else
                    TotalPrice = Round_Up(RetailPrice * KURS_TENGE, -1)
                End If
            End If
        ElseIf TorudaKomponent = False Then
            If NCity < 123 Then
                If TotalPrice < RetailPrice Then
                    If OptPriceProschet > 7500 Then
                        TotalPrice = Round_Up(RetailPrice, -1)
                    Else
                        TotalPrice = Round_Up(RetailPrice, 0)
                    End If
                End If
            ElseIf NCity >= 123 And NCity < 135 Then
                If (TotalPrice / KURS_TENGE) < RetailPrice Then
                    TotalPrice = Round_Up(RetailPrice * KURS_TENGE, -1)
                End If
            ElseIf NCity >= 135 And NCity < 141 Then
                If (TotalPrice / KURS_BELRUB) < RetailPrice Then
                    If RetailPrice > 150000 Then
                        TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, -1)
                    ElseIf RetailPrice <= 150000 Then
                        TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, 0)
                    End If
                End If
            ElseIf NCity = 141 Then
                If (TotalPrice / KURS_SOM) < RetailPrice Then
                    If RetailPrice > 5000 Then
                        TotalPrice = Round_Up(RetailPrice * KURS_SOM, -1)
                    ElseIf RetailPrice <= 5000 Then
                        TotalPrice = Round_Up(RetailPrice * KURS_SOM, 0)
                    End If
                End If
            ElseIf NCity = 142 Then
                If (TotalPrice / KURS_DRAM) < RetailPrice Then
                    TotalPrice = Round_Up(RetailPrice * KURS_DRAM, -1)
                End If
            End If
        End If
    
        'Заполняются ценой (без уникализации) столбцы одной строки в прогрузке по городам с ценой(TotalPrice)
unrrc:
        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next
    'Конец цикла записи цен в прогрузку по столбцам в одной строке


    'Общая уникализация по городам в рамках одного кода и компонента:
    If UserForm1.CheckBox1 = False Then
        Call fMixer(ProgruzkaMainArray, TotalCitiesCount, NRowProgruzka)
    End If

    'Федеральная цена по Оренбургу:
    ProgruzkaMainArray(NRowProgruzka, 156) = ProgruzkaMainArray(NRowProgruzka, 59)
            
    'Добавление к инфостолбцу GorodInfo федеральной цены без транспортных
    ProgruzkaMainArray(NRowProgruzka, 168) = ProgruzkaMainArray(NRowProgruzka, 168) & "125" & "::" & ProgruzkaMainArray(NRowProgruzka, 156) & "::" & "0" & "::" & "0" & "::" & "DPD "

    'Если в юзерформе галочка в "Без РРЦ" не поставлена,
    'И федеральная цена < РРЦ, тогда:
    If UserForm1.CheckBox2 = False Then
        If ProgruzkaMainArray(NRowProgruzka, 156) < RetailPrice Then
            If OptPriceProschet > 7500 Then
                ProgruzkaMainArray(NRowProgruzka, 156) = Round_Up(RetailPrice, -1)
            Else
                ProgruzkaMainArray(NRowProgruzka, 156) = Round_Up(RetailPrice, 0)
            End If
        End If
    End If
                   
    'Переход на новую строку в прогрузке, смена компонента того же кода 1С
Change_Row_Progruzka:
    Do While smkat <> ((UBound(mvvodin) + 1) - 5)
        smkat = smkat + 5
        NcolumnProgruzka = 5
        If Len(ProgruzkaMainArray(NRowProgruzka, 2)) <> 0 Then
            NRowProgruzka = NRowProgruzka + 1
        End If
        GoTo Komponent_Change
    Loop

    'Если значение в столбце ОПТ выдает ошибку, тогда
    'Переходит к Resetting
    If IsError(OptPriceProschet) Then
        GoTo Resetting
    End If

    'Если значение в столбце ОПТ не числовое, или равно нулю, тогда
    'Переходит к Resetting
    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Resetting
    End If

    'Если компонент в прогрузке отсутствует, тогда
    'Счетчик строки сокращается на 1
    If Len(ProgruzkaMainArray(NRowProgruzka, 2)) = 0 Then
        NRowProgruzka = NRowProgruzka - 1
    End If

    'Уникализация в рамках одного кода на одном городе по разным компонентам
    r3 = (NRowProgruzka - (UBound(mvvodin) + 1) / 5 + 1 + TriggerTeplozhar + TriggerKompkresla + TriggerPromo + TriggerKiska + TriggerPMK)
    For NcolumnProgruzka = 5 To 156
        For r1 = r3 To NRowProgruzka
            PriceStep = 0
            TotalPrice = ProgruzkaMainArray(r1, NcolumnProgruzka)

            'Россия
            If NcolumnProgruzka < 108 Then
                If TotalPrice < 7500 Then
                    PriceStep = 1
                Else
                    PriceStep = 10
                End If
                'Казахстан
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                PriceStep = 10
                'Беларусь
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                PriceStep = 1
                'Киргизия
            ElseIf NcolumnProgruzka = 126 Then
                If TotalPrice < 7500 Then
                    PriceStep = 1
                Else
                    PriceStep = 10
                End If
                'Армения
            ElseIf NcolumnProgruzka = 127 Then
                PriceStep = 10
                'Россия
            ElseIf NcolumnProgruzka >= 128 And NcolumnProgruzka < 149 Then
                If TotalPrice < 7500 Then
                    PriceStep = 1
                Else
                    PriceStep = 10
                End If
                'Казахстан
            ElseIf NcolumnProgruzka >= 149 Then
                PriceStep = 10
            End If

            For i = r3 To NRowProgruzka
                If i <> r1 Then
                    If TotalPrice = ProgruzkaMainArray(i, NcolumnProgruzka) Then
                        ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + PriceStep
                        TotalPrice = TotalPrice + PriceStep
                        i = r3 - 1
                    End If
                End If
            Next i
        Next r1
    Next NcolumnProgruzka


    'Переход к следующему коду 1С
    'Cycle3 = Resetting
Resetting:
    Do While d <> UBound(NRowsArray)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        smkat = 0
        TriggerPMK = 0
        TriggerTeplozhar = 0
        TriggerKompkresla = 0
        TriggerPromo = 0
        TriggerKiska = 0
        If Len(ProgruzkaMainArray(NRowProgruzka, 2)) <> 0 Then
            NRowProgruzka = NRowProgruzka + 1
        End If
        GoTo Kod1C_Change
    Loop

    If Transitive = False Then
        Postavshik = vbNullString
    End If

    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    smkat = 0

    If Len(ProgruzkaMainArray(2, 3)) <> 0 Then

        'Из прогрузки берется федеральная цена и инфостолбцы
        'Если есть триггерторуда, тогда прогрузка разделяется на две:
        'ProgruzkaSOTArray = для СОТ
        'mropgruzkat = для Торуды
        If TriggerToruda <> 0 Then
            ReDim ProgruzkaSOTArray(1 To NRowProgruzka, 1 To 140)
            ReDim ProgruzkaTorudaArray(1 To TriggerToruda + 1, 1 To 168)
    
            For i1 = 1 To 127
                'Значения из прогрузки переносятся в прогрузку СОТ
                'Значения из прогрузки переносятся в прогрузку Торуда
                ProgruzkaSOTArray(1, i1) = ProgruzkaMainArray(1, i1)
                ProgruzkaTorudaArray(1, i1) = ProgruzkaMainArray(1, i1)
            Next i1
        
            'Добавление городов и инфостолбцов к прогрузке SOT
            ProgruzkaSOTArray(1, 128) = ProgruzkaMainArray(1, 156)
            ProgruzkaSOTArray(1, 129) = ProgruzkaMainArray(1, 157)
            ProgruzkaSOTArray(1, 130) = ProgruzkaMainArray(1, 158)
            ProgruzkaSOTArray(1, 131) = ProgruzkaMainArray(1, 159)
            ProgruzkaSOTArray(1, 132) = ProgruzkaMainArray(1, 160)
            ProgruzkaSOTArray(1, 133) = ProgruzkaMainArray(1, 161)
            ProgruzkaSOTArray(1, 134) = ProgruzkaMainArray(1, 162)
            ProgruzkaSOTArray(1, 135) = ProgruzkaMainArray(1, 163)
            ProgruzkaSOTArray(1, 136) = ProgruzkaMainArray(1, 164)
            ProgruzkaSOTArray(1, 137) = ProgruzkaMainArray(1, 165)
            ProgruzkaSOTArray(1, 138) = ProgruzkaMainArray(1, 166)
            ProgruzkaSOTArray(1, 139) = ProgruzkaMainArray(1, 167)
            ProgruzkaSOTArray(1, 140) = ProgruzkaMainArray(1, 168)
    
            'Добавление городов и инфостолбцов к прогрузке Торуды
            ProgruzkaTorudaArray(1, 128) = "126"
            ProgruzkaTorudaArray(1, 129) = "127"
            ProgruzkaTorudaArray(1, 130) = "128"
            ProgruzkaTorudaArray(1, 131) = "129"
            ProgruzkaTorudaArray(1, 132) = "130"
            ProgruzkaTorudaArray(1, 133) = "131"
            ProgruzkaTorudaArray(1, 134) = "132"
            ProgruzkaTorudaArray(1, 135) = "133"
            ProgruzkaTorudaArray(1, 136) = "134"
            ProgruzkaTorudaArray(1, 137) = "135"
            ProgruzkaTorudaArray(1, 138) = "136"
            ProgruzkaTorudaArray(1, 139) = "137"
            ProgruzkaTorudaArray(1, 140) = "138"
            ProgruzkaTorudaArray(1, 141) = "139"
            ProgruzkaTorudaArray(1, 142) = "140"
            ProgruzkaTorudaArray(1, 143) = "141"
            ProgruzkaTorudaArray(1, 144) = "142"
            ProgruzkaTorudaArray(1, 145) = "143"
            ProgruzkaTorudaArray(1, 146) = "144"
            ProgruzkaTorudaArray(1, 147) = "145"
            ProgruzkaTorudaArray(1, 148) = "146"
            ProgruzkaTorudaArray(1, 149) = "147"
            ProgruzkaTorudaArray(1, 150) = "148"
            ProgruzkaTorudaArray(1, 151) = "149"
            ProgruzkaTorudaArray(1, 152) = "150"
            ProgruzkaTorudaArray(1, 153) = "151"
            ProgruzkaTorudaArray(1, 154) = "152"
            ProgruzkaTorudaArray(1, 155) = "153"
            ProgruzkaTorudaArray(1, 156) = "125"
            ProgruzkaTorudaArray(1, 157) = "Gorod"
            ProgruzkaTorudaArray(1, 158) = "Nazenka"
            ProgruzkaTorudaArray(1, 159) = "Typezab"
            ProgruzkaTorudaArray(1, 160) = "Objem"
            ProgruzkaTorudaArray(1, 161) = "Ves"
            ProgruzkaTorudaArray(1, 162) = "Kolvo"
            ProgruzkaTorudaArray(1, 163) = "Ju"
            ProgruzkaTorudaArray(1, 164) = "RetailPrice"
            ProgruzkaTorudaArray(1, 165) = "Zakup"
            ProgruzkaTorudaArray(1, 166) = "Postavshik"
            ProgruzkaTorudaArray(1, 167) = "Naimenovanie"
            ProgruzkaTorudaArray(1, 168) = "Gorodinfo"
    
            f = 0
            f1 = 0
        
            'От 2 до кол-ва строк в прогрузке
            For i = 2 To NRowProgruzka
                'Если в компоненте нет t
                'f - счетчик строк для прогрузки СОТ
                If InStr(ProgruzkaMainArray(i, 2), "t") = 0 Then
                    f = f + 1
                    For i1 = 1 To 127
                        'Значения из внутренней прогрузки переносятся в прогрузку СОТ
                        ProgruzkaSOTArray(f + 1, i1) = ProgruzkaMainArray(i, i1)
                    Next i1
                    'Заполняются инфостолбцы в прогрузке СОТ
                    ProgruzkaSOTArray(f + 1, 128) = ProgruzkaMainArray(i, 156)
                    ProgruzkaSOTArray(f + 1, 129) = ProgruzkaMainArray(i, 157)
                    ProgruzkaSOTArray(f + 1, 130) = ProgruzkaMainArray(i, 158)
                    ProgruzkaSOTArray(f + 1, 131) = ProgruzkaMainArray(i, 159)
                    ProgruzkaSOTArray(f + 1, 132) = ProgruzkaMainArray(i, 160)
                    ProgruzkaSOTArray(f + 1, 133) = ProgruzkaMainArray(i, 161)
                    ProgruzkaSOTArray(f + 1, 134) = ProgruzkaMainArray(i, 162)
                    ProgruzkaSOTArray(f + 1, 135) = ProgruzkaMainArray(i, 163)
                    ProgruzkaSOTArray(f + 1, 136) = ProgruzkaMainArray(i, 164)
                    ProgruzkaSOTArray(f + 1, 137) = ProgruzkaMainArray(i, 165)
                    ProgruzkaSOTArray(f + 1, 138) = ProgruzkaMainArray(i, 166)
                    ProgruzkaSOTArray(f + 1, 139) = ProgruzkaMainArray(i, 167)
                    ProgruzkaSOTArray(f + 1, 140) = ProgruzkaMainArray(i, 168)
                Else
                    'f1 - счетчик строк для прогрузки Торуда
                    f1 = f1 + 1
                    For i1 = 1 To 168
                        'Значения из внутренней прогрузки переносятся в прогрузку Торуда
                        ProgruzkaTorudaArray(f1 + 1, i1) = ProgruzkaMainArray(i, i1)
                    Next i1
                End If
    
            Next i
    
            i1 = 0
            f = 0
            f1 = 0
    
            'В лист прогрузки СОТ и Торуда переносятся значения
            Application.ScreenUpdating = False
        
            'СОТ
            If Len(ProgruzkaSOTArray(2, 3)) <> 0 Then
                Workbooks.Add
                ProgruzkaFilename = ActiveWorkbook.name
                ProgruzkaSheetname = ActiveSheet.name
                Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EJ" & NRowProgruzka).Value = ProgruzkaSOTArray
            
                If UserForm1.Autosave = True Then
                    Call Autosave(Postavshik, ProgruzkaFilename, "Б2С СОТ")
                End If
            
            End If
        
            'ТОРУДА
            If Len(ProgruzkaTorudaArray(2, 3)) <> 0 Then
                Workbooks.Add
                ProgruzkaFilename = ActiveWorkbook.name
                ProgruzkaSheetname = ActiveSheet.name
                Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:FL" & TriggerToruda + 1).Value = ProgruzkaTorudaArray
                If UserForm1.Autosave = True Then
                    Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Торуда")
                End If
            End If
            Application.ScreenUpdating = True
    
        Else
    
            Workbooks.Add
            ProgruzkaFilename = ActiveWorkbook.name
            ProgruzkaSheetname = ActiveSheet.name
        
            'Добавление городов и инфостолбцов к прогрузке SOT
            For i = 1 To UBound(ProgruzkaMainArray, 1)
                ProgruzkaMainArray(i, 128) = ProgruzkaMainArray(i, 156)
                ProgruzkaMainArray(i, 129) = ProgruzkaMainArray(i, 157)
                ProgruzkaMainArray(i, 130) = ProgruzkaMainArray(i, 158)
                ProgruzkaMainArray(i, 131) = ProgruzkaMainArray(i, 159)
                ProgruzkaMainArray(i, 132) = ProgruzkaMainArray(i, 160)
                ProgruzkaMainArray(i, 133) = ProgruzkaMainArray(i, 161)
                ProgruzkaMainArray(i, 134) = ProgruzkaMainArray(i, 162)
                ProgruzkaMainArray(i, 135) = ProgruzkaMainArray(i, 163)
                ProgruzkaMainArray(i, 136) = ProgruzkaMainArray(i, 164)
                ProgruzkaMainArray(i, 137) = ProgruzkaMainArray(i, 165)
                ProgruzkaMainArray(i, 138) = ProgruzkaMainArray(i, 166)
                ProgruzkaMainArray(i, 139) = ProgruzkaMainArray(i, 167)
                ProgruzkaMainArray(i, 140) = ProgruzkaMainArray(i, 168)
            Next i
            'Значения из внутренней прогрузки переносятся на лист
            Application.ScreenUpdating = False
            Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EJ" & NRowProgruzka).Value = ProgruzkaMainArray
            Application.ScreenUpdating = True
    
        
            'Если стоит галочка в юзерформе "Автосохранение", тогда
            'Проставляется наименование книги и сохраняется в путь Environ("PROGRUZKA")
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, ProgruzkaFilename, "Б2С СОТ")
            End If
    
        End If
    End If

    If TotalRowsPMK <> 0 Then Call Прогрузка_ПМК(PMKArray, TotalRowsPMK)
    
    If TotalRowsTeplozhar <> 0 Then Call Прогрузка_Тепложар(TeplozharArray, TotalRowsTeplozhar)

    If TotalRowsKompkresla <> 0 Then Call Прогрузка_Компкресла(KompkreslaArray, TotalRowsKompkresla)
    
    If TotalRowsKiska <> 0 Then Call Прогрузка_Киска(KiskaArray, TotalRowsKiska)

    ReDim TeplozharPromoArray(3, TotalRowsPromo)
    ReDim KompkreslaPromoArray(3, TotalRowsPromo)
    ReDim SOTPromoArray(4, TotalRowsPromo)
    ReDim KiskaPromoArray(3, TotalRowsPromo)
    ReDim PMKPromoArray(4, TotalRowsPromo)

    If TotalRowsPromo <> 0 Then
    
        If TriggerTorudaPromo <> 0 Then
            UserForm4.Show
        End If
        
        If TotalRowsPromo <> TriggerTorudaPromo Then
            UserForm2.Show
        End If
        For i = 0 To UBound(mvvod)
            For TotalRowsPromo = 0 To UBound(TotalPromoArray, 2)
                If TotalPromoArray(4, TotalRowsPromo, i) = "НП" Then
                    KompkreslaPromoArray(0, TotalRowsKompkreslaPromo) = TotalPromoArray(0, TotalRowsPromo, i)
                    KompkreslaPromoArray(1, TotalRowsKompkreslaPromo) = TotalPromoArray(1, TotalRowsPromo, i)
                    KompkreslaPromoArray(2, TotalRowsKompkreslaPromo) = TotalPromoArray(2, TotalRowsPromo, i)
                    KompkreslaPromoArray(3, TotalRowsKompkreslaPromo) = TotalPromoArray(3, TotalRowsPromo, i)
                    TotalRowsKompkreslaPromo = TotalRowsKompkreslaPromo + 1
                ElseIf TotalPromoArray(4, TotalRowsPromo, i) = "620П" Or TotalPromoArray(4, TotalRowsPromo, i) = "620tП" Then
                    TeplozharPromoArray(0, TotalRowsTeplozharPromo) = TotalPromoArray(0, TotalRowsPromo, i)
                    TeplozharPromoArray(1, TotalRowsTeplozharPromo) = TotalPromoArray(1, TotalRowsPromo, i)
                    TeplozharPromoArray(2, TotalRowsTeplozharPromo) = TotalPromoArray(2, TotalRowsPromo, i)
                    TeplozharPromoArray(3, TotalRowsTeplozharPromo) = TotalPromoArray(3, TotalRowsPromo, i)
                    TotalRowsTeplozharPromo = TotalRowsTeplozharPromo + 1
                ElseIf TotalPromoArray(4, TotalRowsPromo, i) = "1116П" Or TotalPromoArray(4, TotalRowsPromo, i) = "1116tП" Then
                    KiskaPromoArray(0, TotalRowsKiskaPromo) = TotalPromoArray(0, TotalRowsPromo, i)
                    KiskaPromoArray(1, TotalRowsKiskaPromo) = TotalPromoArray(1, TotalRowsPromo, i)
                    KiskaPromoArray(2, TotalRowsKiskaPromo) = TotalPromoArray(2, TotalRowsPromo, i)
                    KiskaPromoArray(3, TotalRowsKiskaPromo) = TotalPromoArray(3, TotalRowsPromo, i)
                    TotalRowsKiskaPromo = TotalRowsKiskaPromo + 1
                ElseIf TotalPromoArray(4, TotalRowsPromo, i) = "796П" Or TotalPromoArray(4, TotalRowsPromo, i) = "1034П" Then
                    PMKPromoArray(0, TotalRowsPMKPromo) = TotalPromoArray(0, TotalRowsPromo, i)
                    PMKPromoArray(1, TotalRowsPMKPromo) = TotalPromoArray(1, TotalRowsPromo, i)
                    PMKPromoArray(2, TotalRowsPMKPromo) = TotalPromoArray(2, TotalRowsPromo, i)
                    PMKPromoArray(3, TotalRowsPMKPromo) = TotalPromoArray(3, TotalRowsPromo, i)
                    PMKPromoArray(4, TotalRowsPMKPromo) = TotalPromoArray(4, TotalRowsPromo, i)
                    TotalRowsPMKPromo = TotalRowsPMKPromo + 1
                ElseIf Len(TotalPromoArray(4, TotalRowsPromo, i)) <> 0 Then
                    SOTPromoArray(0, TotalRowsSOTPromo) = TotalPromoArray(0, TotalRowsPromo, i)
                    SOTPromoArray(1, TotalRowsSOTPromo) = TotalPromoArray(1, TotalRowsPromo, i)
                    SOTPromoArray(2, TotalRowsSOTPromo) = TotalPromoArray(2, TotalRowsPromo, i)
                    SOTPromoArray(3, TotalRowsSOTPromo) = TotalPromoArray(3, TotalRowsPromo, i)
                    SOTPromoArray(4, TotalRowsSOTPromo) = TotalPromoArray(4, TotalRowsPromo, i)
                    TotalRowsSOTPromo = TotalRowsSOTPromo + 1
                End If
            Next TotalRowsPromo
        Next i
    End If

    If TotalRowsPMKPromo <> 0 Then
        Call Прогрузка_Промо_ПМК(PMKPromoArray, TotalRowsPMKPromo)
    End If

    If TotalRowsTeplozharPromo <> 0 Then
        Call Прогрузка_Промо_Тепложар(TeplozharPromoArray, TotalRowsTeplozharPromo)
    End If

    If TotalRowsKompkreslaPromo <> 0 Then
        Call Прогрузка_Промо_Компкресла(KompkreslaPromoArray, TotalRowsKompkreslaPromo)
    End If

    If TotalRowsKiskaPromo <> 0 Then
        Call Прогрузка_Промо_Киска(KiskaPromoArray, TotalRowsKiskaPromo)
    End If

    If TotalRowsSOTPromo <> 0 Then
        Call Прогрузка_Промо_СОТ_Торуда(SOTPromoArray, TriggerTorudaPromo)
    End If

    'If TriggerTeplo = 1 Then
    '    Call Изменения_Тепложар(mteplotoxls, mteplo, TriggerTeplo)
    'End If

ended:
    Unload UserForm1

End Sub

'Промо SOT и Toruda
Public Sub Прогрузка_Промо_СОТ_Торуда(ByRef PromoArray, ByRef TriggerTorudaPromo)

    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NCity As Long
    Dim r1 As Long
    Dim NRowProgruzka As Long
    Dim d As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a
    Dim OptPrice
    Dim Postavshik As String
    Dim mvvodin() As String
    Dim CitiesArray() As String
    Dim MinTerminalTerminalPrice As Double
    Dim PlatniiVes
    Dim at() As String
    Dim mvvodint()
    Dim ProgruzkaSOTArray As Variant
    Dim ProgruzkaTorudaArray As Variant
    Dim TotalCitiesCount As Long
    Dim TypeZabor As String
    Dim Transitive As Boolean
    Dim NcolumnProgruzka As Long
    Dim smkat As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim TorudaKomponent As Boolean
    Dim NeiskluchenieToruda As Boolean
    Dim Tarif As Double
    Dim r4 As Long
    Dim r3 As Long
    Dim i1 As Long
    Dim f As Long
    Dim f1 As Long
    Dim i As Long
    Dim TriggerTeplo As Long
    
    Transitive = True

    For i = UBound(PromoArray, 2) To 0 Step -1
        If PromoArray(0, i) = vbNullString Then
            ReDim Preserve PromoArray(4, UBound(PromoArray, 2) - 1)
        End If
    Next

    ReDim ProgruzkaMainArray(1 To UBound(PromoArray, 2) + 2, 1 To 168)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С СОТ/Торуда")
    
    ReDim CitiesArray(20 To 170)
    CitiesArray = CreateCities(CitiesArray, "Б2С СОТ/Торуда")
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    smkat = 0
    NRowProgruzka = 2

    If UserForm4.CheckBox_Margin_Promo = False Then
        If UserForm4.TextBox1 = 0 And TriggerTorudaPromo <> 0 Or Len(UserForm4.TextBox1) = 0 And TriggerTorudaPromo <> 0 Then
            MsgBox "Некорректная наценка/категория"
            GoTo ended
        Else
            at = Split(UserForm4.TextBox1, ";")
        End If
    Else
        at = Split(UserForm4.Label2.Caption, ";")
    End If
       
    If UserForm4.CheckBox_Margin_Promo = True Or Len(UserForm4.TextBox1) <> 0 Then
        ReDim mvvodint(UBound(at))

        For i = 0 To UBound(at)
            If Len(at(i)) = 0 Then
                ReDim Preserve mvvodint(UBound(mvvodint) - 1)
            Else
                mvvodint(i - (UBound(at) - UBound(mvvodint))) = at(i)
            End If
        Next

        If (UBound(mvvodint) + 1) Mod 3 <> 0 Then
            MsgBox "Некорректная наценка/категория"
            GoTo ended
        End If
    End If

    If UserForm2.CheckBox_Margin_Promo = True Then
        a = Split(UserForm2.Label2.Caption, ";")
    ElseIf UserForm2.TextBox1 <> 0 And Len(UserForm2.TextBox1) <> 0 Then
        a = Split(UserForm2.TextBox1, ";")
    End If

    If UserForm2.CheckBox_Margin_Promo = True Or Len(UserForm2.TextBox1) <> 0 Then
        ReDim mvvodin(UBound(a))

        For i = 0 To UBound(a)
            If Len(a(i)) = 0 Then
                ReDim Preserve mvvodin(UBound(mvvodin) - 1)
            Else
                mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
            End If
        Next

        If (UBound(mvvodin) + 1) Mod 3 <> 0 Then
            MsgBox "Некорректная наценка/категория"
            GoTo ended
        End If
    End If


Kod1C_Change:
    
    NRowProschet = PromoArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    Dim Artikul As String, GorodOtpravki As String, TypeZaborProschet As String, Kod1CProschet As String, Komplektuyeshee As String
    Dim Quantity As Long, RetailPrice As Long, OptPriceProschet As Long, RetailPriceInfo As Long
    Dim Volume As Double, VolumetricWeight As Double, RazgruzkaPrice As Double, ZaborPriceProschet As Double, JU As Double, Weight As Double
    Artikul = ProschetArray(NRowProschet, 2)
    Quantity = ProschetArray(NRowProschet, 5)
    GorodOtpravki = ProschetArray(NRowProschet, 6)
    TypeZaborProschet = ProschetArray(NRowProschet, 8)
    Volume = ProschetArray(NRowProschet, 9)
    Weight = ProschetArray(NRowProschet, 10)
    VolumetricWeight = ProschetArray(NRowProschet, 11)
    RazgruzkaPrice = ProschetArray(NRowProschet, 12)
    ZaborPriceProschet = ProschetArray(NRowProschet, 13)
    JU = ProschetArray(NRowProschet, 14)
    Komplektuyeshee = ProschetArray(NRowProschet, 16)
    Kod1CProschet = ProschetArray(NRowProschet, 17)
    RetailPrice = ProschetArray(NRowProschet, 18)
    OptPriceProschet = ProschetArray(NRowProschet, 19)
    RetailPriceInfo = ProschetArray(NRowProschet, 20)
    
    If Volume * 1440 < 950 Then
        ObreshetkaPrice = 950
    Else
        ObreshetkaPrice = Volume * 1440
    End If

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ProschetArray(NRowProschet, 13) <> 0 Then
        ZaborPrice = 0
        LgotniiGorod = True
    Else
        ZaborPrice = ProschetArray(NRowProschet, 13)
        LgotniiGorod = False
    End If
    
    Dim tps As Double
    tps = fTPS(Weight, Quantity)

    If d <> 0 And ProschetArray(NRowProschet, 1) = Postavshik Then
        GoTo Komponent_Change
    ElseIf d <> 0 Then
        Transitive = False
    End If

    Postavshik = ProschetArray(NRowProschet, 1)

Komponent_Change:
    
    Dim KomponentNumber As Long
    KomponentNumber = Val(PromoArray(4, d))
    If KomponentNumber = 562 Or KomponentNumber = 618 Or KomponentNumber = 620 _
    Or KomponentNumber = 680 Or KomponentNumber = 681 Or KomponentNumber = 682 _
    Or KomponentNumber = 683 Or KomponentNumber = 708 Or KomponentNumber = 768 _
    Or KomponentNumber = 775 Or KomponentNumber = 777 Or KomponentNumber = 786 _
    Or KomponentNumber = 793 Or KomponentNumber = 821 Or KomponentNumber = 834 _
    Or KomponentNumber = 849 Or KomponentNumber = 997 Then
        TriggerTeplo = 1
    Else
        TriggerTeplo = 0
    End If
    
    TorudaKomponent = False

    'Если в компоненте есть t,
    'К триггеру плюсуется 1
    If InStr(PromoArray(4, d), "t") <> 0 Then
        TorudaKomponent = True
    End If

    'Определяет - является ли код подходящий под условия для исключения торуда (опт <= 1000)
    NeiskluchenieToruda = False
    If TorudaKomponent = True Then
        If OptPriceProschet * Quantity <= 1000 Then
            If DictIsklucheniyaToruda.exists(Kod1CProschet) = False Then
                NeiskluchenieToruda = True
            End If
        End If
    End If

    TypeZabor = FillInfocolumnTypeZabor(PromoArray(4, d), ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    If TorudaKomponent = True Then
        Margin = fMargin(OptPriceProschet, Quantity, 0, 1, at)
    Else
        Margin = fMargin(OptPriceProschet, Quantity, 0, 1, a)
    End If

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = Replace(PromoArray(4, d), "П", vbNullString)
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 157) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 158) = FillInfocolumnMargin(Komplektuyeshee, Margin, NeiskluchenieToruda, True)
    ProgruzkaMainArray(NRowProgruzka, 159) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 160) = Volume
    ProgruzkaMainArray(NRowProgruzka, 161) = Weight
    ProgruzkaMainArray(NRowProgruzka, 162) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 163) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 164) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 165) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 166) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 167) = Artikul

    If NeiskluchenieToruda = True Then
        ProgruzkaMainArray(NRowProgruzka, 163) = "нет"
    End If
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    If TorudaKomponent = True Then
        TotalCitiesCount = 170
    Else
        TotalCitiesCount = 142
    End If

    For NCity = 20 To TotalCitiesCount
        
        If Komplektuyeshee = "Т" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = OptPriceProschet * Margin
                ElseIf NCity >= 123 And NCity < 135 Then
                    TotalPrice = OptPriceProschet * Margin * KURS_TENGE
                ElseIf NCity >= 135 And NCity < 163 Then
                    TotalPrice = OptPriceProschet * Margin
                Else
                    TotalPrice = OptPriceProschet * Margin * KURS_TENGE
                End If
            ElseIf NCity < 123 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            ElseIf NCity = 142 Then
                TotalPrice = OptPriceProschet * Margin * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К2" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2, 0)
                ElseIf NCity < 135 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
                ElseIf NCity < 163 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2, 0)
                Else
                    TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
                End If
            ElseIf NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * 2, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_DRAM, 0)
            End If
            GoTo zapis
            
        ElseIf NeiskluchenieToruda = True Then
            If NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.35, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * (1.35 + SNG_MARGIN) * KURS_TENGE, 0)
            ElseIf NCity < 143 Then
                TotalPrice = Round_Up(OptPriceProschet * (1.35 + SNG_MARGIN), 0)
            ElseIf NCity < 163 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.35, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * (1.35 + SNG_MARGIN) * KURS_TENGE, 0)
            End If
            GoTo zapis
    
        
        ElseIf Komplektuyeshee = "К" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2, 0)
                ElseIf NCity < 135 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
                ElseIf NCity < 163 Then
                    TotalPrice = Round_Up(OptPriceProschet * 2, 0)
                Else
                    TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
                End If
            ElseIf NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * 2, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * 2 * KURS_DRAM, 0)
            End If
            GoTo zapis
        
        ElseIf Komplektuyeshee = "К1" Then
            If TorudaKomponent = True Then
                If NCity < 123 Then
                    TotalPrice = Round_Up(OptPriceProschet * 1.7, 0)
                ElseIf NCity < 135 Then
                    TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_TENGE, 0)
                ElseIf NCity < 163 Then
                    TotalPrice = Round_Up(OptPriceProschet * 1.7, 0)
                Else
                    TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_TENGE, 0)
                End If
            ElseIf NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.7, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * 1.7 * KURS_DRAM, 0)
            End If
            GoTo zapis
        End If
        
        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity >= 123 And NCity < 143 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 163 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        Else
            Price = OptPriceProschet * Margin
        End If

        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))

        If NCity >= 123 And NCity < 143 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        ElseIf NCity >= 163 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        Else
            PlatniiVes = VolumetricWeight
        End If
        
        If PlatniiVes * Tarif < MinTerminalTerminalPrice Then
            TerminalTerminalPrice = MinTerminalTerminalPrice
        Else
            TerminalTerminalPrice = PlatniiVes * Tarif
        End If
        
        If NCity = 44 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If
        
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps

        TotalPrice = Price + FullDostavkaPrice
        
        If TorudaKomponent = True Then
            If NCity < 123 Or NCity >= 135 And NCity < 163 Then
                If OptPrice > 7500 Then
                    TotalPrice = Round_Up(TotalPrice + JU, -1)
                Else
                    TotalPrice = Round_Up(TotalPrice + JU, 0)
                End If
            Else
                TotalPrice = Round_Up(((TotalPrice + JU) * KURS_TENGE), -1)
            End If
        ElseIf NCity < 123 Or NCity >= 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(TotalPrice + JU, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            Else
                TotalPrice = Round_Up(TotalPrice + JU + 200, 0)
            End If
        ElseIf NCity < 135 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200)) * KURS_TENGE, -1)
            End If
        ElseIf NCity < 141 Then
            If OptPrice > 150000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_BELRUB, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_BELRUB, 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200)) * KURS_BELRUB, 0)
            End If
        ElseIf NCity < 142 Then
            If OptPrice > 5000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_SOM, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200)) * KURS_SOM, 0)
            End If
        ElseIf NCity = 142 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_DRAM, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200)) * KURS_DRAM, -1)
            End If
        End If

zapis:

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
        
        ProgruzkaMainArray(NRowProgruzka, 168) = ProgruzkaMainArray(NRowProgruzka, 168) & ProgruzkaMainArray(1, NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
        
        If TorudaKomponent = True Then
            If TotalPrice < RetailPrice Then
                If NCity < 123 Or NCity >= 135 And NCity < 163 Then
                    If OptPriceProschet > 7500 Then
                        TotalPrice = Round_Up(RetailPrice, -1)
                    Else
                        TotalPrice = Round_Up(RetailPrice, 0)
                    End If
                Else
                    TotalPrice = Round_Up(RetailPrice * KURS_TENGE, -1)
                End If
            End If
        ElseIf NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If OptPriceProschet > 7500 Then
                    TotalPrice = Round_Up(RetailPrice, -1)
                Else
                    TotalPrice = Round_Up(RetailPrice, 0)
                End If
            End If
        ElseIf NCity >= 123 And NCity < 135 Then
            If (TotalPrice / KURS_TENGE) < RetailPrice Then
                TotalPrice = Round_Up(RetailPrice * KURS_TENGE, -1)
            End If
        ElseIf NCity >= 135 And NCity < 141 Then
            If (TotalPrice / KURS_BELRUB) < RetailPrice Then
                If RetailPrice > 150000 Then
                    TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, -1)
                Else
                    TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, 0)
                End If
            End If
        ElseIf NCity >= 141 And NCity < 142 Then
            If (TotalPrice / KURS_SOM) < RetailPrice Then
                If RetailPrice > 5000 Then
                    TotalPrice = Round_Up(RetailPrice * KURS_SOM, -1)
                ElseIf RetailPrice <= 5000 Then
                    TotalPrice = Round_Up(RetailPrice * KURS_SOM, 0)
                End If
            End If
        ElseIf NCity = 142 Then
            If (TotalPrice / KURS_DRAM) < RetailPrice Then
                TotalPrice = Round_Up(RetailPrice * KURS_DRAM, -1)
            End If
        End If
unrrc:

        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К2" Then
        For NcolumnProgruzka = 5 To TotalCitiesCount - 15
            If TorudaKomponent = True Then
                If NcolumnProgruzka < 108 Or NcolumnProgruzka >= 120 And NcolumnProgruzka < 148 Then
                    If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                    ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                    Else
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                    End If
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                End If
            ElseIf NcolumnProgruzka < 108 Then
                If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 5000 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
                End If
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If

        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If

    If NeiskluchenieToruda = True Then
        For NcolumnProgruzka = 5 To TotalCitiesCount - 15
            ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then
        For NcolumnProgruzka = 5 To TotalCitiesCount - 15
        
            If TorudaKomponent = True Then
                If NcolumnProgruzka < 108 Or NcolumnProgruzka >= 120 And NcolumnProgruzka < 148 Then
                    If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                    ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                    Else
                        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                    End If
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                End If
            
            ElseIf NcolumnProgruzka < 108 Then
                If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
            
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 5000 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
                End If
            
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If
        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If

    Call fMixer(ProgruzkaMainArray, TotalCitiesCount, NRowProgruzka)

federal_orenburg:
   
    ProgruzkaMainArray(NRowProgruzka, 156) = ProgruzkaMainArray(NRowProgruzka, 59)
    GoTo orenburg_infocolumn
    
orenburg_infocolumn:
    
    ProgruzkaMainArray(NRowProgruzka, 168) = ProgruzkaMainArray(NRowProgruzka, 168) & "125" & "::" & ProgruzkaMainArray(NRowProgruzka, 156) & "::" & "0" & "::" & "0" & "::" & "DPD "
    
    If UserForm1.CheckBox2 = False Then
        If ProgruzkaMainArray(NRowProgruzka, 156) < RetailPrice Then
            If OptPriceProschet > 7500 Then
                ProgruzkaMainArray(NRowProgruzka, 156) = Round_Up(RetailPrice, -1)
            Else
                ProgruzkaMainArray(NRowProgruzka, 156) = Round_Up(RetailPrice, 0)
            End If
        End If
    End If

    Postavshik = ProschetArray(NRowProschet, 1)

Change_Row_Progruzka:

    If ProgruzkaMainArray(NRowProgruzka, 5) = 0 Then
        GoTo Resetting
    ElseIf d = 0 Then
        GoTo Resetting
    ElseIf PromoArray(2, d) <> PromoArray(2, d - 1) Then
        GoTo Resetting
    ElseIf UBound(PromoArray, 2) <> d Then
        If PromoArray(2, d) = PromoArray(2, d + 1) Then
            GoTo Resetting
        End If
    End If

    r4 = 0
    For i = 0 To UBound(PromoArray, 2)
        If PromoArray(2, i) = PromoArray(2, d) Then
            r4 = r4 + 1
        End If
    Next i

    r4 = d - r4 + 3

    r3 = d
    For NcolumnProgruzka = 5 To 156
        For r1 = r4 To r3 + 2
            TotalPrice = ProgruzkaMainArray(r1, NcolumnProgruzka)
            For i = r4 To r3 + 2
                If TotalPrice = ProgruzkaMainArray(i, NcolumnProgruzka) Then
                    If i <> r1 Then
                        If InStr(ProgruzkaMainArray(i, 2), "t") <> 0 Then
                            If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            ElseIf ProgruzkaMainArray(r4, 4) > 7500 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r4 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            End If
                        ElseIf NcolumnProgruzka < 108 Then
                            If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            ElseIf ProgruzkaMainArray(r4, 4) > 7500 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r4 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r4 - 1
                        ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                            If ProgruzkaMainArray(r4, 4) > 150000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r4 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                            If ProgruzkaMainArray(r4, 4) > 5000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r4 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            End If
                        ElseIf NcolumnProgruzka = 127 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r4 - 1
                        ElseIf NcolumnProgruzka = 156 Then
                            If DictSpecPostavshik.exists(ProschetArray(PromoArray(2, d), 1)) = True Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            ElseIf ProgruzkaMainArray(r4, 4) > 7500 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r4 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r4 - 1
                            End If
                        End If
                    End If
                End If
            Next i
        Next r1
    Next NcolumnProgruzka

Resetting:
    Do While d <> UBound(PromoArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        smkat = 0
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change
    Loop

    If Transitive = False Then
        Postavshik = vbNullString
    End If

    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    smkat = 0

    If Len(ProgruzkaMainArray(2, 3)) <> 0 Then
        If TriggerTorudaPromo <> 0 Then
            ReDim ProgruzkaSOTArray(1 To NRowProgruzka, 1 To 140)
            ReDim ProgruzkaTorudaArray(1 To TriggerTorudaPromo + 1, 1 To 168)
    
            For i1 = 1 To 127
                ProgruzkaSOTArray(1, i1) = ProgruzkaMainArray(1, i1)
                ProgruzkaTorudaArray(1, i1) = ProgruzkaMainArray(1, i1)
            Next i1
    
            ProgruzkaSOTArray(1, 128) = ProgruzkaMainArray(1, 156)
            ProgruzkaSOTArray(1, 129) = ProgruzkaMainArray(1, 157)
            ProgruzkaSOTArray(1, 130) = ProgruzkaMainArray(1, 158)
            ProgruzkaSOTArray(1, 131) = ProgruzkaMainArray(1, 159)
            ProgruzkaSOTArray(1, 132) = ProgruzkaMainArray(1, 160)
            ProgruzkaSOTArray(1, 133) = ProgruzkaMainArray(1, 161)
            ProgruzkaSOTArray(1, 134) = ProgruzkaMainArray(1, 162)
            ProgruzkaSOTArray(1, 135) = ProgruzkaMainArray(1, 163)
            ProgruzkaSOTArray(1, 136) = ProgruzkaMainArray(1, 164)
            ProgruzkaSOTArray(1, 137) = ProgruzkaMainArray(1, 165)
            ProgruzkaSOTArray(1, 138) = ProgruzkaMainArray(1, 166)
            ProgruzkaSOTArray(1, 139) = ProgruzkaMainArray(1, 167)
            ProgruzkaSOTArray(1, 140) = ProgruzkaMainArray(1, 168)
    
            ProgruzkaTorudaArray(1, 128) = "126"
            ProgruzkaTorudaArray(1, 129) = "127"
            ProgruzkaTorudaArray(1, 130) = "128"
            ProgruzkaTorudaArray(1, 131) = "129"
            ProgruzkaTorudaArray(1, 132) = "130"
            ProgruzkaTorudaArray(1, 133) = "131"
            ProgruzkaTorudaArray(1, 134) = "132"
            ProgruzkaTorudaArray(1, 135) = "133"
            ProgruzkaTorudaArray(1, 136) = "134"
            ProgruzkaTorudaArray(1, 137) = "135"
            ProgruzkaTorudaArray(1, 138) = "136"
            ProgruzkaTorudaArray(1, 139) = "137"
            ProgruzkaTorudaArray(1, 140) = "138"
            ProgruzkaTorudaArray(1, 141) = "139"
            ProgruzkaTorudaArray(1, 142) = "140"
            ProgruzkaTorudaArray(1, 143) = "141"
            ProgruzkaTorudaArray(1, 144) = "142"
            ProgruzkaTorudaArray(1, 145) = "143"
            ProgruzkaTorudaArray(1, 146) = "144"
            ProgruzkaTorudaArray(1, 147) = "145"
            ProgruzkaTorudaArray(1, 148) = "146"
            ProgruzkaTorudaArray(1, 149) = "147"
            ProgruzkaTorudaArray(1, 150) = "148"
            ProgruzkaTorudaArray(1, 151) = "149"
            ProgruzkaTorudaArray(1, 152) = "150"
            ProgruzkaTorudaArray(1, 153) = "151"
            ProgruzkaTorudaArray(1, 154) = "152"
            ProgruzkaTorudaArray(1, 155) = "153"
            ProgruzkaTorudaArray(1, 156) = "125"
            ProgruzkaTorudaArray(1, 157) = "Gorod"
            ProgruzkaTorudaArray(1, 158) = "Nazenka"
            ProgruzkaTorudaArray(1, 159) = "Typezab"
            ProgruzkaTorudaArray(1, 160) = "Objem"
            ProgruzkaTorudaArray(1, 161) = "Ves"
            ProgruzkaTorudaArray(1, 162) = "Kolvo"
            ProgruzkaTorudaArray(1, 163) = "Ju"
            ProgruzkaTorudaArray(1, 164) = "RetailPrice"
            ProgruzkaTorudaArray(1, 165) = "Zakup"
            ProgruzkaTorudaArray(1, 166) = "Postavshik"
            ProgruzkaTorudaArray(1, 167) = "Naimenovanie"
            ProgruzkaTorudaArray(1, 168) = "Gorodinfo"
    
            f = 0
            f1 = 0
            For i = 2 To NRowProgruzka
                If InStr(ProgruzkaMainArray(i, 2), "t") = 0 Then
                    f = f + 1
                    For i1 = 1 To 127
                        ProgruzkaSOTArray(f + 1, i1) = ProgruzkaMainArray(i, i1)
                    Next i1
                    ProgruzkaSOTArray(f + 1, 128) = ProgruzkaMainArray(i, 156)
                    ProgruzkaSOTArray(f + 1, 129) = ProgruzkaMainArray(i, 157)
                    ProgruzkaSOTArray(f + 1, 130) = ProgruzkaMainArray(i, 158)
                    ProgruzkaSOTArray(f + 1, 131) = ProgruzkaMainArray(i, 159)
                    ProgruzkaSOTArray(f + 1, 132) = ProgruzkaMainArray(i, 160)
                    ProgruzkaSOTArray(f + 1, 133) = ProgruzkaMainArray(i, 161)
                    ProgruzkaSOTArray(f + 1, 134) = ProgruzkaMainArray(i, 162)
                    ProgruzkaSOTArray(f + 1, 135) = ProgruzkaMainArray(i, 163)
                    ProgruzkaSOTArray(f + 1, 136) = ProgruzkaMainArray(i, 164)
                    ProgruzkaSOTArray(f + 1, 137) = ProgruzkaMainArray(i, 165)
                    ProgruzkaSOTArray(f + 1, 138) = ProgruzkaMainArray(i, 166)
                    ProgruzkaSOTArray(f + 1, 139) = ProgruzkaMainArray(i, 167)
                    ProgruzkaSOTArray(f + 1, 140) = ProgruzkaMainArray(i, 168)
                Else
                    f1 = f1 + 1
                    For i1 = 1 To 168
                        ProgruzkaTorudaArray(f1 + 1, i1) = ProgruzkaMainArray(i, i1)
                    Next i1
                End If
    
            Next i
    
            i1 = 0
            f = 0
            f1 = 0
            
            Application.ScreenUpdating = False
            
            'SOT
            If Len(ProgruzkaSOTArray(2, 3)) <> 0 Then
                Workbooks.Add
                ProgruzkaFilename = ActiveWorkbook.name
                ProgruzkaSheetname = ActiveSheet.name
                Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EJ" & NRowProgruzka).Value = ProgruzkaSOTArray
                
                If UserForm1.Autosave = True Then
                    Call Autosave(Postavshik, ProgruzkaFilename, "Б2С СОТ Промо")
                End If
            End If
            
            'TORUDA
            If Len(ProgruzkaTorudaArray(2, 3)) <> 0 Then
                Workbooks.Add
                ProgruzkaFilename = ActiveWorkbook.name
                ProgruzkaSheetname = ActiveSheet.name
                Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:FL" & TriggerTorudaPromo + 1).Value = ProgruzkaTorudaArray
                
                If UserForm1.Autosave = True Then
                    Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Торуда Промо")
                End If
            End If
            Application.ScreenUpdating = True
    
        Else
            
            Workbooks.Add
            ProgruzkaFilename = ActiveWorkbook.name
            ProgruzkaSheetname = ActiveSheet.name
            
            For i = 1 To UBound(ProgruzkaMainArray, 1)
                ProgruzkaMainArray(i, 128) = ProgruzkaMainArray(i, 156)
                ProgruzkaMainArray(i, 129) = ProgruzkaMainArray(i, 157)
                ProgruzkaMainArray(i, 130) = ProgruzkaMainArray(i, 158)
                ProgruzkaMainArray(i, 131) = ProgruzkaMainArray(i, 159)
                ProgruzkaMainArray(i, 132) = ProgruzkaMainArray(i, 160)
                ProgruzkaMainArray(i, 133) = ProgruzkaMainArray(i, 161)
                ProgruzkaMainArray(i, 134) = ProgruzkaMainArray(i, 162)
                ProgruzkaMainArray(i, 135) = ProgruzkaMainArray(i, 163)
                ProgruzkaMainArray(i, 136) = ProgruzkaMainArray(i, 164)
                ProgruzkaMainArray(i, 137) = ProgruzkaMainArray(i, 165)
                ProgruzkaMainArray(i, 138) = ProgruzkaMainArray(i, 166)
                ProgruzkaMainArray(i, 139) = ProgruzkaMainArray(i, 167)
                ProgruzkaMainArray(i, 140) = ProgruzkaMainArray(i, 168)
            Next i
            
            Application.ScreenUpdating = False
            Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EJ" & NRowProgruzka).Value = ProgruzkaMainArray
            
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, ProgruzkaFilename, "Б2С СОТ Промо")
            End If
            
            Application.ScreenUpdating = True
    
        End If
    End If

ended:
    Unload UserForm2
    Unload UserForm4
End Sub


