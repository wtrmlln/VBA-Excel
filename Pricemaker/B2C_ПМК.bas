Attribute VB_Name = "B2C_ПМК"
Option Private Module
Option Explicit

Public Sub Прогрузка_ПМК(ByRef PMKArray, ByRef TotalRowsPMK)
    
    Dim ProgruzkaFilename As String
    Dim NCity As Long
    Dim d As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a() As String
    Dim OptPrice
    Dim MinTerminalTerminalPrice As Double
    Dim str As String
    
    Dim ProgruzkaPMKArray() As Variant
    Dim PlatniiVes
    Dim CitiesArray() As String
    
    Dim Transitive As Boolean
    Dim NcolumnProgruzka As Long
    Dim NRowProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim TypeZabor As String
    Dim TotalCitiesCount As Long
    Dim ProgruzkaSheetname As String
    Dim Tarif As Double
    Dim Postavshik As String
    Dim i As Long
    Dim j As Long
    Dim tps As Double
    Dim ArrayNRow() As Long
    
    For i = UBound(PMKArray, 2) To 0 Step -1
        If PMKArray(0, i) = 0 Or Len(PMKArray(0, i)) = 0 Then
            ReDim Preserve PMKArray(3, UBound(PMKArray, 2) - 1)
        End If
    Next

    Transitive = True
    
    ReDim CitiesArray(20 To 143)
    CitiesArray = CreateCities(CitiesArray, "Б2С ПМК")
    
    ReDim ProgruzkaPMKArray(1 To (TotalRowsPMK + 1), 1 To 140)
    ProgruzkaPMKArray = CreateHeader(ProgruzkaPMKArray, "Б2С ПМК")
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

Kod1C_Change:

    NRowProschet = PMKArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    a = Split(ProschetArray(NRowProschet, 7), ";")

    str = a(4 + PMKArray(3, d))
    
    If str = "1034" Then
        j = j + 1
        ReDim Preserve ArrayNRow(1 To j)
        ArrayNRow(j) = NRowProschet
    End If
    
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

    tps = fTPS(Weight, Quantity)

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ProschetArray(NRowProschet, 13) <> 0 Then
        ZaborPrice = 200
        LgotniiGorod = True
    Else
        ZaborPrice = ProschetArray(NRowProschet, 13)
        LgotniiGorod = False
    End If

    If d <> 0 And ProschetArray(NRowProschet, 1) = Postavshik Then
        GoTo Komponent_Change
    ElseIf d <> 0 Then
        Transitive = False
    End If

Komponent_Change:
    TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, ProschetArray(NRowProschet, 12), NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    Margin = fMargin(OptPriceProschet, Quantity, PMKArray(3, d), 0, a)
    
    'Заполнение инфостолбцов в прогрузке:
    '99 - Gorod = "Город" из просчета
    '101 - TypeZabor = по FillInfocolumnTypeZabor
    '102 - Volume = "объем" из просчета
    '103 - Ves = "вес" из просчета
    '104 - Kolvo = "Кол-во штук" из просчета
    '107 - Zakup = Закуп из просчета
    '108 - Postavshik = Поставщик из просчета
    '109 - Naimenovanie = Артикул из просчета
    ProgruzkaPMKArray(NRowProgruzka, 1) = a(4 + PMKArray(3, d))
    ProgruzkaPMKArray(NRowProgruzka, 2) = Artikul
    ProgruzkaPMKArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaPMKArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaPMKArray(NRowProgruzka, 129) = GorodOtpravki
    ProgruzkaPMKArray(NRowProgruzka, 130) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, False)
    ProgruzkaPMKArray(NRowProgruzka, 131) = TypeZabor
    ProgruzkaPMKArray(NRowProgruzka, 132) = Volume
    ProgruzkaPMKArray(NRowProgruzka, 133) = Weight
    ProgruzkaPMKArray(NRowProgruzka, 134) = Quantity
    ProgruzkaPMKArray(NRowProgruzka, 135) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaPMKArray(NRowProgruzka, 136) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaPMKArray(NRowProgruzka, 137) = OptPriceProschet
    ProgruzkaPMKArray(NRowProgruzka, 138) = ProschetArray(NRowProschet, 1)
    ProgruzkaPMKArray(NRowProgruzka, 139) = Artikul
    'Заполнение инфостолбца "Наценка"

    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To UBound(CitiesArray)
        
        If Komplektuyeshee = "Т" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
            
        End If
        
        If NCity = 100 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity < 107 Then
            Price = OptPriceProschet * Margin
        ElseIf NCity >= 107 And NCity < 114 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 114 And NCity < 116 Then
            Price = OptPriceProschet * (Margin)
        ElseIf NCity >= 116 And NCity < 122 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 122 And NCity < 124 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 124 And NCity < 143 Then
            Price = OptPriceProschet * Margin
        ElseIf NCity = 143 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        End If

        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))

        If NCity >= 107 And NCity < 114 _
        Or NCity >= 116 And NCity < 122 _
        Or NCity >= 122 And NCity < 124 _
        Or NCity = 143 Then
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
        
        If NCity = 100 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet, "Kab") + tps

        TotalPrice = Price + FullDostavkaPrice
       
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
        
        If NCity < 107 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity >= 107 And NCity < 114 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)) * KURS_TENGE, -1)
            End If
        ElseIf NCity >= 114 And NCity < 116 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity >= 116 And NCity < 122 Then
            TotalPrice = Round_Up((TotalPrice + RazgruzkaPrice + JU) * KURS_BELRUB, 0)
        ElseIf NCity >= 122 And NCity < 124 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)) * KURS_TENGE, -1)
            End If
        ElseIf NCity >= 124 And NCity < 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity = 143 Then
            TotalPrice = Round_Up((TotalPrice + RazgruzkaPrice + JU) * KURS_SOM, 0)
        End If

zapis:

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
               
        'Заполнение инфостолбца "GorodInfo"
        If str = "796" Then
            If NcolumnProgruzka <= 98 Then
                ProgruzkaPMKArray(NRowProgruzka, 140) = ProgruzkaPMKArray(NRowProgruzka, 140) & ProgruzkaPMKArray(1, NcolumnProgruzka) & _
                "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
            End If
        ElseIf str = "1034" Then
            If NcolumnProgruzka <= 133 Then
                ProgruzkaPMKArray(NRowProgruzka, 140) = ProgruzkaPMKArray(NRowProgruzka, 140) & ProgruzkaPMKArray(1, NcolumnProgruzka) & _
                "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
            End If
        End If
        
        If NcolumnProgruzka < 91 And TotalPrice < RetailPrice And OptPriceProschet > 7500 _
        Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And TotalPrice < RetailPrice And OptPriceProschet > 7500 _
        Or NcolumnProgruzka >= 124 And NcolumnProgruzka < 143 And TotalPrice < RetailPrice And OptPriceProschet > 7500 Then
            TotalPrice = Round_Up(RetailPrice, -1)
        ElseIf NcolumnProgruzka < 92 And TotalPrice < RetailPrice _
        Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And TotalPrice < RetailPrice _
        Or NcolumnProgruzka >= 124 And NcolumnProgruzka < 143 And TotalPrice < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice, 0)
        ElseIf NcolumnProgruzka >= 107 And NcolumnProgruzka < 114 And TotalPrice / KURS_TENGE < RetailPrice _
        Or NcolumnProgruzka >= 122 And NcolumnProgruzka < 124 And TotalPrice / KURS_TENGE < RetailPrice Then
            TotalPrice = Round_Up(((RetailPrice * (1))) * KURS_TENGE, -1)
        ElseIf NcolumnProgruzka >= 116 And NcolumnProgruzka < 122 And TotalPrice / KURS_BELRUB < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * (1) * KURS_BELRUB, 0)
        ElseIf NcolumnProgruzka = 143 And TotalPrice / KURS_SOM < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * (1) * KURS_SOM, 0)
        End If

unrrc:

        If str = "1034" Then
            ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        ElseIf str = "796" Then
            If NcolumnProgruzka > 98 Then
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka)
            Else
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
            End If
        End If
        NcolumnProgruzka = NcolumnProgruzka + 1
        
    Next
    
    If str = "796" Then
        TotalCitiesCount = 98
    
    ElseIf str = "1034" Then
        TotalCitiesCount = 128
    End If
    
    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Or Komplektuyeshee = "К2" Then

        For NcolumnProgruzka = 5 To TotalCitiesCount
            If ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) > 0 Then
                
                If NcolumnProgruzka < 91 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka = 128 Then
                    ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka < 91 _
                Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 _
                Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 _
                Or NcolumnProgruzka = 128 Then
                    ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka >= 91 And NcolumnProgruzka < 99 _
                Or NcolumnProgruzka >= 107 And NcolumnProgruzka < 109 Then
                    ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka >= 101 And NcolumnProgruzka < 107 Then
                    ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
                
            End If
        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If
    
    If str = "796" Then
        Call fMixer_ПМК(ProgruzkaPMKArray, UBound(CitiesArray) - 29, NRowProgruzka)
        GoTo federal_orenburg
    ElseIf str = "1034" Then
        Call fMixer_ПМК(ProgruzkaPMKArray, UBound(CitiesArray), NRowProgruzka)
        GoTo federal_orenburg
    End If

federal_orenburg:
    For NcolumnProgruzka = 5 To TotalCitiesCount
        If ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) > 0 Then
            
            If NcolumnProgruzka < 91 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 And ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka = 127 Then
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 10 * (PMKArray(1, d) + 1) / 5
            ElseIf NcolumnProgruzka < 91 _
            Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 _
            Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 _
            Or NcolumnProgruzka = 128 Then
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 1 * (PMKArray(1, d) + 1) / 5
            ElseIf NcolumnProgruzka >= 91 And NcolumnProgruzka < 99 _
            Or NcolumnProgruzka >= 107 And NcolumnProgruzka < 109 Then
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 10 * (PMKArray(1, d) + 1) / 5
            ElseIf NcolumnProgruzka >= 101 And NcolumnProgruzka < 107 Then
                ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKArray(NRowProgruzka, NcolumnProgruzka) + 1 * (PMKArray(1, d) + 1) / 5
            End If
            
        End If

    Next NcolumnProgruzka

Change_Row_Progruzka:
    Postavshik = ProschetArray(PMKArray(2, d), 1)

    Do While d <> UBound(PMKArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change
    Loop

    If Transitive = False Then
        Postavshik = vbNullString
    End If
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    
    Dim f As Long
    Dim f1 As Long
    Dim mprogruzka_1034() As Variant
    Dim mprogruzka_cleared() As Variant
    Dim column_progruzka As Long
    Dim i1 As Long
    Dim ProgruzkaClearedName As String
    Dim ProgruzkaClearedSheet As String
    Dim Progruzka1034Name As String
    Dim Progruzka1034Sheet As String
            
    If (Not ArrayNRow) <> -1 Then
        f = 1
        f1 = 1
        ReDim Preserve mprogruzka_1034(1 To UBound(ArrayNRow) + 1, 1 To UBound(ProgruzkaPMKArray, 2))
        ReDim Preserve mprogruzka_cleared(1 To (UBound(ProgruzkaPMKArray, 1) - UBound(mprogruzka_1034, 1) + 1), 1 To (UBound(ProgruzkaPMKArray, 2) - 29))
    
        mprogruzka_1034(1, 99) = "74"
        mprogruzka_1034(1, 100) = "79"
        mprogruzka_1034(1, 101) = "85"
        mprogruzka_1034(1, 102) = "86"
        mprogruzka_1034(1, 103) = "87"
        mprogruzka_1034(1, 104) = "88"
        mprogruzka_1034(1, 105) = "89"
        mprogruzka_1034(1, 106) = "90"
        mprogruzka_1034(1, 107) = "95"
        mprogruzka_1034(1, 107) = "98"
        mprogruzka_1034(1, 109) = "101"
        mprogruzka_1034(1, 110) = "105"
        mprogruzka_1034(1, 111) = "107"
        mprogruzka_1034(1, 112) = "107"
        mprogruzka_1034(1, 113) = "109"
        mprogruzka_1034(1, 114) = "110"
        mprogruzka_1034(1, 115) = "111"
        mprogruzka_1034(1, 116) = "113"
        mprogruzka_1034(1, 117) = "114"
        mprogruzka_1034(1, 118) = "115"
        mprogruzka_1034(1, 119) = "116"
        mprogruzka_1034(1, 120) = "117"
        mprogruzka_1034(1, 121) = "118"
        mprogruzka_1034(1, 122) = "119"
        mprogruzka_1034(1, 123) = "120"
        mprogruzka_1034(1, 124) = "121"
        mprogruzka_1034(1, 125) = "124"
        mprogruzka_1034(1, 126) = "125"
        mprogruzka_1034(1, 127) = "129"
        mprogruzka_1034(1, 128) = "130"

        mprogruzka_1034(1, 129) = "Gorod"
        mprogruzka_1034(1, 130) = "Nazenka"
        mprogruzka_1034(1, 131) = "Typezab"
        mprogruzka_1034(1, 132) = "Objem"
        mprogruzka_1034(1, 133) = "Ves"
        mprogruzka_1034(1, 134) = "Kolvo"
        mprogruzka_1034(1, 135) = "Ju"
        mprogruzka_1034(1, 136) = "Rrc"
        mprogruzka_1034(1, 137) = "Zakup"
        mprogruzka_1034(1, 138) = "Postavshik"
        mprogruzka_1034(1, 139) = "Naimenovanie"
        mprogruzka_1034(1, 140) = "Gorodinfo"
    
        mprogruzka_cleared(1, 99) = "Gorod"
        mprogruzka_cleared(1, 100) = "Nazenka"
        mprogruzka_cleared(1, 101) = "Typezab"
        mprogruzka_cleared(1, 102) = "Objem"
        mprogruzka_cleared(1, 103) = "Ves"
        mprogruzka_cleared(1, 104) = "Kolvo"
        mprogruzka_cleared(1, 105) = "Ju"
        mprogruzka_cleared(1, 106) = "Rrc"
        mprogruzka_cleared(1, 107) = "Zakup"
        mprogruzka_cleared(1, 108) = "Postavshik"
        mprogruzka_cleared(1, 109) = "Naimenovanie"
        mprogruzka_cleared(1, 110) = "Gorodinfo"
    
        For i = 1 To UBound(mprogruzka_cleared, 2)
            If i < 99 Or i > 128 Then
                mprogruzka_cleared(1, f1) = ProgruzkaPMKArray(1, i)
                f1 = f1 + 1
            End If
        Next i
        f1 = 1
        For i = 1 To UBound(mprogruzka_1034, 2)
            mprogruzka_1034(1, i) = ProgruzkaPMKArray(1, i)
        Next i
        For i = 2 To UBound(ProgruzkaPMKArray, 1)
            column_progruzka = 1
            If ProgruzkaPMKArray(i, 1) = "1034" Then
                f = f + 1
                For i1 = 1 To UBound(mprogruzka_1034, 2)
                    mprogruzka_1034(f, i1) = ProgruzkaPMKArray(i, i1)
                Next i1
            Else
                f1 = f1 + 1
            
                For i1 = 1 To UBound(ProgruzkaPMKArray, 2)
                    If i1 < 99 Or i1 > 128 Then
                        mprogruzka_cleared(f1, column_progruzka) = ProgruzkaPMKArray(i, i1)
                        column_progruzka = column_progruzka + 1
                    End If
                Next i1
            End If
        Next i
    
        If UBound(mprogruzka_cleared, 1) > 1 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            ProgruzkaClearedName = ActiveWorkbook.name
            ProgruzkaClearedSheet = ActiveSheet.name
            Workbooks(ProgruzkaClearedName).Worksheets(ProgruzkaClearedSheet).Range("A1:DF" & UBound(mprogruzka_cleared, 1)).Value = mprogruzka_cleared
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, ProgruzkaClearedName, "Б2С ПМК")
            End If
            Application.ScreenUpdating = True
        End If
    
        If UBound(mprogruzka_1034, 1) > 1 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            Progruzka1034Name = ActiveWorkbook.name
            Progruzka1034Sheet = ActiveSheet.name
            Workbooks(Progruzka1034Name).Worksheets(Progruzka1034Sheet).Range("A1:EJ" & UBound(mprogruzka_1034, 1)).Value = mprogruzka_1034
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, Progruzka1034Name, "Б2С ПМК 1034")
            End If
            Application.ScreenUpdating = True
        End If
    Else
        f = 1
        f1 = 1
        ReDim Preserve mprogruzka_cleared(1 To (UBound(ProgruzkaPMKArray, 1) + 1), 1 To (UBound(ProgruzkaPMKArray, 2) - 29))
        mprogruzka_cleared(1, 99) = "Gorod"
        mprogruzka_cleared(1, 100) = "Nazenka"
        mprogruzka_cleared(1, 101) = "Typezab"
        mprogruzka_cleared(1, 102) = "Objem"
        mprogruzka_cleared(1, 103) = "Ves"
        mprogruzka_cleared(1, 104) = "Kolvo"
        mprogruzka_cleared(1, 105) = "Ju"
        mprogruzka_cleared(1, 106) = "Rrc"
        mprogruzka_cleared(1, 107) = "Zakup"
        mprogruzka_cleared(1, 108) = "Postavshik"
        mprogruzka_cleared(1, 109) = "Naimenovanie"
        mprogruzka_cleared(1, 110) = "Gorodinfo"
    
        For i = 1 To UBound(mprogruzka_cleared, 2)
            If i < 99 Or i > 128 Then
                mprogruzka_cleared(1, f1) = ProgruzkaPMKArray(1, i)
                f1 = f1 + 1
            End If
        Next i
        f1 = 1
        For i = 2 To UBound(ProgruzkaPMKArray, 1)
            f1 = f1 + 1
            f = 1
            For i1 = 1 To UBound(ProgruzkaPMKArray, 2)
                If i1 < 99 Or i1 > 128 Then
                    mprogruzka_cleared(f1, f) = ProgruzkaPMKArray(i, i1)
                    f = f + 1
                End If
            Next i1
        Next i
    
        Workbooks.Add
        ProgruzkaClearedName = ActiveWorkbook.name
        ProgruzkaClearedSheet = ActiveSheet.name
    
        Application.ScreenUpdating = False
        Workbooks(ProgruzkaClearedName).Worksheets(ProgruzkaClearedSheet).Range("A1:DF" & UBound(mprogruzka_cleared, 1)).Value = mprogruzka_cleared
        Application.ScreenUpdating = True
      
        If UserForm1.Autosave = True Then
            Call Autosave(Postavshik, ProgruzkaClearedName, "Б2С ПМК")
        End If
    End If

ended:

End Sub

Public Sub Прогрузка_Промо_ПМК(ByRef PMKPromoArray, ByRef TotalRowsPMKPromo)

    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NCity As Long
    Dim d As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a() As String
    Dim OptPrice
    Dim str As String
    Dim ProgruzkaPMKPromoArray As Variant, PlatniiVes, TypeZabor
    Dim CitiesArray() As String
    Dim MinTerminalTerminalPrice As Double
    Dim NcolumnProgruzka As Long
    Dim NRowProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Postavshik As String
    Dim Transitive As Boolean
    Dim TotalCitiesCount As Long
    Dim Tarif As Double
    Dim i As Long
    Dim tps As Double
    Dim ArrayNRow() As Long
    
    For i = UBound(PMKPromoArray, 2) To 0 Step -1
        If PMKPromoArray(0, i) = 0 Or PMKPromoArray(0, i) = vbNullString Then
            ReDim Preserve PMKPromoArray(4, UBound(PMKPromoArray, 2) - 1)
        End If
    Next
    
    ReDim CitiesArray(20 To 143)
    CitiesArray = CreateCities(CitiesArray, "Б2С ПМК")
    
    ReDim ProgruzkaPMKPromoArray(1 To (TotalRowsPMKPromo + 1), 1 To 140)
    ProgruzkaPMKPromoArray = CreateHeader(ProgruzkaPMKPromoArray, "Б2С ПМК")
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2
    
Kod1C_Change:
    
    NRowProschet = PMKPromoArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
        
    If UserForm2.CheckBox_Margin_Promo = False Then
        If UserForm2.TextBox1 = 0 Or UserForm2.TextBox1 = vbNullString Then
            MsgBox "Некорректная наценка/категория"
            GoTo ended
        Else
            a = Split(UserForm2.TextBox1, ";")
        End If
    ElseIf UserForm2.CheckBox_Margin_Promo = True Then
        a = Split(UserForm2.Label2.Caption, ";")
    End If
    
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
    
    tps = fTPS(Weight, Quantity)

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ZaborPriceProschet <> 0 Then
        ZaborPrice = 0
        LgotniiGorod = True
    Else
        ZaborPrice = ZaborPriceProschet
        LgotniiGorod = False
    End If
    
    If d <> 0 And ProschetArray(NRowProschet, 1) = Postavshik Then
        GoTo Komponent_Change
    ElseIf d <> 0 Then
        Transitive = False
    End If

Komponent_Change:
    str = Replace(PMKPromoArray(4, d), "П", vbNullString)
    
    TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, PMKPromoArray(3, d), 1, a)
    
    'Заполнение инфостолбцов в прогрузке:
    '99 - Gorod = "Город" из просчета
    '101 - TypeZabor = по FillInfocolumnTypeZabor
    '102 - ProschetArray(NRowProschet, 9) = "объем" из просчета
    '103 - Ves = "вес" из просчета
    '104 - Kolvo = "Кол-во штук" из просчета
    '107 - Zakup = Закуп из просчета
    '108 - Postavshik = Поставщик из просчета
    '109 - Naimenovanie = Артикул из просчета
    ProgruzkaPMKPromoArray(NRowProgruzka, 1) = str
    ProgruzkaPMKPromoArray(NRowProgruzka, 2) = Artikul
    ProgruzkaPMKPromoArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaPMKPromoArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaPMKPromoArray(NRowProgruzka, 129) = GorodOtpravki
    ProgruzkaPMKPromoArray(NRowProgruzka, 130) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, True)
    ProgruzkaPMKPromoArray(NRowProgruzka, 131) = TypeZabor
    ProgruzkaPMKPromoArray(NRowProgruzka, 132) = Volume
    ProgruzkaPMKPromoArray(NRowProgruzka, 133) = Weight
    ProgruzkaPMKPromoArray(NRowProgruzka, 134) = Quantity
    ProgruzkaPMKPromoArray(NRowProgruzka, 135) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaPMKPromoArray(NRowProgruzka, 136) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaPMKPromoArray(NRowProgruzka, 137) = OptPriceProschet
    ProgruzkaPMKPromoArray(NRowProgruzka, 138) = ProschetArray(PMKPromoArray(2, d), 1)
    ProgruzkaPMKPromoArray(NRowProgruzka, 139) = Artikul
    
    

    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To UBound(CitiesArray)
        
        If Komplektuyeshee = "Т" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 107 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 107 And NCity < 114 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 114 And NCity < 116 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 116 And NCity < 122 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity >= 122 And NCity < 124 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 124 And NCity < 143 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity = 143 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        End If
        
        If NCity = 100 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity < 107 Then
            Price = OptPriceProschet * Margin
        ElseIf NCity >= 107 And NCity < 114 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 114 And NCity < 116 Then
            Price = OptPriceProschet * (Margin)
        ElseIf NCity >= 116 And NCity < 122 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 122 And NCity < 124 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity >= 124 And NCity < 143 Then
            Price = OptPriceProschet * Margin
        ElseIf NCity = 143 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        End If
        
        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))
            
        If NCity >= 107 And NCity < 114 _
        Or NCity >= 116 And NCity < 122 _
        Or NCity >= 122 And NCity < 124 _
        Or NCity = 143 Then
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
        
        If NCity = 100 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet, "Kab") + tps

        TotalPrice = Price + FullDostavkaPrice
        
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
        
        If NCity < 107 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity >= 107 And NCity < 114 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)) * KURS_TENGE, -1)
            End If
        ElseIf NCity >= 114 And NCity < 116 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity >= 116 And NCity < 122 Then
            TotalPrice = Round_Up((TotalPrice + RazgruzkaPrice + JU) * KURS_BELRUB, 0)
        ElseIf NCity >= 122 And NCity < 124 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)) * KURS_TENGE, -1)
            End If
        ElseIf NCity >= 124 And NCity < 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU)), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + RazgruzkaPrice + JU + 200)), 0)
            End If
        ElseIf NCity = 143 Then
            TotalPrice = Round_Up((TotalPrice + RazgruzkaPrice + JU) * KURS_SOM, 0)
        End If
        
zapis:

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
        
        'Заполнение инфостолбца "GorodInfo"
        If str = "796" Then
            If NcolumnProgruzka <= 98 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, 140) = ProgruzkaPMKPromoArray(NRowProgruzka, 140) & ProgruzkaPMKPromoArray(1, NcolumnProgruzka) & _
                "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
            End If
        ElseIf str = "1034" Then
            If NcolumnProgruzka <= 133 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, 140) = ProgruzkaPMKPromoArray(NRowProgruzka, 140) & ProgruzkaPMKPromoArray(1, NcolumnProgruzka) & _
                "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
            End If
        End If
        
        If NcolumnProgruzka < 91 And TotalPrice < RetailPrice And OptPriceProschet > 7500 _
        Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And TotalPrice < RetailPrice And OptPriceProschet > 7500 _
        Or NcolumnProgruzka >= 124 And NcolumnProgruzka < 143 And TotalPrice < RetailPrice And OptPriceProschet > 7500 Then
            TotalPrice = Round_Up(RetailPrice, -1)
        ElseIf NcolumnProgruzka < 92 And TotalPrice < RetailPrice _
        Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And TotalPrice < RetailPrice _
        Or NcolumnProgruzka >= 124 And NcolumnProgruzka < 143 And TotalPrice < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice, 0)
        ElseIf NcolumnProgruzka >= 107 And NcolumnProgruzka < 114 And TotalPrice / KURS_TENGE < RetailPrice _
        Or NcolumnProgruzka >= 122 And NcolumnProgruzka < 124 And TotalPrice / KURS_TENGE < RetailPrice Then
            TotalPrice = Round_Up(((RetailPrice * (1))) * KURS_TENGE, -1)
        ElseIf NcolumnProgruzka >= 116 And NcolumnProgruzka < 122 And TotalPrice / KURS_BELRUB < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * (1) * KURS_BELRUB, 0)
        ElseIf NcolumnProgruzka = 143 And TotalPrice / KURS_SOM < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * (1) * KURS_SOM, 0)
        End If
        
unrrc:
        If str = "1034" Then
            ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        ElseIf str = "796" Then
            If NcolumnProgruzka > 98 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka)
            Else
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
            End If
        End If
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next
    
    TotalCitiesCount = 98

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If
   
    If Komplektuyeshee = "К" _
    Or Komplektuyeshee = "К1" _
    Or Komplektuyeshee = "К2" Then
        For NcolumnProgruzka = 5 To TotalCitiesCount
            If ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) > 0 Then
                
                If NcolumnProgruzka < 91 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
                Or NcolumnProgruzka = 128 Then
                    ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka < 91 _
                Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 _
                Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 _
                Or NcolumnProgruzka = 128 Then
                    ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka >= 91 And NcolumnProgruzka < 99 _
                Or NcolumnProgruzka >= 107 And NcolumnProgruzka < 109 Then
                    ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                ElseIf NcolumnProgruzka >= 101 And NcolumnProgruzka < 107 Then
                    ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
                
            End If
        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If
    
    If str = "796" Then
        Call fMixer_ПМК(ProgruzkaPMKPromoArray, UBound(CitiesArray) - 29, NRowProgruzka)
        GoTo federal_orenburg
    ElseIf str = "1034" Then
        Call fMixer_ПМК(ProgruzkaPMKPromoArray, UBound(CitiesArray), NRowProgruzka)
        GoTo federal_orenburg
    End If
       
federal_orenburg:
    For NcolumnProgruzka = 5 To TotalCitiesCount
        If ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) > 0 Then
            
            If NcolumnProgruzka < 91 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 And ProgruzkaPMKPromoArray(NRowProgruzka, 4) > 7500 _
            Or NcolumnProgruzka = 127 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 10 * (PMKPromoArray(1, d)) / 5
            ElseIf NcolumnProgruzka < 91 _
            Or NcolumnProgruzka >= 99 And NcolumnProgruzka < 101 _
            Or NcolumnProgruzka >= 109 And NcolumnProgruzka < 128 _
            Or NcolumnProgruzka = 128 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 1 * (PMKPromoArray(1, d)) / 5
            ElseIf NcolumnProgruzka >= 91 And NcolumnProgruzka < 99 _
            Or NcolumnProgruzka >= 107 And NcolumnProgruzka < 109 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 10 * (PMKPromoArray(1, d)) / 5
            ElseIf NcolumnProgruzka >= 101 And NcolumnProgruzka < 107 Then
                ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaPMKPromoArray(NRowProgruzka, NcolumnProgruzka) + 1 * (PMKPromoArray(1, d)) / 5
            End If
            
        End If

    Next NcolumnProgruzka
    
Change_Row_Progruzka:
    Postavshik = ProschetArray(NRowProschet, 1)
    
    Do While d <> UBound(PMKPromoArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change
    Loop
    
    If Transitive = False Then
        Postavshik = vbNullString
    End If
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    
    Dim f As Long
    Dim f1 As Long
    Dim mprogruzka_1034() As Variant
    Dim mprogruzka_cleared() As Variant
    Dim column_progruzka As Long
    Dim i1 As Long
    Dim ProgruzkaClearedName As String
    Dim ProgruzkaClearedSheet As String
    Dim Progruzka1034Name As String
    Dim Progruzka1034Sheet As String
    
    If (Not ArrayNRow) <> -1 Then
        f = 1
        f1 = 1
        ReDim Preserve mprogruzka_1034(1 To UBound(ArrayNRow) + 1, 1 To UBound(ProgruzkaPMKPromoArray, 2))
        ReDim Preserve mprogruzka_cleared(1 To (UBound(ProgruzkaPMKPromoArray, 1) - UBound(mprogruzka_1034, 1) + 1), 1 To (UBound(ProgruzkaPMKPromoArray, 2) - 29))
    
        mprogruzka_1034(1, 99) = "74"
        mprogruzka_1034(1, 100) = "79"
        mprogruzka_1034(1, 101) = "85"
        mprogruzka_1034(1, 102) = "86"
        mprogruzka_1034(1, 103) = "87"
        mprogruzka_1034(1, 104) = "88"
        mprogruzka_1034(1, 105) = "89"
        mprogruzka_1034(1, 106) = "90"
        mprogruzka_1034(1, 107) = "95"
        mprogruzka_1034(1, 108) = "98"
        mprogruzka_1034(1, 109) = "101"
        mprogruzka_1034(1, 110) = "105"
        mprogruzka_1034(1, 111) = "107"
        mprogruzka_1034(1, 112) = "107"
        mprogruzka_1034(1, 113) = "109"
        mprogruzka_1034(1, 114) = "110"
        mprogruzka_1034(1, 115) = "111"
        mprogruzka_1034(1, 116) = "113"
        mprogruzka_1034(1, 117) = "114"
        mprogruzka_1034(1, 118) = "115"
        mprogruzka_1034(1, 119) = "116"
        mprogruzka_1034(1, 120) = "117"
        mprogruzka_1034(1, 121) = "118"
        mprogruzka_1034(1, 122) = "119"
        mprogruzka_1034(1, 123) = "120"
        mprogruzka_1034(1, 124) = "121"
        mprogruzka_1034(1, 125) = "124"
        mprogruzka_1034(1, 126) = "125"
        mprogruzka_1034(1, 127) = "129"
        mprogruzka_1034(1, 128) = "130"

        mprogruzka_1034(1, 129) = "Gorod"
        mprogruzka_1034(1, 130) = "Nazenka"
        mprogruzka_1034(1, 131) = "Typezab"
        mprogruzka_1034(1, 132) = "Objem"
        mprogruzka_1034(1, 133) = "Ves"
        mprogruzka_1034(1, 134) = "Kolvo"
        mprogruzka_1034(1, 135) = "Ju"
        mprogruzka_1034(1, 136) = "Rrc"
        mprogruzka_1034(1, 137) = "Zakup"
        mprogruzka_1034(1, 138) = "Postavshik"
        mprogruzka_1034(1, 139) = "Naimenovanie"
        mprogruzka_1034(1, 140) = "Gorodinfo"
    
        mprogruzka_cleared(1, 99) = "Gorod"
        mprogruzka_cleared(1, 100) = "Nazenka"
        mprogruzka_cleared(1, 101) = "Typezab"
        mprogruzka_cleared(1, 102) = "Objem"
        mprogruzka_cleared(1, 103) = "Ves"
        mprogruzka_cleared(1, 104) = "Kolvo"
        mprogruzka_cleared(1, 105) = "Ju"
        mprogruzka_cleared(1, 106) = "Rrc"
        mprogruzka_cleared(1, 107) = "Zakup"
        mprogruzka_cleared(1, 108) = "Postavshik"
        mprogruzka_cleared(1, 109) = "Naimenovanie"
        mprogruzka_cleared(1, 110) = "Gorodinfo"
    
        For i = 1 To UBound(mprogruzka_cleared, 2)
            If i < 99 Or i > 128 Then
                mprogruzka_cleared(1, f1) = ProgruzkaPMKPromoArray(1, i)
                f1 = f1 + 1
            End If
        Next i
        f1 = 1
        For i = 1 To UBound(mprogruzka_1034, 2)
            mprogruzka_1034(1, i) = ProgruzkaPMKPromoArray(1, i)
        Next i
        For i = 2 To UBound(ProgruzkaPMKPromoArray, 1)
            column_progruzka = 1
            If ProgruzkaPMKPromoArray(i, 1) = "1034" Then
                f = f + 1
                For i1 = 1 To UBound(mprogruzka_1034, 2)
                    mprogruzka_1034(f, i1) = ProgruzkaPMKPromoArray(i, i1)
                Next i1
            Else
                f1 = f1 + 1
            
                For i1 = 1 To UBound(ProgruzkaPMKPromoArray, 2)
                    If i1 < 99 Or i1 > 128 Then
                        mprogruzka_cleared(f1, column_progruzka) = ProgruzkaPMKPromoArray(i, i1)
                        column_progruzka = column_progruzka + 1
                    End If
                Next i1
            End If
        Next i
    
        If UBound(mprogruzka_cleared, 1) > 1 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            ProgruzkaClearedName = ActiveWorkbook.name
            ProgruzkaClearedSheet = ActiveSheet.name
            Workbooks(ProgruzkaClearedName).Worksheets(ProgruzkaClearedSheet).Range("A1:DF" & UBound(mprogruzka_cleared, 1)).Value = mprogruzka_cleared
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, ProgruzkaClearedName, "Б2С ПМК Промо")
            End If
            Application.ScreenUpdating = True
        End If
    
        If UBound(mprogruzka_1034, 1) > 1 Then
            Application.ScreenUpdating = False
            Workbooks.Add
            Progruzka1034Name = ActiveWorkbook.name
            Progruzka1034Sheet = ActiveSheet.name
            Workbooks(Progruzka1034Name).Worksheets(Progruzka1034Sheet).Range("A1:EJ" & UBound(mprogruzka_1034, 1)).Value = mprogruzka_1034
            If UserForm1.Autosave = True Then
                Call Autosave(Postavshik, Progruzka1034Name, "Б2С ПМК 1034 Промо")
            End If
            Application.ScreenUpdating = True
        End If
    Else
        f = 1
        f1 = 1
        ReDim Preserve mprogruzka_cleared(1 To (UBound(ProgruzkaPMKPromoArray, 1) + 1), 1 To (UBound(ProgruzkaPMKPromoArray, 2) - 29))
        mprogruzka_cleared(1, 99) = "Gorod"
        mprogruzka_cleared(1, 100) = "Nazenka"
        mprogruzka_cleared(1, 101) = "Typezab"
        mprogruzka_cleared(1, 102) = "Objem"
        mprogruzka_cleared(1, 103) = "Ves"
        mprogruzka_cleared(1, 104) = "Kolvo"
        mprogruzka_cleared(1, 105) = "Ju"
        mprogruzka_cleared(1, 106) = "Rrc"
        mprogruzka_cleared(1, 107) = "Zakup"
        mprogruzka_cleared(1, 108) = "Postavshik"
        mprogruzka_cleared(1, 109) = "Naimenovanie"
        mprogruzka_cleared(1, 110) = "Gorodinfo"
    
        For i = 1 To UBound(mprogruzka_cleared, 2)
            If i < 99 Or i > 128 Then
                mprogruzka_cleared(1, f1) = ProgruzkaPMKPromoArray(1, i)
                f1 = f1 + 1
            End If
        Next i
        f1 = 1
        For i = 2 To UBound(ProgruzkaPMKPromoArray, 1)
            f1 = f1 + 1
            f = 1
            For i1 = 1 To UBound(ProgruzkaPMKPromoArray, 2)
                If i1 < 99 Or i1 > 128 Then
                    mprogruzka_cleared(f1, f) = ProgruzkaPMKPromoArray(i, i1)
                    f = f + 1
                End If
            Next i1
        Next i
    
        Workbooks.Add
        ProgruzkaClearedName = ActiveWorkbook.name
        ProgruzkaClearedSheet = ActiveSheet.name
    
        Application.ScreenUpdating = False
        Workbooks(ProgruzkaClearedName).Worksheets(ProgruzkaClearedSheet).Range("A1:DF" & UBound(mprogruzka_cleared, 1)).Value = mprogruzka_cleared
        Application.ScreenUpdating = True
      
        If UserForm1.Autosave = True Then
            Call Autosave(Postavshik, ProgruzkaClearedName, "Б2С ПМК Промо")
        End If
    End If
    
ended:
   
End Sub


