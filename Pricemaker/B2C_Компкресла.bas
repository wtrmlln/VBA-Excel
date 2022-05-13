Attribute VB_Name = "B2C_Компкресла"
Option Private Module
Option Explicit

Public Sub Прогрузка_Компкресла(ByRef KompkreslaArray, ByRef TotalRowsKompkresla)

    Dim ProgruzkaFilename As String
    Dim NCity As Long
    Dim r1 As Long
    Dim d As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a
    Dim OptPrice
    Dim MinTerminalTerminalPrice As Double
    Dim mdostavkadpd As Variant
    Dim CitiesArray() As String

    Dim ProgruzkaKompkreslaDverArray() As Variant
    Dim ProgruzkaKompkreslaTerminalArray() As Variant
    Dim ProgruzkaKompkreslaArray() As Variant
    Dim PlatniiVes
    Dim TypeZabor
    Dim ProgruzkaSheetname As String
    Dim TransferColumnsDict As Object
    Dim NcolumnProgruzka As Long
    Dim NRowProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Tarif As Double
    Dim Postavshik As String
    Dim i As Long
    Dim tps As Double
    
    Dim Transitive As Boolean
    Transitive = True

    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    ReDim ProgruzkaKompkreslaDverArray(1 To TotalRowsKompkresla + 1, 1 To 193)
    ReDim ProgruzkaKompkreslaTerminalArray(1 To TotalRowsKompkresla + 1, 1 To 193)
    ReDim ProgruzkaKompkreslaArray(1 To TotalRowsKompkresla + 1, 1 To 204)
    ProgruzkaKompkreslaArray = CreateHeader(ProgruzkaKompkreslaArray, "Б2С Компкресла")
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Компкресла")
    
    For i = UBound(KompkreslaArray, 2) To 0 Step -1
        If KompkreslaArray(0, i) = 0 Or KompkreslaArray(0, i) = vbNullString Then
            ReDim Preserve KompkreslaArray(3, UBound(KompkreslaArray, 2) - 1)
        End If
    Next

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Для перевода столбцов из СОТ-прогрузки (числовое обозначение столбцов в прогрузке) в Компкресла (SITE#, SITE#)
    Set TransferColumnsDict = CreateObject("Scripting.Dictionary")
    TransferColumnsDict.Add 6, "SITE38"
    TransferColumnsDict.Add 7, "SITE21"
    TransferColumnsDict.Add 8, "SITE24"
    TransferColumnsDict.Add 9, "SITE70"
    TransferColumnsDict.Add 10, "SITE71"
    TransferColumnsDict.Add 11, "SITE72"
    TransferColumnsDict.Add 12, "SITE73"
    TransferColumnsDict.Add 13, "SITE49"
    TransferColumnsDict.Add 14, "SITE74"
    TransferColumnsDict.Add 15, "SITE28"
    TransferColumnsDict.Add 16, "SITE75"
    TransferColumnsDict.Add 17, "SITE76"
    TransferColumnsDict.Add 18, "SITE17"
    TransferColumnsDict.Add 21, "SITE37"
    TransferColumnsDict.Add 22, "SITE33"
    TransferColumnsDict.Add 23, "SITE4"
    TransferColumnsDict.Add 24, "SITE50"
    TransferColumnsDict.Add 25, "SITE2"
    TransferColumnsDict.Add 26, "SITE77"
    TransferColumnsDict.Add 27, "SITE29"
    TransferColumnsDict.Add 28, "SITE32"
    TransferColumnsDict.Add 29, "SITE78"
    TransferColumnsDict.Add 30, "SITE79"
    TransferColumnsDict.Add 31, "SITE26"
    TransferColumnsDict.Add 32, "SITE36"
    TransferColumnsDict.Add 33, "SITE80"
    TransferColumnsDict.Add 34, "SITE81"
    TransferColumnsDict.Add 35, "SITE19"
    TransferColumnsDict.Add 36, "SITE20"
    TransferColumnsDict.Add 37, "SITE82"
    TransferColumnsDict.Add 38, "SITE83"
    TransferColumnsDict.Add 40, "SITE35"
    TransferColumnsDict.Add 41, "SITE43"
    TransferColumnsDict.Add 43, "SITE3"
    TransferColumnsDict.Add 44, "SITE39"
    TransferColumnsDict.Add 45, "SITE45"
    TransferColumnsDict.Add 46, "SITE51"
    TransferColumnsDict.Add 47, "SITE84"
    TransferColumnsDict.Add 48, "SITE52"
    TransferColumnsDict.Add 49, "SITE31"
    TransferColumnsDict.Add 50, "SITE44"
    TransferColumnsDict.Add 51, "SITE22"
    TransferColumnsDict.Add 52, "SITE53"
    TransferColumnsDict.Add 53, "SITE14"
    TransferColumnsDict.Add 56, "SITE85"
    TransferColumnsDict.Add 57, "SITE11"
    TransferColumnsDict.Add 58, "SITE54"
    TransferColumnsDict.Add 59, "SITE12"
    TransferColumnsDict.Add 60, "SITE86"
    TransferColumnsDict.Add 61, "SITE34"
    TransferColumnsDict.Add 62, "SITE5"
    TransferColumnsDict.Add 63, "SITE87"
    TransferColumnsDict.Add 65, "SITE55"
    TransferColumnsDict.Add 67, "SITE15"
    TransferColumnsDict.Add 69, "SITE88"
    TransferColumnsDict.Add 70, "SITE56"
    TransferColumnsDict.Add 71, "SITE16"
    TransferColumnsDict.Add 72, "SITE27"
    TransferColumnsDict.Add 73, "SITE57"
    TransferColumnsDict.Add 74, "SITE18"
    TransferColumnsDict.Add 75, "SITE58"
    TransferColumnsDict.Add 78, "SITE48"
    TransferColumnsDict.Add 79, "SITE59"
    TransferColumnsDict.Add 80, "SITE89"
    TransferColumnsDict.Add 81, "SITE41"
    TransferColumnsDict.Add 82, "SITE90"
    TransferColumnsDict.Add 83, "SITE60"
    TransferColumnsDict.Add 84, "SITE8"
    TransferColumnsDict.Add 85, "SITE61"
    TransferColumnsDict.Add 86, "SITE62"
    TransferColumnsDict.Add 87, "SITE63"
    TransferColumnsDict.Add 88, "SITE46"
    TransferColumnsDict.Add 89, "SITE23"
    TransferColumnsDict.Add 90, "SITE64"
    TransferColumnsDict.Add 91, "SITE9"
    TransferColumnsDict.Add 92, "SITE40"
    TransferColumnsDict.Add 93, "SITE13"
    TransferColumnsDict.Add 95, "SITE6"
    TransferColumnsDict.Add 97, "SITE30"
    TransferColumnsDict.Add 99, "SITE65"
    TransferColumnsDict.Add 100, "SITE10"
    TransferColumnsDict.Add 101, "SITE66"
    TransferColumnsDict.Add 102, "SITE67"
    TransferColumnsDict.Add 103, "SITE91"
    TransferColumnsDict.Add 105, "SITE68"
    TransferColumnsDict.Add 106, "SITE42"
    TransferColumnsDict.Add 107, "SITE69"
    TransferColumnsDict.Add 109, "SITE98"
    TransferColumnsDict.Add 110, "SITE99"
    TransferColumnsDict.Add 111, "SITE92"
    TransferColumnsDict.Add 112, "SITE100"
    TransferColumnsDict.Add 113, "SITE93"
    TransferColumnsDict.Add 114, "SITE101"
    TransferColumnsDict.Add 115, "SITE94"
    TransferColumnsDict.Add 116, "SITE95"
    TransferColumnsDict.Add 117, "SITE96"
    TransferColumnsDict.Add 118, "SITE102"
    TransferColumnsDict.Add 119, "SITE97"
    TransferColumnsDict.Add 120, "SITE106"
    TransferColumnsDict.Add 121, "SITE107"
    TransferColumnsDict.Add 122, "SITE108"
    TransferColumnsDict.Add 123, "SITE103"
    TransferColumnsDict.Add 124, "SITE104"
    TransferColumnsDict.Add 125, "SITE105"

    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

Kod1C_Change:
    
    NRowProschet = KompkreslaArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    a = Split(ProschetArray(NRowProschet, 7), ";")
    
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

    TypeZabor = FillInfocolumnTypeZabor("Н", ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, KompkreslaArray(3, d), 0, a)

    ProgruzkaKompkreslaArray(NRowProgruzka, 1) = Kod1CProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 193) = GorodOtpravki
    ProgruzkaKompkreslaArray(NRowProgruzka, 194) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, True)
    ProgruzkaKompkreslaArray(NRowProgruzka, 195) = TypeZabor
    ProgruzkaKompkreslaArray(NRowProgruzka, 196) = Volume
    ProgruzkaKompkreslaArray(NRowProgruzka, 197) = Weight
    ProgruzkaKompkreslaArray(NRowProgruzka, 198) = Quantity
    ProgruzkaKompkreslaArray(NRowProgruzka, 199) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaKompkreslaArray(NRowProgruzka, 200) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaKompkreslaArray(NRowProgruzka, 201) = OptPriceProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 202) = ProschetArray(NRowProschet, 1)
    ProgruzkaKompkreslaArray(NRowProgruzka, 203) = Artikul
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To 142

        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_DRAM, 0)
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 123 Then
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
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
            
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 123 Then
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
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
            
        End If

        If NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        Else
            Price = OptPriceProschet * (Margin)
        End If

        Tarif = fTarif(KompkreslaArray(2, d), CitiesArray(NCity), VolumetricWeight, IsKresla:=1)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity), IsKresla:=1)

        If NCity >= 123 Then
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

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps

        TotalPrice = Price + FullDostavkaPrice

        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
      
        If NCity < 123 Or NCity >= 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(TotalPrice + JU, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            Else
                TotalPrice = Round_Up(TotalPrice + JU + 200, 0)
            End If
        ElseIf NCity < 135 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_TENGE, -1)
            End If
        ElseIf NCity < 141 Then
            If OptPrice > 150000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, 0)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_BELRUB, 0)
            End If
        ElseIf NCity < 142 Then
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
        TotalPrice = TotalPrice / 0.96

zapis:
        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
      
        If NCity < 123 And TotalPrice < RetailPrice And RetailPrice > 7500 Then
            TotalPrice = Round_Up(RetailPrice, -1)
        ElseIf NCity < 123 And TotalPrice < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice, 0)
        ElseIf NCity >= 123 And NCity < 135 And TotalPrice / KURS_TENGE < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * KURS_TENGE, -1)
        ElseIf NCity >= 135 And NCity < 141 And TotalPrice / KURS_BELRUB < RetailPrice And RetailPrice > 150000 Then
            TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, -1)
        ElseIf NCity >= 135 And NCity < 141 And TotalPrice / KURS_BELRUB < RetailPrice And RetailPrice <= 150000 Then
            TotalPrice = Round_Up(RetailPrice * KURS_BELRUB, 0)
        ElseIf NCity >= 141 And NCity < 142 And TotalPrice / KURS_SOM < RetailPrice And RetailPrice > 5000 Then
            TotalPrice = Round_Up(RetailPrice * KURS_SOM, -1)
        ElseIf NCity >= 141 And NCity < 142 And TotalPrice / KURS_SOM < RetailPrice And RetailPrice <= 5000 Then
            TotalPrice = Round_Up(RetailPrice * KURS_SOM, 0)
        ElseIf NCity = 142 And TotalPrice / KURS_DRAM < RetailPrice Then
            TotalPrice = Round_Up(RetailPrice * KURS_DRAM, -1)
        End If

unrrc:
    
        ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127

            If NcolumnProgruzka < 108 And ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 7500 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
            ElseIf NcolumnProgruzka < 108 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 And ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 150000 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 And ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 5000 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If
        Next NcolumnProgruzka
        GoTo federal_orenburg
    End If

    
    Call fMixerK(ProgruzkaKompkreslaDverArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg

federal_orenburg:
Change_Row_Progruzka:

    Postavshik = ProschetArray(NRowProschet, 1)

    If ProgruzkaKompkreslaDverArray(NRowProgruzka, 5) = 0 Then
        GoTo Resetting
    End If

Resetting:
    Do While d <> UBound(KompkreslaArray, 2)
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
    NRowProgruzka = 2

Kod1C_Change1:
    
    NRowProschet = KompkreslaArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    a = Split(ProschetArray(NRowProschet, 7), ";")
    
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
    
    tps = fTPS(Weight, Quantity)

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ProschetArray(KompkreslaArray(2, d), 13) <> 0 Then
        ZaborPrice = 200
        LgotniiGorod = True
    Else
        ZaborPrice = ProschetArray(KompkreslaArray(2, d), 13)
        LgotniiGorod = False
    End If

    If d <> 0 And ProschetArray(KompkreslaArray(2, d), 1) = Postavshik Then
        GoTo Komponent_Change1
    End If

Komponent_Change1:
    
    NRowProschet = KompkreslaArray(2, d)
    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) = OptPriceProschet

    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka1
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka1
    End If

    Margin = fMargin(OptPriceProschet, Quantity, KompkreslaArray(3, d), 0, a)

    For NCity = 20 To 142
    
        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_DRAM, 0)
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 123 Then
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
            GoTo zapis1
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 123 Then
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
            GoTo zapis1
        End If
    
        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        Else
            Price = OptPriceProschet * (Margin)
        End If

        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))

        If NCity >= 123 Then
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

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps

        TotalPrice = Price + FullDostavkaPrice
    
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
    
        If NCity < 123 Or NCity >= 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            Else
                TotalPrice = Round_Up(TotalPrice + JU + 200, 0)
            End If
        ElseIf NCity < 135 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_TENGE, -1)
            End If
        ElseIf NCity < 141 Then
            If OptPrice > 150000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, 0)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_BELRUB, 0)
            End If
        ElseIf NCity < 142 Then
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

zapis1:
        If UserForm1.CheckBox2 = True Then
            GoTo unrrc1
        End If
    
        'Если в словаре есть ключ с номером столбца СОТ-прогрузки, тогда
        'Добавляем в прогрузку КК в столбец допинфы наименование столбца прогрузки КК (SITE#, SITE#), полную цену, межтерминалку и полные транспортные
        'После этого добавляем то же наименование столбца со значениями 0::0::0::DPD
        If TransferColumnsDict.exists(NcolumnProgruzka) = True Then
            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 129) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 129) & TransferColumnsDict.Item(NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
        End If
    
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If RetailPrice > 7500 Then
                    TotalPrice = Round_Up(RetailPrice, 0)
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

unrrc1:

        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg1
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127
            If NcolumnProgruzka < 108 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 5000 Then
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
                Else
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
                End If
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If
        Next NcolumnProgruzka

        GoTo federal_orenburg1

    End If

    
    Call fMixerEK(ProgruzkaKompkreslaTerminalArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg1

federal_orenburg1:

Change_Row_Progruzka1:

    Postavshik = ProschetArray(NRowProschet, 1)

    If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 5) = 0 Then
        GoTo Resetting1
    End If

Resetting1:
    Do While d <> UBound(KompkreslaArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change1
    Loop

    For i = 2 To NRowProgruzka
        ProgruzkaKompkreslaArray(i, 1) = ProgruzkaKompkreslaArray(i, 1)
        ProgruzkaKompkreslaArray(i, 2) = ProgruzkaKompkreslaTerminalArray(i, 6)
        ProgruzkaKompkreslaArray(i, 3) = 0
        ProgruzkaKompkreslaArray(i, 4) = ProgruzkaKompkreslaTerminalArray(i, 7)
        ProgruzkaKompkreslaArray(i, 5) = 0
        ProgruzkaKompkreslaArray(i, 6) = ProgruzkaKompkreslaTerminalArray(i, 8)
        ProgruzkaKompkreslaArray(i, 7) = 0
        ProgruzkaKompkreslaArray(i, 8) = ProgruzkaKompkreslaTerminalArray(i, 9)
        ProgruzkaKompkreslaArray(i, 9) = 0
        ProgruzkaKompkreslaArray(i, 10) = ProgruzkaKompkreslaTerminalArray(i, 10)
        ProgruzkaKompkreslaArray(i, 11) = 0
        ProgruzkaKompkreslaArray(i, 12) = ProgruzkaKompkreslaTerminalArray(i, 11)
        ProgruzkaKompkreslaArray(i, 13) = 0
        ProgruzkaKompkreslaArray(i, 14) = ProgruzkaKompkreslaTerminalArray(i, 12)
        ProgruzkaKompkreslaArray(i, 15) = 0
        ProgruzkaKompkreslaArray(i, 16) = ProgruzkaKompkreslaTerminalArray(i, 13)
        ProgruzkaKompkreslaArray(i, 17) = 0
        ProgruzkaKompkreslaArray(i, 18) = ProgruzkaKompkreslaTerminalArray(i, 14)
        ProgruzkaKompkreslaArray(i, 19) = 0
        ProgruzkaKompkreslaArray(i, 20) = ProgruzkaKompkreslaTerminalArray(i, 15)
        ProgruzkaKompkreslaArray(i, 21) = 0
        ProgruzkaKompkreslaArray(i, 22) = ProgruzkaKompkreslaTerminalArray(i, 16)
        ProgruzkaKompkreslaArray(i, 23) = 0
        ProgruzkaKompkreslaArray(i, 24) = ProgruzkaKompkreslaTerminalArray(i, 17)
        ProgruzkaKompkreslaArray(i, 25) = 0
        ProgruzkaKompkreslaArray(i, 26) = ProgruzkaKompkreslaTerminalArray(i, 18)
        ProgruzkaKompkreslaArray(i, 27) = 0
        ProgruzkaKompkreslaArray(i, 28) = ProgruzkaKompkreslaTerminalArray(i, 21)
        ProgruzkaKompkreslaArray(i, 29) = 0
        ProgruzkaKompkreslaArray(i, 30) = ProgruzkaKompkreslaTerminalArray(i, 22)
        ProgruzkaKompkreslaArray(i, 31) = 0
        ProgruzkaKompkreslaArray(i, 32) = ProgruzkaKompkreslaTerminalArray(i, 23)
        ProgruzkaKompkreslaArray(i, 33) = 0
        ProgruzkaKompkreslaArray(i, 34) = ProgruzkaKompkreslaTerminalArray(i, 24)
        ProgruzkaKompkreslaArray(i, 35) = 0
        ProgruzkaKompkreslaArray(i, 36) = ProgruzkaKompkreslaTerminalArray(i, 25)
        ProgruzkaKompkreslaArray(i, 37) = 0
        ProgruzkaKompkreslaArray(i, 38) = ProgruzkaKompkreslaTerminalArray(i, 26)
        ProgruzkaKompkreslaArray(i, 39) = 0
        ProgruzkaKompkreslaArray(i, 40) = ProgruzkaKompkreslaTerminalArray(i, 27)
        ProgruzkaKompkreslaArray(i, 41) = 0
        ProgruzkaKompkreslaArray(i, 42) = ProgruzkaKompkreslaTerminalArray(i, 28)
        ProgruzkaKompkreslaArray(i, 43) = 0
        ProgruzkaKompkreslaArray(i, 44) = ProgruzkaKompkreslaTerminalArray(i, 29)
        ProgruzkaKompkreslaArray(i, 45) = 0
        ProgruzkaKompkreslaArray(i, 46) = ProgruzkaKompkreslaTerminalArray(i, 30)
        ProgruzkaKompkreslaArray(i, 47) = 0
        ProgruzkaKompkreslaArray(i, 48) = ProgruzkaKompkreslaTerminalArray(i, 31)
        ProgruzkaKompkreslaArray(i, 49) = 0
        ProgruzkaKompkreslaArray(i, 50) = ProgruzkaKompkreslaTerminalArray(i, 32)
        ProgruzkaKompkreslaArray(i, 51) = 0
        ProgruzkaKompkreslaArray(i, 52) = ProgruzkaKompkreslaTerminalArray(i, 33)
        ProgruzkaKompkreslaArray(i, 53) = 0
        ProgruzkaKompkreslaArray(i, 54) = ProgruzkaKompkreslaTerminalArray(i, 34)
        ProgruzkaKompkreslaArray(i, 55) = 0
        ProgruzkaKompkreslaArray(i, 56) = ProgruzkaKompkreslaTerminalArray(i, 35)
        ProgruzkaKompkreslaArray(i, 57) = 0
        ProgruzkaKompkreslaArray(i, 58) = ProgruzkaKompkreslaTerminalArray(i, 36)
        ProgruzkaKompkreslaArray(i, 59) = 0
        ProgruzkaKompkreslaArray(i, 60) = ProgruzkaKompkreslaTerminalArray(i, 37)
        ProgruzkaKompkreslaArray(i, 61) = 0
        ProgruzkaKompkreslaArray(i, 62) = ProgruzkaKompkreslaTerminalArray(i, 38)
        ProgruzkaKompkreslaArray(i, 63) = 0
        ProgruzkaKompkreslaArray(i, 64) = ProgruzkaKompkreslaTerminalArray(i, 40)
        ProgruzkaKompkreslaArray(i, 65) = 0
        ProgruzkaKompkreslaArray(i, 66) = ProgruzkaKompkreslaTerminalArray(i, 41)
        ProgruzkaKompkreslaArray(i, 67) = 0
        ProgruzkaKompkreslaArray(i, 68) = ProgruzkaKompkreslaTerminalArray(i, 43)
        ProgruzkaKompkreslaArray(i, 69) = 0
        ProgruzkaKompkreslaArray(i, 70) = ProgruzkaKompkreslaTerminalArray(i, 44)
        ProgruzkaKompkreslaArray(i, 71) = 0
        ProgruzkaKompkreslaArray(i, 72) = ProgruzkaKompkreslaTerminalArray(i, 45)
        ProgruzkaKompkreslaArray(i, 73) = 0
        ProgruzkaKompkreslaArray(i, 74) = ProgruzkaKompkreslaTerminalArray(i, 46)
        ProgruzkaKompkreslaArray(i, 75) = 0
        ProgruzkaKompkreslaArray(i, 76) = ProgruzkaKompkreslaTerminalArray(i, 47)
        ProgruzkaKompkreslaArray(i, 77) = 0
        ProgruzkaKompkreslaArray(i, 78) = ProgruzkaKompkreslaTerminalArray(i, 48)
        ProgruzkaKompkreslaArray(i, 79) = 0
        ProgruzkaKompkreslaArray(i, 80) = ProgruzkaKompkreslaTerminalArray(i, 49)
        ProgruzkaKompkreslaArray(i, 81) = 0
        ProgruzkaKompkreslaArray(i, 82) = ProgruzkaKompkreslaTerminalArray(i, 50)
        ProgruzkaKompkreslaArray(i, 83) = 0
        ProgruzkaKompkreslaArray(i, 84) = ProgruzkaKompkreslaTerminalArray(i, 51)
        ProgruzkaKompkreslaArray(i, 85) = 0
        ProgruzkaKompkreslaArray(i, 86) = ProgruzkaKompkreslaTerminalArray(i, 52)
        ProgruzkaKompkreslaArray(i, 87) = 0
        ProgruzkaKompkreslaArray(i, 88) = ProgruzkaKompkreslaTerminalArray(i, 53)
        ProgruzkaKompkreslaArray(i, 89) = 0
        ProgruzkaKompkreslaArray(i, 90) = ProgruzkaKompkreslaTerminalArray(i, 56)
        ProgruzkaKompkreslaArray(i, 91) = 0
        ProgruzkaKompkreslaArray(i, 92) = ProgruzkaKompkreslaTerminalArray(i, 57)
        ProgruzkaKompkreslaArray(i, 93) = 0
        ProgruzkaKompkreslaArray(i, 94) = ProgruzkaKompkreslaTerminalArray(i, 58)
        ProgruzkaKompkreslaArray(i, 95) = 0
        ProgruzkaKompkreslaArray(i, 96) = ProgruzkaKompkreslaTerminalArray(i, 59)
        ProgruzkaKompkreslaArray(i, 97) = 0
        ProgruzkaKompkreslaArray(i, 98) = ProgruzkaKompkreslaTerminalArray(i, 60)
        ProgruzkaKompkreslaArray(i, 99) = 0
        ProgruzkaKompkreslaArray(i, 100) = ProgruzkaKompkreslaTerminalArray(i, 61)
        ProgruzkaKompkreslaArray(i, 101) = 0
        ProgruzkaKompkreslaArray(i, 102) = ProgruzkaKompkreslaTerminalArray(i, 62)
        ProgruzkaKompkreslaArray(i, 103) = 0
        ProgruzkaKompkreslaArray(i, 104) = ProgruzkaKompkreslaTerminalArray(i, 63)
        ProgruzkaKompkreslaArray(i, 105) = 0
        ProgruzkaKompkreslaArray(i, 106) = ProgruzkaKompkreslaTerminalArray(i, 65)
        ProgruzkaKompkreslaArray(i, 107) = 0
        ProgruzkaKompkreslaArray(i, 108) = ProgruzkaKompkreslaTerminalArray(i, 67)
        ProgruzkaKompkreslaArray(i, 109) = 0
        ProgruzkaKompkreslaArray(i, 110) = ProgruzkaKompkreslaTerminalArray(i, 69)
        ProgruzkaKompkreslaArray(i, 111) = 0
        ProgruzkaKompkreslaArray(i, 112) = ProgruzkaKompkreslaTerminalArray(i, 70)
        ProgruzkaKompkreslaArray(i, 113) = 0
        ProgruzkaKompkreslaArray(i, 114) = ProgruzkaKompkreslaTerminalArray(i, 71)
        ProgruzkaKompkreslaArray(i, 115) = 0
        ProgruzkaKompkreslaArray(i, 116) = ProgruzkaKompkreslaTerminalArray(i, 72)
        ProgruzkaKompkreslaArray(i, 117) = 0
        ProgruzkaKompkreslaArray(i, 118) = ProgruzkaKompkreslaTerminalArray(i, 73)
        ProgruzkaKompkreslaArray(i, 119) = 0
        ProgruzkaKompkreslaArray(i, 120) = ProgruzkaKompkreslaTerminalArray(i, 74)
        ProgruzkaKompkreslaArray(i, 121) = 0
        ProgruzkaKompkreslaArray(i, 122) = ProgruzkaKompkreslaTerminalArray(i, 75)
        ProgruzkaKompkreslaArray(i, 123) = 0
        ProgruzkaKompkreslaArray(i, 124) = ProgruzkaKompkreslaTerminalArray(i, 78)
        ProgruzkaKompkreslaArray(i, 125) = 0
        ProgruzkaKompkreslaArray(i, 126) = ProgruzkaKompkreslaTerminalArray(i, 79)
        ProgruzkaKompkreslaArray(i, 127) = 0
        ProgruzkaKompkreslaArray(i, 128) = ProgruzkaKompkreslaTerminalArray(i, 80)
        ProgruzkaKompkreslaArray(i, 129) = 0
        ProgruzkaKompkreslaArray(i, 130) = ProgruzkaKompkreslaTerminalArray(i, 81)
        ProgruzkaKompkreslaArray(i, 131) = 0
        ProgruzkaKompkreslaArray(i, 132) = ProgruzkaKompkreslaTerminalArray(i, 82)
        ProgruzkaKompkreslaArray(i, 133) = 0
        ProgruzkaKompkreslaArray(i, 134) = ProgruzkaKompkreslaTerminalArray(i, 83)
        ProgruzkaKompkreslaArray(i, 135) = 0
        ProgruzkaKompkreslaArray(i, 136) = ProgruzkaKompkreslaTerminalArray(i, 84)
        ProgruzkaKompkreslaArray(i, 137) = 0
        ProgruzkaKompkreslaArray(i, 138) = ProgruzkaKompkreslaTerminalArray(i, 85)
        ProgruzkaKompkreslaArray(i, 139) = 0
        ProgruzkaKompkreslaArray(i, 140) = ProgruzkaKompkreslaTerminalArray(i, 86)
        ProgruzkaKompkreslaArray(i, 141) = 0
        ProgruzkaKompkreslaArray(i, 142) = ProgruzkaKompkreslaTerminalArray(i, 87)
        ProgruzkaKompkreslaArray(i, 143) = 0
        ProgruzkaKompkreslaArray(i, 144) = ProgruzkaKompkreslaTerminalArray(i, 88)
        ProgruzkaKompkreslaArray(i, 145) = 0
        ProgruzkaKompkreslaArray(i, 146) = ProgruzkaKompkreslaTerminalArray(i, 89)
        ProgruzkaKompkreslaArray(i, 147) = 0
        ProgruzkaKompkreslaArray(i, 148) = ProgruzkaKompkreslaTerminalArray(i, 90)
        ProgruzkaKompkreslaArray(i, 149) = 0
        ProgruzkaKompkreslaArray(i, 150) = ProgruzkaKompkreslaTerminalArray(i, 91)
        ProgruzkaKompkreslaArray(i, 151) = 0
        ProgruzkaKompkreslaArray(i, 152) = ProgruzkaKompkreslaTerminalArray(i, 92)
        ProgruzkaKompkreslaArray(i, 153) = 0
        ProgruzkaKompkreslaArray(i, 154) = ProgruzkaKompkreslaTerminalArray(i, 93)
        ProgruzkaKompkreslaArray(i, 155) = 0
        ProgruzkaKompkreslaArray(i, 156) = ProgruzkaKompkreslaTerminalArray(i, 95)
        ProgruzkaKompkreslaArray(i, 157) = 0
        ProgruzkaKompkreslaArray(i, 158) = ProgruzkaKompkreslaTerminalArray(i, 97)
        ProgruzkaKompkreslaArray(i, 159) = 0
        ProgruzkaKompkreslaArray(i, 160) = ProgruzkaKompkreslaTerminalArray(i, 99)
        ProgruzkaKompkreslaArray(i, 161) = 0
        ProgruzkaKompkreslaArray(i, 162) = ProgruzkaKompkreslaTerminalArray(i, 100)
        ProgruzkaKompkreslaArray(i, 163) = 0
        ProgruzkaKompkreslaArray(i, 164) = ProgruzkaKompkreslaTerminalArray(i, 101)
        ProgruzkaKompkreslaArray(i, 165) = 0
        ProgruzkaKompkreslaArray(i, 166) = ProgruzkaKompkreslaTerminalArray(i, 102)
        ProgruzkaKompkreslaArray(i, 167) = 0
        ProgruzkaKompkreslaArray(i, 168) = ProgruzkaKompkreslaTerminalArray(i, 103)
        ProgruzkaKompkreslaArray(i, 169) = 0
        ProgruzkaKompkreslaArray(i, 170) = ProgruzkaKompkreslaTerminalArray(i, 105)
        ProgruzkaKompkreslaArray(i, 171) = 0
        ProgruzkaKompkreslaArray(i, 172) = ProgruzkaKompkreslaTerminalArray(i, 106)
        ProgruzkaKompkreslaArray(i, 173) = 0
        ProgruzkaKompkreslaArray(i, 174) = ProgruzkaKompkreslaTerminalArray(i, 107)
        ProgruzkaKompkreslaArray(i, 175) = 0
        ProgruzkaKompkreslaArray(i, 176) = ProgruzkaKompkreslaTerminalArray(i, 109)
        ProgruzkaKompkreslaArray(i, 177) = ProgruzkaKompkreslaTerminalArray(i, 110)
        ProgruzkaKompkreslaArray(i, 178) = ProgruzkaKompkreslaTerminalArray(i, 111)
        ProgruzkaKompkreslaArray(i, 179) = ProgruzkaKompkreslaTerminalArray(i, 112)
        ProgruzkaKompkreslaArray(i, 180) = ProgruzkaKompkreslaTerminalArray(i, 113)
        ProgruzkaKompkreslaArray(i, 181) = ProgruzkaKompkreslaTerminalArray(i, 114)
        ProgruzkaKompkreslaArray(i, 182) = ProgruzkaKompkreslaTerminalArray(i, 115)
        ProgruzkaKompkreslaArray(i, 183) = ProgruzkaKompkreslaTerminalArray(i, 116)
        ProgruzkaKompkreslaArray(i, 184) = ProgruzkaKompkreslaTerminalArray(i, 117)
        ProgruzkaKompkreslaArray(i, 185) = ProgruzkaKompkreslaTerminalArray(i, 118)
        ProgruzkaKompkreslaArray(i, 186) = ProgruzkaKompkreslaTerminalArray(i, 119)
        ProgruzkaKompkreslaArray(i, 187) = ProgruzkaKompkreslaTerminalArray(i, 120)
        ProgruzkaKompkreslaArray(i, 188) = ProgruzkaKompkreslaTerminalArray(i, 121)
        ProgruzkaKompkreslaArray(i, 189) = ProgruzkaKompkreslaTerminalArray(i, 122)
        ProgruzkaKompkreslaArray(i, 190) = ProgruzkaKompkreslaTerminalArray(i, 123)
        ProgruzkaKompkreslaArray(i, 191) = ProgruzkaKompkreslaTerminalArray(i, 124)
        ProgruzkaKompkreslaArray(i, 192) = ProgruzkaKompkreslaTerminalArray(i, 125)
        ProgruzkaKompkreslaArray(i, 204) = ProgruzkaKompkreslaTerminalArray(i, 129)
    Next

    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:GV" & TotalRowsKompkresla + 1).Value = ProgruzkaKompkreslaArray
    Application.ScreenUpdating = True

    If UserForm1.Autosave = True Then
        
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Компкресла")

    End If

ended:
    'Unload UserForm1

End Sub


'Прогрузка Н Промо
Public Sub Прогрузка_Промо_Компкресла(ByRef KompkreslaPromoArray, ByRef TotalRowsKompkreslaPromo)
  
    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NCity As Integer
    Dim r1 As Long
    Dim d As Integer
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a() As String
    Dim OptPrice
    Dim mvvodin() As String
    Dim ProgruzkaKompkreslaDverArray() As Variant
    Dim ProgruzkaKompkreslaTerminalArray() As Variant
    Dim ProgruzkaKompkreslaArray() As Variant
    Dim PlatniiVes
    Dim TypeZabor As String
    Dim CitiesArray() As String
    Dim MinTerminalTerminalPrice As Double
    Dim Transitive As Boolean
    Dim mdostavkadpd() As Variant
    Dim TransferColumnsDict As Object
    Dim NcolumnProgruzka As Long
    Dim NRowProgruzka As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Postavshik As String
    Dim Tarif As Double
    Dim i As Long
    Dim NRowProschet As Long
    Dim tps As Long
    
    Transitive = True

    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    ReDim ProgruzkaKompkreslaDverArray(1 To TotalRowsKompkreslaPromo + 1, 1 To 193)
    ReDim ProgruzkaKompkreslaTerminalArray(1 To TotalRowsKompkreslaPromo + 1, 1 To 193)
    ReDim ProgruzkaKompkreslaArray(1 To TotalRowsKompkreslaPromo + 1, 1 To 204)
    ProgruzkaKompkreslaArray = CreateHeader(ProgruzkaKompkreslaArray, "Б2С Компкресла")
    
    For i = UBound(KompkreslaPromoArray, 2) To 0 Step -1
        If KompkreslaPromoArray(0, i) = 0 Or KompkreslaPromoArray(0, i) = vbNullString Then
            ReDim Preserve KompkreslaPromoArray(3, UBound(KompkreslaPromoArray, 2) - 1)
        End If
    Next
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Компкресла")
    
    'Для перевода столбцов из СОТ-прогрузки (числовое обозначение столбцов в прогрузке) в Компкресла (SITE#, SITE#)
    Set TransferColumnsDict = CreateObject("Scripting.Dictionary")
    TransferColumnsDict.Add 6, "SITE38"
    TransferColumnsDict.Add 7, "SITE21"
    TransferColumnsDict.Add 8, "SITE24"
    TransferColumnsDict.Add 9, "SITE70"
    TransferColumnsDict.Add 10, "SITE71"
    TransferColumnsDict.Add 11, "SITE72"
    TransferColumnsDict.Add 12, "SITE73"
    TransferColumnsDict.Add 13, "SITE49"
    TransferColumnsDict.Add 14, "SITE74"
    TransferColumnsDict.Add 15, "SITE28"
    TransferColumnsDict.Add 16, "SITE75"
    TransferColumnsDict.Add 17, "SITE76"
    TransferColumnsDict.Add 18, "SITE17"
    TransferColumnsDict.Add 21, "SITE37"
    TransferColumnsDict.Add 22, "SITE33"
    TransferColumnsDict.Add 23, "SITE4"
    TransferColumnsDict.Add 24, "SITE50"
    TransferColumnsDict.Add 25, "SITE2"
    TransferColumnsDict.Add 26, "SITE77"
    TransferColumnsDict.Add 27, "SITE29"
    TransferColumnsDict.Add 28, "SITE32"
    TransferColumnsDict.Add 29, "SITE78"
    TransferColumnsDict.Add 30, "SITE79"
    TransferColumnsDict.Add 31, "SITE26"
    TransferColumnsDict.Add 32, "SITE36"
    TransferColumnsDict.Add 33, "SITE80"
    TransferColumnsDict.Add 34, "SITE81"
    TransferColumnsDict.Add 35, "SITE19"
    TransferColumnsDict.Add 36, "SITE20"
    TransferColumnsDict.Add 37, "SITE82"
    TransferColumnsDict.Add 38, "SITE83"
    TransferColumnsDict.Add 40, "SITE35"
    TransferColumnsDict.Add 41, "SITE43"
    TransferColumnsDict.Add 43, "SITE3"
    TransferColumnsDict.Add 44, "SITE39"
    TransferColumnsDict.Add 45, "SITE45"
    TransferColumnsDict.Add 46, "SITE51"
    TransferColumnsDict.Add 47, "SITE84"
    TransferColumnsDict.Add 48, "SITE52"
    TransferColumnsDict.Add 49, "SITE31"
    TransferColumnsDict.Add 50, "SITE44"
    TransferColumnsDict.Add 51, "SITE22"
    TransferColumnsDict.Add 52, "SITE53"
    TransferColumnsDict.Add 53, "SITE14"
    TransferColumnsDict.Add 56, "SITE85"
    TransferColumnsDict.Add 57, "SITE11"
    TransferColumnsDict.Add 58, "SITE54"
    TransferColumnsDict.Add 59, "SITE12"
    TransferColumnsDict.Add 60, "SITE86"
    TransferColumnsDict.Add 61, "SITE34"
    TransferColumnsDict.Add 62, "SITE5"
    TransferColumnsDict.Add 63, "SITE87"
    TransferColumnsDict.Add 65, "SITE55"
    TransferColumnsDict.Add 67, "SITE15"
    TransferColumnsDict.Add 69, "SITE88"
    TransferColumnsDict.Add 70, "SITE56"
    TransferColumnsDict.Add 71, "SITE16"
    TransferColumnsDict.Add 72, "SITE27"
    TransferColumnsDict.Add 73, "SITE57"
    TransferColumnsDict.Add 74, "SITE18"
    TransferColumnsDict.Add 75, "SITE58"
    TransferColumnsDict.Add 78, "SITE48"
    TransferColumnsDict.Add 79, "SITE59"
    TransferColumnsDict.Add 80, "SITE89"
    TransferColumnsDict.Add 81, "SITE41"
    TransferColumnsDict.Add 82, "SITE90"
    TransferColumnsDict.Add 83, "SITE60"
    TransferColumnsDict.Add 84, "SITE8"
    TransferColumnsDict.Add 85, "SITE61"
    TransferColumnsDict.Add 86, "SITE62"
    TransferColumnsDict.Add 87, "SITE63"
    TransferColumnsDict.Add 88, "SITE46"
    TransferColumnsDict.Add 89, "SITE23"
    TransferColumnsDict.Add 90, "SITE64"
    TransferColumnsDict.Add 91, "SITE9"
    TransferColumnsDict.Add 92, "SITE40"
    TransferColumnsDict.Add 93, "SITE13"
    TransferColumnsDict.Add 95, "SITE6"
    TransferColumnsDict.Add 97, "SITE30"
    TransferColumnsDict.Add 99, "SITE65"
    TransferColumnsDict.Add 100, "SITE10"
    TransferColumnsDict.Add 101, "SITE66"
    TransferColumnsDict.Add 102, "SITE67"
    TransferColumnsDict.Add 103, "SITE91"
    TransferColumnsDict.Add 105, "SITE68"
    TransferColumnsDict.Add 106, "SITE42"
    TransferColumnsDict.Add 107, "SITE69"
    TransferColumnsDict.Add 109, "SITE98"
    TransferColumnsDict.Add 110, "SITE99"
    TransferColumnsDict.Add 111, "SITE92"
    TransferColumnsDict.Add 112, "SITE100"
    TransferColumnsDict.Add 113, "SITE93"
    TransferColumnsDict.Add 114, "SITE101"
    TransferColumnsDict.Add 115, "SITE94"
    TransferColumnsDict.Add 116, "SITE95"
    TransferColumnsDict.Add 117, "SITE96"
    TransferColumnsDict.Add 118, "SITE102"
    TransferColumnsDict.Add 119, "SITE97"
    TransferColumnsDict.Add 120, "SITE106"
    TransferColumnsDict.Add 121, "SITE107"
    TransferColumnsDict.Add 122, "SITE108"
    TransferColumnsDict.Add 123, "SITE103"
    TransferColumnsDict.Add 124, "SITE104"
    TransferColumnsDict.Add 125, "SITE105"

    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

Kod1C_Change:
    
    NRowProschet = KompkreslaPromoArray(2, d)
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
    
    ReDim mvvodin(UBound(a))
    
    For i = 0 To UBound(a)
        If a(i) = vbNullString Then
            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
        Else
            mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
        End If
    Next
    
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

    TypeZabor = FillInfocolumnTypeZabor("Н", ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps, 1)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, 0, 1, a)
    
    ProgruzkaKompkreslaArray(NRowProgruzka, 1) = Kod1CProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 193) = GorodOtpravki
    ProgruzkaKompkreslaArray(NRowProgruzka, 194) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, True)
    ProgruzkaKompkreslaArray(NRowProgruzka, 195) = TypeZabor
    ProgruzkaKompkreslaArray(NRowProgruzka, 196) = Volume
    ProgruzkaKompkreslaArray(NRowProgruzka, 197) = Weight
    ProgruzkaKompkreslaArray(NRowProgruzka, 198) = Quantity
    ProgruzkaKompkreslaArray(NRowProgruzka, 199) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaKompkreslaArray(NRowProgruzka, 200) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaKompkreslaArray(NRowProgruzka, 201) = OptPriceProschet
    ProgruzkaKompkreslaArray(NRowProgruzka, 202) = ProschetArray(NRowProschet, 1)
    ProgruzkaKompkreslaArray(NRowProgruzka, 203) = Artikul
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To 142
        
        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_DRAM, 0)
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 2 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 1.7 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        
        End If

        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity < 123 Then
            Price = OptPriceProschet * (Margin)
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        End If

        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight, IsKresla:=1)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity), IsKresla:=1)

        If NCity >= 123 Then
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

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps

        TotalPrice = Price + FullDostavkaPrice
        
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
        
        If NCity < 123 Or NCity >= 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(TotalPrice + JU, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            Else
                TotalPrice = Round_Up(TotalPrice + JU + 200, 0)
            End If
        ElseIf NCity < 135 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_TENGE, -1)
            End If
        ElseIf NCity < 141 Then
            If OptPrice > 150000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up((TotalPrice + JU) * KURS_BELRUB, 0)
            Else
                TotalPrice = Round_Up((TotalPrice + JU + 200) * KURS_BELRUB, 0)
            End If
        ElseIf NCity < 142 Then
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
        TotalPrice = TotalPrice / 0.96

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
        
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If RetailPrice > 7500 Then
                    TotalPrice = Round_Up(RetailPrice, -1)
                Else
                    TotalPrice = Round_Up(RetailPrice, 0)
                End If
            End If
        ElseIf NCity >= 123 And NcolumnProgruzka < 135 Then
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
                Else
                    TotalPrice = Round_Up(RetailPrice * KURS_SOM, 0)
                End If
            End If
        ElseIf NCity = 142 Then
            If (TotalPrice / KURS_DRAM) < RetailPrice Then
                TotalPrice = Round_Up(RetailPrice * KURS_DRAM, -1)
            End If
        End If
unrrc:

zapis:
        ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127
            
            If NcolumnProgruzka < 108 Then
                If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 7500 Then
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                Else
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            ElseIf NcolumnProgruzka = 126 Then
                If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 5000 Then
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
                Else
                    ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
                End If
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaDverArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If
        Next NcolumnProgruzka

        GoTo federal_orenburg

    End If

    
    Call fMixerK(ProgruzkaKompkreslaDverArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg
    
federal_orenburg:

Change_Row_Progruzka:

    Postavshik = ProschetArray(KompkreslaPromoArray(2, d), 1)

    Do While d <> UBound(KompkreslaPromoArray, 2)
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
    NRowProgruzka = 2

Kod1C_Change1:
    
    NRowProschet = KompkreslaPromoArray(2, d)
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

    ReDim mvvodin(UBound(a))

    For i = 0 To UBound(a)
        If a(i) = vbNullString Then
            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
        Else
            mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
        End If
    Next

    If (UBound(mvvodin) + 1) Mod 3 <> 0 Then
        MsgBox "Некорректная наценка/категория"
        GoTo ended
    End If
    
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
        ZaborPrice = 0
        LgotniiGorod = True
    Else
        ZaborPrice = ProschetArray(NRowProschet, 13)
        LgotniiGorod = False
    End If
    
    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 1) = Kod1CProschet
    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) = OptPriceProschet
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka1
    End If
    
    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka1
    End If

    Margin = fMargin(OptPriceProschet, Quantity, 0, 1, a)

    For NCity = 20 To 142
        
        'Расчет цены для комплектующих
        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin, 0)
            ElseIf NCity < 135 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_TENGE, 0)
            ElseIf NCity < 141 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_BELRUB, 0)
            ElseIf NCity < 142 Then
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(OptPriceProschet * Margin * KURS_DRAM, 0)
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 2 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis1
            
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * 1.7 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis1
        End If
        
        'Расчет опта на наценку
        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity < 123 Then
            Price = OptPriceProschet * Margin
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        End If
        
        'Расчет транспортных:
        Tarif = fTarif(NRowProschet, CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(NRowProschet, CitiesArray(NCity))

        If NCity >= 123 Then
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

        FullDostavkaPrice = fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, GorodOtpravki, JU, _
        TypeZaborProschet, Quantity, RazgruzkaPrice, ZaborPriceProschet) + tps
    
        'Итого с транспортными:
        TotalPrice = Price + FullDostavkaPrice
            
        If Quantity > 1 Then
            OptPrice = OptPriceProschet * Quantity
        Else
            OptPrice = OptPriceProschet
        End If
        
        'Округление цены и умножение на СНГ валюту
        If NCity < 123 Or NCity >= 143 Then
            If OptPrice > 7500 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(TotalPrice + JU, 0)
            Else
                TotalPrice = Round_Up(TotalPrice + JU + 200, 0)
            End If
        ElseIf NCity < 135 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_TENGE, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200) / 0.96) * KURS_TENGE, -1)
            End If
        ElseIf NCity < 141 Then
            If OptPrice > 150000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_BELRUB, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_BELRUB, 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200) / 0.96) * KURS_BELRUB, 0)
            End If
        ElseIf NCity < 142 Then
            If OptPrice > 5000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_SOM, -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_SOM, 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200) / 0.96) * KURS_SOM, 0)
            End If
        ElseIf NCity = 142 Then
            If OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96) * KURS_DRAM, -1)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200) / 0.96) * KURS_DRAM, -1)
            End If
        End If

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc1
        End If
        
        'Приведение к РРЦ, если цена с транспортными < РРЦ
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If RetailPrice > 7500 Then
                    TotalPrice = Round_Up(RetailPrice, 0)
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

unrrc1:
zapis1:
        'Если в словаре есть ключ с номером столбца СОТ-прогрузки, тогда
        'Добавляем в прогрузку КК в столбец допинфы наименование столбца прогрузки КК (SITE#, SITE#), полную цену, межтерминалку и полные транспортные
        'После этого добавляем то же наименование столбца со значениями 0::0::0::DPD
        If TransferColumnsDict.exists(NcolumnProgruzka) = True Then
            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 129) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 129) & TransferColumnsDict.Item(NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
        End If

        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg1
    End If
    
    'Уникализация для комплектующих
    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127
            
            If NcolumnProgruzka < 108 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
            ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 5000 Then
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 126)
                Else
                    ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 126)
                End If
            ElseIf NcolumnProgruzka = 127 Then
                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 127)
            End If
        Next NcolumnProgruzka

        GoTo federal_orenburg1
    End If

    
    Call fMixerEK(ProgruzkaKompkreslaTerminalArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg1

    For NcolumnProgruzka = 5 To 127
        TotalPrice = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka)
        NCity = NcolumnProgruzka
        For i = 5 To 127
            If TotalPrice = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, i) Then
                If NCity <> i Then
                    If NcolumnProgruzka < 108 Then
                        If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 7500 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1
                            i = 4
                            TotalPrice = TotalPrice + 1
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1
                            i = 4
                            TotalPrice = TotalPrice + 1
                        End If
                    ElseIf NcolumnProgruzka >= 108 And NCity < 120 Then
                        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10
                        i = 4
                        TotalPrice = TotalPrice + 10
                    ElseIf NcolumnProgruzka >= 120 And NCity < 126 Then
                        If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10
                            i = 4
                            TotalPrice = TotalPrice + 10
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1
                            i = 4
                            TotalPrice = TotalPrice + 1
                        End If
                    ElseIf NcolumnProgruzka >= 126 And NCity < 127 Then
                        If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 5000 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10
                            i = 4
                            TotalPrice = TotalPrice + 10
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 1
                            i = 4
                            TotalPrice = TotalPrice + 1
                        End If
                    ElseIf NcolumnProgruzka = 127 Then
                        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, NcolumnProgruzka) + 10
                        i = 4
                        TotalPrice = TotalPrice + 10
                    End If
                End If
            End If
        Next i
    Next NcolumnProgruzka

federal_orenburg1:

Change_Row_Progruzka1:

    Postavshik = ProschetArray(NRowProschet, 1)

    Do While d <> UBound(KompkreslaPromoArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change1
    Loop

    For i = 2 To NRowProgruzka
        ProgruzkaKompkreslaArray(i, 1) = ProgruzkaKompkreslaArray(i, 1)
        ProgruzkaKompkreslaArray(i, 2) = ProgruzkaKompkreslaTerminalArray(i, 6)
        ProgruzkaKompkreslaArray(i, 3) = 0
        ProgruzkaKompkreslaArray(i, 4) = ProgruzkaKompkreslaTerminalArray(i, 7)
        ProgruzkaKompkreslaArray(i, 5) = 0
        ProgruzkaKompkreslaArray(i, 6) = ProgruzkaKompkreslaTerminalArray(i, 8)
        ProgruzkaKompkreslaArray(i, 7) = 0
        ProgruzkaKompkreslaArray(i, 8) = ProgruzkaKompkreslaTerminalArray(i, 9)
        ProgruzkaKompkreslaArray(i, 9) = 0
        ProgruzkaKompkreslaArray(i, 10) = ProgruzkaKompkreslaTerminalArray(i, 10)
        ProgruzkaKompkreslaArray(i, 11) = 0
        ProgruzkaKompkreslaArray(i, 12) = ProgruzkaKompkreslaTerminalArray(i, 11)
        ProgruzkaKompkreslaArray(i, 13) = 0
        ProgruzkaKompkreslaArray(i, 14) = ProgruzkaKompkreslaTerminalArray(i, 12)
        ProgruzkaKompkreslaArray(i, 15) = 0
        ProgruzkaKompkreslaArray(i, 16) = ProgruzkaKompkreslaTerminalArray(i, 13)
        ProgruzkaKompkreslaArray(i, 17) = 0
        ProgruzkaKompkreslaArray(i, 18) = ProgruzkaKompkreslaTerminalArray(i, 14)
        ProgruzkaKompkreslaArray(i, 19) = 0
        ProgruzkaKompkreslaArray(i, 20) = ProgruzkaKompkreslaTerminalArray(i, 15)
        ProgruzkaKompkreslaArray(i, 21) = 0
        ProgruzkaKompkreslaArray(i, 22) = ProgruzkaKompkreslaTerminalArray(i, 16)
        ProgruzkaKompkreslaArray(i, 23) = 0
        ProgruzkaKompkreslaArray(i, 24) = ProgruzkaKompkreslaTerminalArray(i, 17)
        ProgruzkaKompkreslaArray(i, 25) = 0
        ProgruzkaKompkreslaArray(i, 26) = ProgruzkaKompkreslaTerminalArray(i, 18)
        ProgruzkaKompkreslaArray(i, 27) = 0
        ProgruzkaKompkreslaArray(i, 28) = ProgruzkaKompkreslaTerminalArray(i, 21)
        ProgruzkaKompkreslaArray(i, 29) = 0
        ProgruzkaKompkreslaArray(i, 30) = ProgruzkaKompkreslaTerminalArray(i, 22)
        ProgruzkaKompkreslaArray(i, 31) = 0
        ProgruzkaKompkreslaArray(i, 32) = ProgruzkaKompkreslaTerminalArray(i, 23)
        ProgruzkaKompkreslaArray(i, 33) = 0
        ProgruzkaKompkreslaArray(i, 34) = ProgruzkaKompkreslaTerminalArray(i, 24)
        ProgruzkaKompkreslaArray(i, 35) = 0
        ProgruzkaKompkreslaArray(i, 36) = ProgruzkaKompkreslaTerminalArray(i, 25)
        ProgruzkaKompkreslaArray(i, 37) = 0
        ProgruzkaKompkreslaArray(i, 38) = ProgruzkaKompkreslaTerminalArray(i, 26)
        ProgruzkaKompkreslaArray(i, 39) = 0
        ProgruzkaKompkreslaArray(i, 40) = ProgruzkaKompkreslaTerminalArray(i, 27)
        ProgruzkaKompkreslaArray(i, 41) = 0
        ProgruzkaKompkreslaArray(i, 42) = ProgruzkaKompkreslaTerminalArray(i, 28)
        ProgruzkaKompkreslaArray(i, 43) = 0
        ProgruzkaKompkreslaArray(i, 44) = ProgruzkaKompkreslaTerminalArray(i, 29)
        ProgruzkaKompkreslaArray(i, 45) = 0
        ProgruzkaKompkreslaArray(i, 46) = ProgruzkaKompkreslaTerminalArray(i, 30)
        ProgruzkaKompkreslaArray(i, 47) = 0
        ProgruzkaKompkreslaArray(i, 48) = ProgruzkaKompkreslaTerminalArray(i, 31)
        ProgruzkaKompkreslaArray(i, 49) = 0
        ProgruzkaKompkreslaArray(i, 50) = ProgruzkaKompkreslaTerminalArray(i, 32)
        ProgruzkaKompkreslaArray(i, 51) = 0
        ProgruzkaKompkreslaArray(i, 52) = ProgruzkaKompkreslaTerminalArray(i, 33)
        ProgruzkaKompkreslaArray(i, 53) = 0
        ProgruzkaKompkreslaArray(i, 54) = ProgruzkaKompkreslaTerminalArray(i, 34)
        ProgruzkaKompkreslaArray(i, 55) = 0
        ProgruzkaKompkreslaArray(i, 56) = ProgruzkaKompkreslaTerminalArray(i, 35)
        ProgruzkaKompkreslaArray(i, 57) = 0
        ProgruzkaKompkreslaArray(i, 58) = ProgruzkaKompkreslaTerminalArray(i, 36)
        ProgruzkaKompkreslaArray(i, 59) = 0
        ProgruzkaKompkreslaArray(i, 60) = ProgruzkaKompkreslaTerminalArray(i, 37)
        ProgruzkaKompkreslaArray(i, 61) = 0
        ProgruzkaKompkreslaArray(i, 62) = ProgruzkaKompkreslaTerminalArray(i, 38)
        ProgruzkaKompkreslaArray(i, 63) = 0
        ProgruzkaKompkreslaArray(i, 64) = ProgruzkaKompkreslaTerminalArray(i, 40)
        ProgruzkaKompkreslaArray(i, 65) = 0
        ProgruzkaKompkreslaArray(i, 66) = ProgruzkaKompkreslaTerminalArray(i, 41)
        ProgruzkaKompkreslaArray(i, 67) = 0
        ProgruzkaKompkreslaArray(i, 68) = ProgruzkaKompkreslaTerminalArray(i, 43)
        ProgruzkaKompkreslaArray(i, 69) = 0
        ProgruzkaKompkreslaArray(i, 70) = ProgruzkaKompkreslaTerminalArray(i, 44)
        ProgruzkaKompkreslaArray(i, 71) = 0
        ProgruzkaKompkreslaArray(i, 72) = ProgruzkaKompkreslaTerminalArray(i, 45)
        ProgruzkaKompkreslaArray(i, 73) = 0
        ProgruzkaKompkreslaArray(i, 74) = ProgruzkaKompkreslaTerminalArray(i, 46)
        ProgruzkaKompkreslaArray(i, 75) = 0
        ProgruzkaKompkreslaArray(i, 76) = ProgruzkaKompkreslaTerminalArray(i, 47)
        ProgruzkaKompkreslaArray(i, 77) = 0
        ProgruzkaKompkreslaArray(i, 78) = ProgruzkaKompkreslaTerminalArray(i, 48)
        ProgruzkaKompkreslaArray(i, 79) = 0
        ProgruzkaKompkreslaArray(i, 80) = ProgruzkaKompkreslaTerminalArray(i, 49)
        ProgruzkaKompkreslaArray(i, 81) = 0
        ProgruzkaKompkreslaArray(i, 82) = ProgruzkaKompkreslaTerminalArray(i, 50)
        ProgruzkaKompkreslaArray(i, 83) = 0
        ProgruzkaKompkreslaArray(i, 84) = ProgruzkaKompkreslaTerminalArray(i, 51)
        ProgruzkaKompkreslaArray(i, 85) = 0
        ProgruzkaKompkreslaArray(i, 86) = ProgruzkaKompkreslaTerminalArray(i, 52)
        ProgruzkaKompkreslaArray(i, 87) = 0
        ProgruzkaKompkreslaArray(i, 88) = ProgruzkaKompkreslaTerminalArray(i, 53)
        ProgruzkaKompkreslaArray(i, 89) = 0
        ProgruzkaKompkreslaArray(i, 90) = ProgruzkaKompkreslaTerminalArray(i, 56)
        ProgruzkaKompkreslaArray(i, 91) = 0
        ProgruzkaKompkreslaArray(i, 92) = ProgruzkaKompkreslaTerminalArray(i, 57)
        ProgruzkaKompkreslaArray(i, 93) = 0
        ProgruzkaKompkreslaArray(i, 94) = ProgruzkaKompkreslaTerminalArray(i, 58)
        ProgruzkaKompkreslaArray(i, 95) = 0
        ProgruzkaKompkreslaArray(i, 96) = ProgruzkaKompkreslaTerminalArray(i, 59)
        ProgruzkaKompkreslaArray(i, 97) = 0
        ProgruzkaKompkreslaArray(i, 98) = ProgruzkaKompkreslaTerminalArray(i, 60)
        ProgruzkaKompkreslaArray(i, 99) = 0
        ProgruzkaKompkreslaArray(i, 100) = ProgruzkaKompkreslaTerminalArray(i, 61)
        ProgruzkaKompkreslaArray(i, 101) = 0
        ProgruzkaKompkreslaArray(i, 102) = ProgruzkaKompkreslaTerminalArray(i, 62)
        ProgruzkaKompkreslaArray(i, 103) = 0
        ProgruzkaKompkreslaArray(i, 104) = ProgruzkaKompkreslaTerminalArray(i, 63)
        ProgruzkaKompkreslaArray(i, 105) = 0
        ProgruzkaKompkreslaArray(i, 106) = ProgruzkaKompkreslaTerminalArray(i, 65)
        ProgruzkaKompkreslaArray(i, 107) = 0
        ProgruzkaKompkreslaArray(i, 108) = ProgruzkaKompkreslaTerminalArray(i, 67)
        ProgruzkaKompkreslaArray(i, 109) = 0
        ProgruzkaKompkreslaArray(i, 110) = ProgruzkaKompkreslaTerminalArray(i, 69)
        ProgruzkaKompkreslaArray(i, 111) = 0
        ProgruzkaKompkreslaArray(i, 112) = ProgruzkaKompkreslaTerminalArray(i, 70)
        ProgruzkaKompkreslaArray(i, 113) = 0
        ProgruzkaKompkreslaArray(i, 114) = ProgruzkaKompkreslaTerminalArray(i, 71)
        ProgruzkaKompkreslaArray(i, 115) = 0
        ProgruzkaKompkreslaArray(i, 116) = ProgruzkaKompkreslaTerminalArray(i, 72)
        ProgruzkaKompkreslaArray(i, 117) = 0
        ProgruzkaKompkreslaArray(i, 118) = ProgruzkaKompkreslaTerminalArray(i, 73)
        ProgruzkaKompkreslaArray(i, 119) = 0
        ProgruzkaKompkreslaArray(i, 120) = ProgruzkaKompkreslaTerminalArray(i, 74)
        ProgruzkaKompkreslaArray(i, 121) = 0
        ProgruzkaKompkreslaArray(i, 122) = ProgruzkaKompkreslaTerminalArray(i, 75)
        ProgruzkaKompkreslaArray(i, 123) = 0
        ProgruzkaKompkreslaArray(i, 124) = ProgruzkaKompkreslaTerminalArray(i, 78)
        ProgruzkaKompkreslaArray(i, 125) = 0
        ProgruzkaKompkreslaArray(i, 126) = ProgruzkaKompkreslaTerminalArray(i, 79)
        ProgruzkaKompkreslaArray(i, 127) = 0
        ProgruzkaKompkreslaArray(i, 128) = ProgruzkaKompkreslaTerminalArray(i, 80)
        ProgruzkaKompkreslaArray(i, 129) = 0
        ProgruzkaKompkreslaArray(i, 130) = ProgruzkaKompkreslaTerminalArray(i, 81)
        ProgruzkaKompkreslaArray(i, 131) = 0
        ProgruzkaKompkreslaArray(i, 132) = ProgruzkaKompkreslaTerminalArray(i, 82)
        ProgruzkaKompkreslaArray(i, 133) = 0
        ProgruzkaKompkreslaArray(i, 134) = ProgruzkaKompkreslaTerminalArray(i, 83)
        ProgruzkaKompkreslaArray(i, 135) = 0
        ProgruzkaKompkreslaArray(i, 136) = ProgruzkaKompkreslaTerminalArray(i, 84)
        ProgruzkaKompkreslaArray(i, 137) = 0
        ProgruzkaKompkreslaArray(i, 138) = ProgruzkaKompkreslaTerminalArray(i, 85)
        ProgruzkaKompkreslaArray(i, 139) = 0
        ProgruzkaKompkreslaArray(i, 140) = ProgruzkaKompkreslaTerminalArray(i, 86)
        ProgruzkaKompkreslaArray(i, 141) = 0
        ProgruzkaKompkreslaArray(i, 142) = ProgruzkaKompkreslaTerminalArray(i, 87)
        ProgruzkaKompkreslaArray(i, 143) = 0
        ProgruzkaKompkreslaArray(i, 144) = ProgruzkaKompkreslaTerminalArray(i, 88)
        ProgruzkaKompkreslaArray(i, 145) = 0
        ProgruzkaKompkreslaArray(i, 146) = ProgruzkaKompkreslaTerminalArray(i, 89)
        ProgruzkaKompkreslaArray(i, 147) = 0
        ProgruzkaKompkreslaArray(i, 148) = ProgruzkaKompkreslaTerminalArray(i, 90)
        ProgruzkaKompkreslaArray(i, 149) = 0
        ProgruzkaKompkreslaArray(i, 150) = ProgruzkaKompkreslaTerminalArray(i, 91)
        ProgruzkaKompkreslaArray(i, 151) = 0
        ProgruzkaKompkreslaArray(i, 152) = ProgruzkaKompkreslaTerminalArray(i, 92)
        ProgruzkaKompkreslaArray(i, 153) = 0
        ProgruzkaKompkreslaArray(i, 154) = ProgruzkaKompkreslaTerminalArray(i, 93)
        ProgruzkaKompkreslaArray(i, 155) = 0
        ProgruzkaKompkreslaArray(i, 156) = ProgruzkaKompkreslaTerminalArray(i, 95)
        ProgruzkaKompkreslaArray(i, 157) = 0
        ProgruzkaKompkreslaArray(i, 158) = ProgruzkaKompkreslaTerminalArray(i, 97)
        ProgruzkaKompkreslaArray(i, 159) = 0
        ProgruzkaKompkreslaArray(i, 160) = ProgruzkaKompkreslaTerminalArray(i, 99)
        ProgruzkaKompkreslaArray(i, 161) = 0
        ProgruzkaKompkreslaArray(i, 162) = ProgruzkaKompkreslaTerminalArray(i, 100)
        ProgruzkaKompkreslaArray(i, 163) = 0
        ProgruzkaKompkreslaArray(i, 164) = ProgruzkaKompkreslaTerminalArray(i, 101)
        ProgruzkaKompkreslaArray(i, 165) = 0
        ProgruzkaKompkreslaArray(i, 166) = ProgruzkaKompkreslaTerminalArray(i, 102)
        ProgruzkaKompkreslaArray(i, 167) = 0
        ProgruzkaKompkreslaArray(i, 168) = ProgruzkaKompkreslaTerminalArray(i, 103)
        ProgruzkaKompkreslaArray(i, 169) = 0
        ProgruzkaKompkreslaArray(i, 170) = ProgruzkaKompkreslaTerminalArray(i, 105)
        ProgruzkaKompkreslaArray(i, 171) = 0
        ProgruzkaKompkreslaArray(i, 172) = ProgruzkaKompkreslaTerminalArray(i, 106)
        ProgruzkaKompkreslaArray(i, 173) = 0
        ProgruzkaKompkreslaArray(i, 174) = ProgruzkaKompkreslaTerminalArray(i, 107)
        ProgruzkaKompkreslaArray(i, 175) = 0
        ProgruzkaKompkreslaArray(i, 176) = ProgruzkaKompkreslaTerminalArray(i, 109)
        ProgruzkaKompkreslaArray(i, 177) = ProgruzkaKompkreslaTerminalArray(i, 110)
        ProgruzkaKompkreslaArray(i, 178) = ProgruzkaKompkreslaTerminalArray(i, 111)
        ProgruzkaKompkreslaArray(i, 179) = ProgruzkaKompkreslaTerminalArray(i, 112)
        ProgruzkaKompkreslaArray(i, 180) = ProgruzkaKompkreslaTerminalArray(i, 113)
        ProgruzkaKompkreslaArray(i, 181) = ProgruzkaKompkreslaTerminalArray(i, 114)
        ProgruzkaKompkreslaArray(i, 182) = ProgruzkaKompkreslaTerminalArray(i, 115)
        ProgruzkaKompkreslaArray(i, 183) = ProgruzkaKompkreslaTerminalArray(i, 116)
        ProgruzkaKompkreslaArray(i, 184) = ProgruzkaKompkreslaTerminalArray(i, 117)
        ProgruzkaKompkreslaArray(i, 185) = ProgruzkaKompkreslaTerminalArray(i, 118)
        ProgruzkaKompkreslaArray(i, 186) = ProgruzkaKompkreslaTerminalArray(i, 119)
        ProgruzkaKompkreslaArray(i, 187) = ProgruzkaKompkreslaTerminalArray(i, 120)
        ProgruzkaKompkreslaArray(i, 188) = ProgruzkaKompkreslaTerminalArray(i, 121)
        ProgruzkaKompkreslaArray(i, 189) = ProgruzkaKompkreslaTerminalArray(i, 122)
        ProgruzkaKompkreslaArray(i, 190) = ProgruzkaKompkreslaTerminalArray(i, 123)
        ProgruzkaKompkreslaArray(i, 191) = ProgruzkaKompkreslaTerminalArray(i, 124)
        ProgruzkaKompkreslaArray(i, 192) = ProgruzkaKompkreslaTerminalArray(i, 125)
        ProgruzkaKompkreslaArray(i, 204) = ProgruzkaKompkreslaTerminalArray(i, 129)

    Next
    
    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:GV" & TotalRowsKompkreslaPromo + 1).Value = ProgruzkaKompkreslaArray
    Application.ScreenUpdating = True

    If UserForm1.Autosave = True Then
        
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Компкресла Промо")

    End If

ended:
End Sub


