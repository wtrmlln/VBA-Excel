Attribute VB_Name = "B2C_Тепложар"
Option Private Module
Option Explicit

'Прогрузка Н
'Прогрузка 620
Public Sub Прогрузка_Тепложар(ByRef TeplozharArray, ByRef TotalRowsTeplozhar)

    Dim ProgruzkaFilename As String
    Dim NCity As Integer
    Dim r1 As Long
    Dim NRowProgruzka As Long
    Dim d As Integer
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a
    Dim OptPrice
    Dim MinTerminalTerminalPrice As Double
    Dim str As String
    Dim PlatniiVes, TypeZabor
    Dim Transitive As Boolean
    Dim ProgruzkaSheetname As String
    Dim NcolumnProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Tarif As Double
    Dim r3 As Long
    Dim mgorodinfo() As String
    Dim i As Long
    Dim tps As Double
    Dim CitiesArray() As String
    Dim Postavshik As String
    
    Transitive = True

    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    For i = UBound(TeplozharArray, 2) To 0 Step -1
        If TeplozharArray(0, i) = 0 Or Len(TeplozharArray(0, i)) = 0 Then
            ReDim Preserve TeplozharArray(3, UBound(TeplozharArray, 2) - 1)
        End If
    Next
    
    ReDim ProgruzkaMainArray(1 To (TotalRowsTeplozhar + 1) * 2, 1 To 235)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С Тепложар")
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Тепложар")
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

Kod1C_Change:
    
    NRowProschet = TeplozharArray(2, d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    a = Split(ProschetArray(NRowProschet, 7), ";")

    str = a(4 + TeplozharArray(3, d))

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

    NCity = 20

Komponent_Change:

    TypeZabor = FillInfocolumnTypeZabor("620", ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, TeplozharArray(3, d), 0, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = a(4 + TeplozharArray(3, d))
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 224) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 225) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, False)
    ProgruzkaMainArray(NRowProgruzka, 226) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 227) = Volume
    ProgruzkaMainArray(NRowProgruzka, 228) = Weight
    ProgruzkaMainArray(NRowProgruzka, 229) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 230) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 231) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 232) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 233) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 234) = Artikul
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To 142
        
        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * Margin * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
            
        ElseIf Komplektuyeshee = "К" Or Komplektuyeshee = "К2" Then
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

        If NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        ElseIf NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        Else
            Price = OptPriceProschet * Margin
        End If

        Tarif = fTarif(TeplozharArray(2, d), CitiesArray(NCity), VolumetricWeight)
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

        ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & ProgruzkaMainArray(1, NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
    
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If RetailPrice > 7500 Then
                    TotalPrice = Round_Up(RetailPrice, -1)
                Else
                    TotalPrice = Round_Up(RetailPrice, 0)
                End If
            End If
        ElseIf InStr(str, "t") = 0 Then
            If NCity >= 123 And NCity < 135 Then
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
        End If

unrrc:

        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Or Komplektuyeshee = "К2" Then

        For NcolumnProgruzka = 5 To 127
        
            If NcolumnProgruzka < 108 Then
                If DictSpecPostavshik.exists(ProschetArray(TeplozharArray(2, d), 1)) = True Then
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
            ElseIf NcolumnProgruzka = 126 Then
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

    Call fMixer(ProgruzkaMainArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg

federal_orenburg:

    ProgruzkaMainArray(NRowProgruzka, 223) = ProgruzkaMainArray(NRowProgruzka, 59)
    GoTo orenburg_infocolumn

orenburg_infocolumn:

    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & "125" & "::" & ProgruzkaMainArray(NRowProgruzka, 223) & "::" & "0" & "::" & "0" & "::" & "DPD "

    If UserForm1.CheckBox2 = False Then
        If ProgruzkaMainArray(NRowProgruzka, 223) < RetailPrice Then
            If OptPrice > 7500 Then
                ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(RetailPrice, -1) + 10 * (TeplozharArray(1, d) + 1) / 5
            Else
                ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(RetailPrice, 0) + (TeplozharArray(1, d) + 1) / 5
            End If
        End If
    End If

    mgorodinfo = Split(ProgruzkaMainArray(NRowProgruzka, 235), " ")
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(0), Left(mgorodinfo(0), InStr(mgorodinfo(0), "::") - 1), 128)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(1), Left(mgorodinfo(1), InStr(mgorodinfo(1), "::") - 1), 129)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(2), Left(mgorodinfo(2), InStr(mgorodinfo(2), "::") - 1), 130)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(3), Left(mgorodinfo(3), InStr(mgorodinfo(3), "::") - 1), 131)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(4), Left(mgorodinfo(4), InStr(mgorodinfo(4), "::") - 1), 132)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(5), Left(mgorodinfo(5), InStr(mgorodinfo(5), "::") - 1), 133)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(6), Left(mgorodinfo(6), InStr(mgorodinfo(6), "::") - 1), 134)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(7), Left(mgorodinfo(7), InStr(mgorodinfo(7), "::") - 1), 135)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(8), Left(mgorodinfo(8), InStr(mgorodinfo(8), "::") - 1), 136)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(9), Left(mgorodinfo(9), InStr(mgorodinfo(9), "::") - 1), 137)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(10), Left(mgorodinfo(10), InStr(mgorodinfo(10), "::") - 1), 138)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(11), Left(mgorodinfo(11), InStr(mgorodinfo(11), "::") - 1), 139)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(12), Left(mgorodinfo(12), InStr(mgorodinfo(12), "::") - 1), 140)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(13), Left(mgorodinfo(13), InStr(mgorodinfo(13), "::") - 1), 141)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(14), Left(mgorodinfo(14), InStr(mgorodinfo(14), "::") - 1), 142)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(16), Left(mgorodinfo(16), InStr(mgorodinfo(16), "::") - 1), 143)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(17), Left(mgorodinfo(17), InStr(mgorodinfo(17), "::") - 1), 144)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(18), Left(mgorodinfo(18), InStr(mgorodinfo(18), "::") - 1), 145)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(19), Left(mgorodinfo(19), InStr(mgorodinfo(19), "::") - 1), 146)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(20), Left(mgorodinfo(20), InStr(mgorodinfo(20), "::") - 1), 147)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(21), Left(mgorodinfo(21), InStr(mgorodinfo(21), "::") - 1), 148)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(22), Left(mgorodinfo(22), InStr(mgorodinfo(22), "::") - 1), 149)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(23), Left(mgorodinfo(23), InStr(mgorodinfo(23), "::") - 1), 150)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(24), Left(mgorodinfo(24), InStr(mgorodinfo(24), "::") - 1), 151)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(25), Left(mgorodinfo(25), InStr(mgorodinfo(25), "::") - 1), 152)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(26), Left(mgorodinfo(26), InStr(mgorodinfo(26), "::") - 1), 153)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(27), Left(mgorodinfo(27), InStr(mgorodinfo(27), "::") - 1), 154)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(28), Left(mgorodinfo(28), InStr(mgorodinfo(28), "::") - 1), 155)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(29), Left(mgorodinfo(29), InStr(mgorodinfo(29), "::") - 1), 156)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(30), Left(mgorodinfo(30), InStr(mgorodinfo(30), "::") - 1), 157)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(31), Left(mgorodinfo(31), InStr(mgorodinfo(31), "::") - 1), 158)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(32), Left(mgorodinfo(32), InStr(mgorodinfo(32), "::") - 1), 159)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(33), Left(mgorodinfo(33), InStr(mgorodinfo(33), "::") - 1), 160)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(35), Left(mgorodinfo(35), InStr(mgorodinfo(35), "::") - 1), 161)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(36), Left(mgorodinfo(36), InStr(mgorodinfo(36), "::") - 1), 162)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(37), Left(mgorodinfo(37), InStr(mgorodinfo(37), "::") - 1), 163)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(38), Left(mgorodinfo(38), InStr(mgorodinfo(38), "::") - 1), 164)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(39), Left(mgorodinfo(39), InStr(mgorodinfo(39), "::") - 1), 165)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(40), Left(mgorodinfo(40), InStr(mgorodinfo(40), "::") - 1), 166)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(41), Left(mgorodinfo(41), InStr(mgorodinfo(41), "::") - 1), 167)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(43), Left(mgorodinfo(43), InStr(mgorodinfo(43), "::") - 1), 168)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(44), Left(mgorodinfo(44), InStr(mgorodinfo(44), "::") - 1), 169)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(45), Left(mgorodinfo(45), InStr(mgorodinfo(45), "::") - 1), 170)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(46), Left(mgorodinfo(46), InStr(mgorodinfo(46), "::") - 1), 171)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(48), Left(mgorodinfo(48), InStr(mgorodinfo(48), "::") - 1), 172)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(52), Left(mgorodinfo(52), InStr(mgorodinfo(52), "::") - 1), 173)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(53), Left(mgorodinfo(53), InStr(mgorodinfo(53), "::") - 1), 174)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(54), Left(mgorodinfo(54), InStr(mgorodinfo(54), "::") - 1), 175)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(55), Left(mgorodinfo(55), InStr(mgorodinfo(55), "::") - 1), 176)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(56), Left(mgorodinfo(56), InStr(mgorodinfo(56), "::") - 1), 177)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(57), Left(mgorodinfo(57), InStr(mgorodinfo(57), "::") - 1), 178)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(58), Left(mgorodinfo(58), InStr(mgorodinfo(58), "::") - 1), 179)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(60), Left(mgorodinfo(60), InStr(mgorodinfo(60), "::") - 1), 180)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(62), Left(mgorodinfo(62), InStr(mgorodinfo(62), "::") - 1), 181)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(65), Left(mgorodinfo(65), InStr(mgorodinfo(65), "::") - 1), 182)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(66), Left(mgorodinfo(66), InStr(mgorodinfo(66), "::") - 1), 183)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(67), Left(mgorodinfo(67), InStr(mgorodinfo(67), "::") - 1), 184)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(68), Left(mgorodinfo(68), InStr(mgorodinfo(68), "::") - 1), 185)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(69), Left(mgorodinfo(69), InStr(mgorodinfo(69), "::") - 1), 186)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(70), Left(mgorodinfo(70), InStr(mgorodinfo(70), "::") - 1), 187)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(73), Left(mgorodinfo(73), InStr(mgorodinfo(73), "::") - 1), 188)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(74), Left(mgorodinfo(74), InStr(mgorodinfo(74), "::") - 1), 189)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(75), Left(mgorodinfo(75), InStr(mgorodinfo(75), "::") - 1), 190)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(76), Left(mgorodinfo(76), InStr(mgorodinfo(76), "::") - 1), 191)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(78), Left(mgorodinfo(78), InStr(mgorodinfo(78), "::") - 1), 192)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(79), Left(mgorodinfo(79), InStr(mgorodinfo(79), "::") - 1), 193)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(80), Left(mgorodinfo(80), InStr(mgorodinfo(80), "::") - 1), 194)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(81), Left(mgorodinfo(81), InStr(mgorodinfo(81), "::") - 1), 195)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(82), Left(mgorodinfo(82), InStr(mgorodinfo(82), "::") - 1), 196)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(83), Left(mgorodinfo(83), InStr(mgorodinfo(83), "::") - 1), 197)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(84), Left(mgorodinfo(84), InStr(mgorodinfo(84), "::") - 1), 198)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(85), Left(mgorodinfo(85), InStr(mgorodinfo(85), "::") - 1), 199)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(86), Left(mgorodinfo(86), InStr(mgorodinfo(86), "::") - 1), 200)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(87), Left(mgorodinfo(87), InStr(mgorodinfo(87), "::") - 1), 201)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(88), Left(mgorodinfo(88), InStr(mgorodinfo(88), "::") - 1), 202)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(90), Left(mgorodinfo(90), InStr(mgorodinfo(90), "::") - 1), 203)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(92), Left(mgorodinfo(92), InStr(mgorodinfo(92), "::") - 1), 204)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(94), Left(mgorodinfo(94), InStr(mgorodinfo(94), "::") - 1), 205)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(95), Left(mgorodinfo(95), InStr(mgorodinfo(95), "::") - 1), 206)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(96), Left(mgorodinfo(96), InStr(mgorodinfo(96), "::") - 1), 207)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(97), Left(mgorodinfo(97), InStr(mgorodinfo(97), "::") - 1), 208)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(98), Left(mgorodinfo(98), InStr(mgorodinfo(98), "::") - 1), 209)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(100), Left(mgorodinfo(100), InStr(mgorodinfo(100), "::") - 1), 210)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(101), Left(mgorodinfo(101), InStr(mgorodinfo(101), "::") - 1), 211)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(102), Left(mgorodinfo(102), InStr(mgorodinfo(102), "::") - 1), 212)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(104), Left(mgorodinfo(104), InStr(mgorodinfo(104), "::") - 1), 213)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(105), Left(mgorodinfo(105), InStr(mgorodinfo(105), "::") - 1), 214)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(106), Left(mgorodinfo(106), InStr(mgorodinfo(106), "::") - 1), 215)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(108), Left(mgorodinfo(108), InStr(mgorodinfo(108), "::") - 1), 216)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(111), Left(mgorodinfo(111), InStr(mgorodinfo(111), "::") - 1), 217)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(113), Left(mgorodinfo(113), InStr(mgorodinfo(113), "::") - 1), 218)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(114), Left(mgorodinfo(114), InStr(mgorodinfo(114), "::") - 1), 219)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(119), Left(mgorodinfo(119), InStr(mgorodinfo(119), "::") - 1), 220)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(121), Left(mgorodinfo(121), InStr(mgorodinfo(121), "::") - 1), 221)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & " " & Replace(mgorodinfo(122), Left(mgorodinfo(122), InStr(mgorodinfo(122), "::") - 1), 222)


Change_Row_Progruzka:

    Postavshik = ProschetArray(NRowProschet, 1)

    If ProgruzkaMainArray(NRowProgruzka, 5) = 0 Then
        GoTo Resetting
    End If

    r3 = NRowProgruzka - 1
    For NcolumnProgruzka = 5 To 127
        For r1 = r3 To NRowProgruzka
            TotalPrice = ProgruzkaMainArray(r1, NcolumnProgruzka)
            For i = r3 To NRowProgruzka
                If TotalPrice = ProgruzkaMainArray(i, NcolumnProgruzka) Then
                    If i <> r1 Then
                        If NcolumnProgruzka < 108 Then
                            If DictSpecPostavshik.exists(ProschetArray(TeplozharArray(2, d), 1)) = True Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r3 - 1
                        ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                            If ProgruzkaMainArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                            If ProgruzkaMainArray(NRowProgruzka, 4) > 5000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka = 127 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r3 - 1
                        End If
                    End If
                End If
            Next i
        Next r1
    Next NcolumnProgruzka

Resetting:
    Do While d <> UBound(TeplozharArray, 2)
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

    For i = 2 To (UBound(TeplozharArray, 2) + 1) * 2
        ProgruzkaMainArray(i, 128) = ProgruzkaMainArray(i, 5)
        ProgruzkaMainArray(i, 129) = ProgruzkaMainArray(i, 6)
        ProgruzkaMainArray(i, 130) = ProgruzkaMainArray(i, 7)
        ProgruzkaMainArray(i, 131) = ProgruzkaMainArray(i, 8)
        ProgruzkaMainArray(i, 132) = ProgruzkaMainArray(i, 9)
        ProgruzkaMainArray(i, 133) = ProgruzkaMainArray(i, 10)
        ProgruzkaMainArray(i, 134) = ProgruzkaMainArray(i, 11)
        ProgruzkaMainArray(i, 135) = ProgruzkaMainArray(i, 12)
        ProgruzkaMainArray(i, 136) = ProgruzkaMainArray(i, 13)
        ProgruzkaMainArray(i, 137) = ProgruzkaMainArray(i, 14)
        ProgruzkaMainArray(i, 138) = ProgruzkaMainArray(i, 15)
        ProgruzkaMainArray(i, 139) = ProgruzkaMainArray(i, 16)
        ProgruzkaMainArray(i, 140) = ProgruzkaMainArray(i, 17)
        ProgruzkaMainArray(i, 141) = ProgruzkaMainArray(i, 18)
        ProgruzkaMainArray(i, 142) = ProgruzkaMainArray(i, 20)
        ProgruzkaMainArray(i, 143) = ProgruzkaMainArray(i, 21)
        ProgruzkaMainArray(i, 144) = ProgruzkaMainArray(i, 22)
        ProgruzkaMainArray(i, 145) = ProgruzkaMainArray(i, 23)
        ProgruzkaMainArray(i, 146) = ProgruzkaMainArray(i, 24)
        ProgruzkaMainArray(i, 147) = ProgruzkaMainArray(i, 25)
        ProgruzkaMainArray(i, 148) = ProgruzkaMainArray(i, 26)
        ProgruzkaMainArray(i, 149) = ProgruzkaMainArray(i, 27)
        ProgruzkaMainArray(i, 150) = ProgruzkaMainArray(i, 28)
        ProgruzkaMainArray(i, 151) = ProgruzkaMainArray(i, 29)
        ProgruzkaMainArray(i, 152) = ProgruzkaMainArray(i, 30)
        ProgruzkaMainArray(i, 153) = ProgruzkaMainArray(i, 31)
        ProgruzkaMainArray(i, 154) = ProgruzkaMainArray(i, 32)
        ProgruzkaMainArray(i, 155) = ProgruzkaMainArray(i, 33)
        ProgruzkaMainArray(i, 156) = ProgruzkaMainArray(i, 34)
        ProgruzkaMainArray(i, 157) = ProgruzkaMainArray(i, 35)
        ProgruzkaMainArray(i, 158) = ProgruzkaMainArray(i, 36)
        ProgruzkaMainArray(i, 159) = ProgruzkaMainArray(i, 37)
        ProgruzkaMainArray(i, 160) = ProgruzkaMainArray(i, 38)
        ProgruzkaMainArray(i, 161) = ProgruzkaMainArray(i, 40)
        ProgruzkaMainArray(i, 162) = ProgruzkaMainArray(i, 41)
        ProgruzkaMainArray(i, 163) = ProgruzkaMainArray(i, 42)
        ProgruzkaMainArray(i, 164) = ProgruzkaMainArray(i, 43)
        ProgruzkaMainArray(i, 165) = ProgruzkaMainArray(i, 44)
        ProgruzkaMainArray(i, 166) = ProgruzkaMainArray(i, 45)
        ProgruzkaMainArray(i, 167) = ProgruzkaMainArray(i, 46)
        ProgruzkaMainArray(i, 168) = ProgruzkaMainArray(i, 48)
        ProgruzkaMainArray(i, 169) = ProgruzkaMainArray(i, 49)
        ProgruzkaMainArray(i, 170) = ProgruzkaMainArray(i, 50)
        ProgruzkaMainArray(i, 171) = ProgruzkaMainArray(i, 51)
        ProgruzkaMainArray(i, 172) = ProgruzkaMainArray(i, 53)
        ProgruzkaMainArray(i, 173) = ProgruzkaMainArray(i, 57)
        ProgruzkaMainArray(i, 174) = ProgruzkaMainArray(i, 58)
        ProgruzkaMainArray(i, 175) = ProgruzkaMainArray(i, 59)
        ProgruzkaMainArray(i, 176) = ProgruzkaMainArray(i, 60)
        ProgruzkaMainArray(i, 177) = ProgruzkaMainArray(i, 61)
        ProgruzkaMainArray(i, 178) = ProgruzkaMainArray(i, 62)
        ProgruzkaMainArray(i, 179) = ProgruzkaMainArray(i, 63)
        ProgruzkaMainArray(i, 180) = ProgruzkaMainArray(i, 65)
        ProgruzkaMainArray(i, 181) = ProgruzkaMainArray(i, 67)
        ProgruzkaMainArray(i, 182) = ProgruzkaMainArray(i, 70)
        ProgruzkaMainArray(i, 183) = ProgruzkaMainArray(i, 71)
        ProgruzkaMainArray(i, 184) = ProgruzkaMainArray(i, 72)
        ProgruzkaMainArray(i, 185) = ProgruzkaMainArray(i, 73)
        ProgruzkaMainArray(i, 186) = ProgruzkaMainArray(i, 74)
        ProgruzkaMainArray(i, 187) = ProgruzkaMainArray(i, 75)
        ProgruzkaMainArray(i, 188) = ProgruzkaMainArray(i, 78)
        ProgruzkaMainArray(i, 189) = ProgruzkaMainArray(i, 79)
        ProgruzkaMainArray(i, 190) = ProgruzkaMainArray(i, 80)
        ProgruzkaMainArray(i, 191) = ProgruzkaMainArray(i, 81)
        ProgruzkaMainArray(i, 192) = ProgruzkaMainArray(i, 83)
        ProgruzkaMainArray(i, 193) = ProgruzkaMainArray(i, 84)
        ProgruzkaMainArray(i, 194) = ProgruzkaMainArray(i, 85)
        ProgruzkaMainArray(i, 195) = ProgruzkaMainArray(i, 86)
        ProgruzkaMainArray(i, 196) = ProgruzkaMainArray(i, 87)
        ProgruzkaMainArray(i, 197) = ProgruzkaMainArray(i, 88)
        ProgruzkaMainArray(i, 198) = ProgruzkaMainArray(i, 89)
        ProgruzkaMainArray(i, 199) = ProgruzkaMainArray(i, 90)
        ProgruzkaMainArray(i, 200) = ProgruzkaMainArray(i, 91)
        ProgruzkaMainArray(i, 201) = ProgruzkaMainArray(i, 92)
        ProgruzkaMainArray(i, 202) = ProgruzkaMainArray(i, 93)
        ProgruzkaMainArray(i, 203) = ProgruzkaMainArray(i, 95)
        ProgruzkaMainArray(i, 204) = ProgruzkaMainArray(i, 97)
        ProgruzkaMainArray(i, 205) = ProgruzkaMainArray(i, 99)
        ProgruzkaMainArray(i, 206) = ProgruzkaMainArray(i, 100)
        ProgruzkaMainArray(i, 207) = ProgruzkaMainArray(i, 101)
        ProgruzkaMainArray(i, 208) = ProgruzkaMainArray(i, 102)
        ProgruzkaMainArray(i, 209) = ProgruzkaMainArray(i, 103)
        ProgruzkaMainArray(i, 210) = ProgruzkaMainArray(i, 105)
        ProgruzkaMainArray(i, 211) = ProgruzkaMainArray(i, 106)
        ProgruzkaMainArray(i, 212) = ProgruzkaMainArray(i, 107)
        ProgruzkaMainArray(i, 213) = ProgruzkaMainArray(i, 109)
        ProgruzkaMainArray(i, 214) = ProgruzkaMainArray(i, 110)
        ProgruzkaMainArray(i, 215) = ProgruzkaMainArray(i, 111)
        ProgruzkaMainArray(i, 216) = ProgruzkaMainArray(i, 113)
        ProgruzkaMainArray(i, 217) = ProgruzkaMainArray(i, 116)
        ProgruzkaMainArray(i, 218) = ProgruzkaMainArray(i, 118)
        ProgruzkaMainArray(i, 219) = ProgruzkaMainArray(i, 119)
        ProgruzkaMainArray(i, 220) = ProgruzkaMainArray(i, 124)
        ProgruzkaMainArray(i, 221) = ProgruzkaMainArray(i, 126)
        ProgruzkaMainArray(i, 222) = ProgruzkaMainArray(i, 127)
    Next

    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:IA" & NRowProgruzka).Value = ProgruzkaMainArray
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:IA" & NRowProgruzka).RemoveDuplicates Columns:=3, Header:=xlYes
    Application.ScreenUpdating = True
    
    If UserForm1.Autosave = True Then
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Тепложар")
    End If

ended:
End Sub

Public Sub Прогрузка_Промо_Тепложар(ByRef TeplozharPromoArray, ByRef TotalRowsTeplozharPromo)
   
    Dim r1 As Long
    Dim NRowProgruzka As Long
    Dim d As Integer
    Dim a
    Dim PlatniiVes, a1() As String, TypeZabor
    Dim MinTerminalTerminalPrice As Double
    Dim CitiesArray() As String
    Dim Transitive As Boolean
    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NCity As Long
    Dim NcolumnProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Postavshik As String
    Dim Margin As Double
    Dim TotalPrice As Long
    Dim Price As Long
    Dim Tarif As Double
    Dim TerminalTerminalPrice As Double
    Dim FullDostavkaPrice As Long
    Dim OptPrice As Long
    Dim r3 As Long
    Dim mgorodinfo() As String
    Dim i As Long
    Dim tps As Double

    Transitive = True
    
    For i = UBound(TeplozharPromoArray, 2) To 0 Step -1
        If TeplozharPromoArray(0, i) = 0 Or Len(TeplozharPromoArray(0, i)) = 0 Then
            ReDim Preserve TeplozharPromoArray(3, UBound(TeplozharPromoArray, 2) - 1)
        End If
    Next

    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    ReDim ProgruzkaMainArray(1 To (TotalRowsTeplozharPromo + 1) * 2, 1 To 235)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С Тепложар")
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Тепложар")

    NCity = 20
    NcolumnProgruzka = 5
    NRowProgruzka = 2
    
    If UserForm2.CheckBox_Margin_Promo = False Then
        If UserForm2.TextBox1 = 0 Or Len(UserForm2.TextBox1) = 0 Then
            MsgBox "Некорректная наценка/категория"
            GoTo ended
        Else
            a = Split(UserForm2.TextBox1, ";")
        End If
    ElseIf UserForm2.CheckBox_Margin_Promo = True Then
        a = Split(UserForm2.Label2.Caption, ";")
    End If

Kod1C_Change:
    
    NRowProschet = TeplozharPromoArray(2, d)
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

    tps = fTPS(Weight, Quantity)

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ProschetArray(NRowProschet, 13) <> 0 Then
        ZaborPrice = 0
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

    NCity = 20

Komponent_Change:

    TypeZabor = FillInfocolumnTypeZabor("620", ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, 0, 1, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = "620"
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 224) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 225) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, True)
    ProgruzkaMainArray(NRowProgruzka, 226) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 227) = Volume
    ProgruzkaMainArray(NRowProgruzka, 228) = Weight
    ProgruzkaMainArray(NRowProgruzka, 229) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 230) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 231) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 232) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 233) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 234) = Artikul

    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If

    For NCity = 20 To 142
        
        If Komplektuyeshee = "Т" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * Margin
            ElseIf NCity < 135 Then
                TotalPrice = OptPriceProschet * Margin * KURS_TENGE
            ElseIf NCity < 141 Then
                TotalPrice = OptPriceProschet * Margin * KURS_BELRUB
            ElseIf NCity < 142 Then
                TotalPrice = OptPriceProschet * Margin * KURS_SOM
            Else
                TotalPrice = OptPriceProschet * Margin * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        ElseIf Komplektuyeshee = "К" Or Komplektuyeshee = "К2" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 2
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * 2 * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * 2 * KURS_SOM
            ElseIf NCity >= 142 Then
                TotalPrice = OptPriceProschet * 2 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
            
        ElseIf Komplektuyeshee = "К1" Then
            If NCity < 123 Then
                TotalPrice = OptPriceProschet * 1.7
            ElseIf NCity >= 123 And NCity < 135 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_TENGE
            ElseIf NCity >= 135 And NCity < 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_BELRUB
            ElseIf NCity = 141 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_SOM
            ElseIf NCity >= 142 Then
                TotalPrice = OptPriceProschet * 1.7 * KURS_DRAM
            End If
            TotalPrice = Round_Up(TotalPrice, 0)
            GoTo zapis
        
        End If

        If NCity = 44 Then
            Price = OptPriceProschet * (Margin + 0.2)
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        Else
            Price = OptPriceProschet * Margin
        End If

        Tarif = fTarif(TeplozharPromoArray(2, d), CitiesArray(NCity), VolumetricWeight)
        MinTerminalTerminalPrice = fTarif(TeplozharPromoArray(2, d), CitiesArray(NCity))

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
                TotalPrice = Round_Up(((TotalPrice + JU)) * KURS_BELRUB, 0)
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
zapis:

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If

        ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & ProgruzkaMainArray(1, NcolumnProgruzka) & "::" & TotalPrice & "::" & TerminalTerminalPrice & "::" & FullDostavkaPrice & "::" & "DPD "
        
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If RetailPrice > 7500 Then
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

        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Or Komplektuyeshee = "К2" Then

        For NcolumnProgruzka = 5 To 127
            If NcolumnProgruzka < 108 Then
                If DictSpecPostavshik.exists(ProschetArray(TeplozharPromoArray(2, d), 1)) = True Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 5)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 5)
                End If
            ElseIf NcolumnProgruzka >= 108 And NCity < 120 Then
                ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 108)
            ElseIf NcolumnProgruzka >= 120 And NCity < 126 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 150000 Then
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 10 * (NcolumnProgruzka - 120)
                Else
                    ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) + 1 * (NcolumnProgruzka - 120)
                End If
            ElseIf NcolumnProgruzka >= 126 And NCity < 127 Then
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

    Call fMixer(ProgruzkaMainArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg

federal_orenburg:

    ProgruzkaMainArray(NRowProgruzka, 223) = ProgruzkaMainArray(NRowProgruzka, 59)
    GoTo orenburg_infocolumn

    a1 = Split(ProschetArray(TeplozharPromoArray(2, d), 7), ";")
        
    If Komplektuyeshee = "К" Then
        If OptPriceProschet > 7500 Then
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.9, -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.9, 0)
        End If
    ElseIf Komplektuyeshee = "К1" Then
        If OptPriceProschet > 7500 Then
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.6, -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.6, 0)
        End If
    ElseIf Komplektuyeshee = "К2" Then
        If OptPriceProschet > 7500 Then
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.9, -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(OptPriceProschet * 1.9, 0)
        End If
    ElseIf OptPriceProschet > 7500 Then
        If Quantity > 1 Then
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up((OptPriceProschet * (Margin - 0.1) + RazgruzkaPrice / Quantity), -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up(((OptPriceProschet * (Margin - 0.1)) + RazgruzkaPrice), -1)
        End If
    ElseIf Quantity > 1 Then
        ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up((OptPriceProschet * (Margin - 0.1) + RazgruzkaPrice / Quantity), 0)
    Else
        ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up((OptPriceProschet * (Margin - 0.1) + RazgruzkaPrice), 0)
    End If

orenburg_infocolumn:
    
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & "125" & "::" & ProgruzkaMainArray(NRowProgruzka, 223) & "::" & "0" & "::" & "0" & "::" & "DPD "
    
    If UserForm1.CheckBox2 = False Then
        If ProgruzkaMainArray(NRowProgruzka, 223) < RetailPrice Then
            If OptPrice > 7500 Then
                ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up((RetailPrice * (1)), -1) + 10 * (UBound(a1) + 1) / 5
            Else
                ProgruzkaMainArray(NRowProgruzka, 223) = Round_Up((RetailPrice * (1)), 0) + (UBound(a1) + 1) / 5
            End If
        End If
    End If

    mgorodinfo = Split(ProgruzkaMainArray(NRowProgruzka, 235), " ")
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(0), Left(mgorodinfo(0), InStr(mgorodinfo(0), "::") - 1), 128)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(1), Left(mgorodinfo(1), InStr(mgorodinfo(1), "::") - 1), 129)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(2), Left(mgorodinfo(2), InStr(mgorodinfo(2), "::") - 1), 130)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(3), Left(mgorodinfo(3), InStr(mgorodinfo(3), "::") - 1), 131)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(4), Left(mgorodinfo(4), InStr(mgorodinfo(4), "::") - 1), 132)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(5), Left(mgorodinfo(5), InStr(mgorodinfo(5), "::") - 1), 133)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(6), Left(mgorodinfo(6), InStr(mgorodinfo(6), "::") - 1), 134)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(7), Left(mgorodinfo(7), InStr(mgorodinfo(7), "::") - 1), 135)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(8), Left(mgorodinfo(8), InStr(mgorodinfo(8), "::") - 1), 136)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(9), Left(mgorodinfo(9), InStr(mgorodinfo(9), "::") - 1), 137)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(10), Left(mgorodinfo(10), InStr(mgorodinfo(10), "::") - 1), 138)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(11), Left(mgorodinfo(11), InStr(mgorodinfo(11), "::") - 1), 139)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(12), Left(mgorodinfo(12), InStr(mgorodinfo(12), "::") - 1), 140)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(13), Left(mgorodinfo(13), InStr(mgorodinfo(13), "::") - 1), 141)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(14), Left(mgorodinfo(14), InStr(mgorodinfo(14), "::") - 1), 142)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(16), Left(mgorodinfo(16), InStr(mgorodinfo(16), "::") - 1), 143)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(17), Left(mgorodinfo(17), InStr(mgorodinfo(17), "::") - 1), 144)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(18), Left(mgorodinfo(18), InStr(mgorodinfo(18), "::") - 1), 145)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(19), Left(mgorodinfo(19), InStr(mgorodinfo(19), "::") - 1), 146)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(20), Left(mgorodinfo(20), InStr(mgorodinfo(20), "::") - 1), 147)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(21), Left(mgorodinfo(21), InStr(mgorodinfo(21), "::") - 1), 148)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(22), Left(mgorodinfo(22), InStr(mgorodinfo(22), "::") - 1), 149)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(23), Left(mgorodinfo(23), InStr(mgorodinfo(23), "::") - 1), 150)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(24), Left(mgorodinfo(24), InStr(mgorodinfo(24), "::") - 1), 151)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(25), Left(mgorodinfo(25), InStr(mgorodinfo(25), "::") - 1), 152)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(26), Left(mgorodinfo(26), InStr(mgorodinfo(26), "::") - 1), 153)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(27), Left(mgorodinfo(27), InStr(mgorodinfo(27), "::") - 1), 154)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(28), Left(mgorodinfo(28), InStr(mgorodinfo(28), "::") - 1), 155)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(29), Left(mgorodinfo(29), InStr(mgorodinfo(29), "::") - 1), 156)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(30), Left(mgorodinfo(30), InStr(mgorodinfo(30), "::") - 1), 157)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(31), Left(mgorodinfo(31), InStr(mgorodinfo(31), "::") - 1), 158)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(32), Left(mgorodinfo(32), InStr(mgorodinfo(32), "::") - 1), 159)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(33), Left(mgorodinfo(33), InStr(mgorodinfo(33), "::") - 1), 160)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(35), Left(mgorodinfo(35), InStr(mgorodinfo(35), "::") - 1), 161)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(36), Left(mgorodinfo(36), InStr(mgorodinfo(36), "::") - 1), 162)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(37), Left(mgorodinfo(37), InStr(mgorodinfo(37), "::") - 1), 163)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(38), Left(mgorodinfo(38), InStr(mgorodinfo(38), "::") - 1), 164)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(39), Left(mgorodinfo(39), InStr(mgorodinfo(39), "::") - 1), 165)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(40), Left(mgorodinfo(40), InStr(mgorodinfo(40), "::") - 1), 166)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(41), Left(mgorodinfo(41), InStr(mgorodinfo(41), "::") - 1), 167)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(43), Left(mgorodinfo(43), InStr(mgorodinfo(43), "::") - 1), 168)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(44), Left(mgorodinfo(44), InStr(mgorodinfo(44), "::") - 1), 169)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(45), Left(mgorodinfo(45), InStr(mgorodinfo(45), "::") - 1), 170)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(46), Left(mgorodinfo(46), InStr(mgorodinfo(46), "::") - 1), 171)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(48), Left(mgorodinfo(48), InStr(mgorodinfo(48), "::") - 1), 172)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(52), Left(mgorodinfo(52), InStr(mgorodinfo(52), "::") - 1), 173)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(53), Left(mgorodinfo(53), InStr(mgorodinfo(53), "::") - 1), 174)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(54), Left(mgorodinfo(54), InStr(mgorodinfo(54), "::") - 1), 175)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(55), Left(mgorodinfo(55), InStr(mgorodinfo(55), "::") - 1), 176)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(56), Left(mgorodinfo(56), InStr(mgorodinfo(56), "::") - 1), 177)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(57), Left(mgorodinfo(57), InStr(mgorodinfo(57), "::") - 1), 178)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(58), Left(mgorodinfo(58), InStr(mgorodinfo(58), "::") - 1), 179)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(60), Left(mgorodinfo(60), InStr(mgorodinfo(60), "::") - 1), 180)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(62), Left(mgorodinfo(62), InStr(mgorodinfo(62), "::") - 1), 181)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(65), Left(mgorodinfo(65), InStr(mgorodinfo(65), "::") - 1), 182)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(66), Left(mgorodinfo(66), InStr(mgorodinfo(66), "::") - 1), 183)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(67), Left(mgorodinfo(67), InStr(mgorodinfo(67), "::") - 1), 184)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(68), Left(mgorodinfo(68), InStr(mgorodinfo(68), "::") - 1), 185)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(69), Left(mgorodinfo(69), InStr(mgorodinfo(69), "::") - 1), 186)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(70), Left(mgorodinfo(70), InStr(mgorodinfo(70), "::") - 1), 187)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(73), Left(mgorodinfo(73), InStr(mgorodinfo(73), "::") - 1), 188)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(74), Left(mgorodinfo(74), InStr(mgorodinfo(74), "::") - 1), 189)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(75), Left(mgorodinfo(75), InStr(mgorodinfo(75), "::") - 1), 190)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(76), Left(mgorodinfo(76), InStr(mgorodinfo(76), "::") - 1), 191)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(78), Left(mgorodinfo(78), InStr(mgorodinfo(78), "::") - 1), 192)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(79), Left(mgorodinfo(79), InStr(mgorodinfo(79), "::") - 1), 193)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(80), Left(mgorodinfo(80), InStr(mgorodinfo(80), "::") - 1), 194)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(81), Left(mgorodinfo(81), InStr(mgorodinfo(81), "::") - 1), 195)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(82), Left(mgorodinfo(82), InStr(mgorodinfo(82), "::") - 1), 196)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(83), Left(mgorodinfo(83), InStr(mgorodinfo(83), "::") - 1), 197)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(84), Left(mgorodinfo(84), InStr(mgorodinfo(84), "::") - 1), 198)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(85), Left(mgorodinfo(85), InStr(mgorodinfo(85), "::") - 1), 199)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(86), Left(mgorodinfo(86), InStr(mgorodinfo(86), "::") - 1), 200)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(87), Left(mgorodinfo(87), InStr(mgorodinfo(87), "::") - 1), 201)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(88), Left(mgorodinfo(88), InStr(mgorodinfo(88), "::") - 1), 202)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(90), Left(mgorodinfo(90), InStr(mgorodinfo(90), "::") - 1), 203)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(92), Left(mgorodinfo(92), InStr(mgorodinfo(92), "::") - 1), 204)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(94), Left(mgorodinfo(94), InStr(mgorodinfo(94), "::") - 1), 205)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(95), Left(mgorodinfo(95), InStr(mgorodinfo(95), "::") - 1), 206)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(96), Left(mgorodinfo(96), InStr(mgorodinfo(96), "::") - 1), 207)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(97), Left(mgorodinfo(97), InStr(mgorodinfo(97), "::") - 1), 208)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(98), Left(mgorodinfo(98), InStr(mgorodinfo(98), "::") - 1), 209)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(100), Left(mgorodinfo(100), InStr(mgorodinfo(100), "::") - 1), 210)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(101), Left(mgorodinfo(101), InStr(mgorodinfo(101), "::") - 1), 211)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(102), Left(mgorodinfo(102), InStr(mgorodinfo(102), "::") - 1), 212)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(104), Left(mgorodinfo(104), InStr(mgorodinfo(104), "::") - 1), 213)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(105), Left(mgorodinfo(105), InStr(mgorodinfo(105), "::") - 1), 214)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(106), Left(mgorodinfo(106), InStr(mgorodinfo(106), "::") - 1), 215)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(108), Left(mgorodinfo(108), InStr(mgorodinfo(108), "::") - 1), 216)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(111), Left(mgorodinfo(111), InStr(mgorodinfo(111), "::") - 1), 217)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(113), Left(mgorodinfo(113), InStr(mgorodinfo(113), "::") - 1), 218)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(114), Left(mgorodinfo(114), InStr(mgorodinfo(114), "::") - 1), 219)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(119), Left(mgorodinfo(119), InStr(mgorodinfo(119), "::") - 1), 220)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(121), Left(mgorodinfo(121), InStr(mgorodinfo(121), "::") - 1), 221)
    ProgruzkaMainArray(NRowProgruzka, 235) = ProgruzkaMainArray(NRowProgruzka, 235) & Replace(mgorodinfo(122), Left(mgorodinfo(122), InStr(mgorodinfo(122), "::") - 1), 222)

Change_Row_Progruzka:
    Postavshik = ProschetArray(NRowProschet, 1)

    If ProgruzkaMainArray(NRowProgruzka, 5) = 0 Then
        GoTo Resetting
    End If

    r3 = NRowProgruzka - 1
    For NcolumnProgruzka = 5 To 127
        For r1 = r3 To NRowProgruzka
            TotalPrice = ProgruzkaMainArray(r1, NcolumnProgruzka)
            For i = r3 To NRowProgruzka
                If TotalPrice = ProgruzkaMainArray(i, NcolumnProgruzka) Then
                    If i <> r1 Then
                        If NcolumnProgruzka < 108 Then
                            If DictSpecPostavshik.exists(ProschetArray(TeplozharPromoArray(2, d), 1)) = True Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            ElseIf ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 108 And NcolumnProgruzka < 120 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r3 - 1
                        ElseIf NcolumnProgruzka >= 120 And NcolumnProgruzka < 126 Then
                            If ProgruzkaMainArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka >= 126 And NcolumnProgruzka < 127 Then
                            If ProgruzkaMainArray(NRowProgruzka, 4) > 5000 Then
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                                TotalPrice = TotalPrice + 10
                                i = r3 - 1
                            Else
                                ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 1
                                TotalPrice = TotalPrice + 1
                                i = r3 - 1
                            End If
                        ElseIf NcolumnProgruzka = 127 Then
                            ProgruzkaMainArray(r1, NcolumnProgruzka) = ProgruzkaMainArray(r1, NcolumnProgruzka) + 10
                            TotalPrice = TotalPrice + 10
                            i = r3 - 1
                        End If
                    End If
                End If
            Next i
        Next r1
    Next NcolumnProgruzka

Resetting:
    Do While d <> UBound(TeplozharPromoArray, 2)
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

    For i = 2 To (UBound(TeplozharPromoArray, 2) + 1) * 2
        ProgruzkaMainArray(i, 128) = ProgruzkaMainArray(i, 5)
        ProgruzkaMainArray(i, 129) = ProgruzkaMainArray(i, 6)
        ProgruzkaMainArray(i, 130) = ProgruzkaMainArray(i, 7)
        ProgruzkaMainArray(i, 131) = ProgruzkaMainArray(i, 8)
        ProgruzkaMainArray(i, 132) = ProgruzkaMainArray(i, 9)
        ProgruzkaMainArray(i, 133) = ProgruzkaMainArray(i, 10)
        ProgruzkaMainArray(i, 134) = ProgruzkaMainArray(i, 11)
        ProgruzkaMainArray(i, 135) = ProgruzkaMainArray(i, 12)
        ProgruzkaMainArray(i, 136) = ProgruzkaMainArray(i, 13)
        ProgruzkaMainArray(i, 137) = ProgruzkaMainArray(i, 14)
        ProgruzkaMainArray(i, 138) = ProgruzkaMainArray(i, 15)
        ProgruzkaMainArray(i, 139) = ProgruzkaMainArray(i, 16)
        ProgruzkaMainArray(i, 140) = ProgruzkaMainArray(i, 17)
        ProgruzkaMainArray(i, 141) = ProgruzkaMainArray(i, 18)
        ProgruzkaMainArray(i, 142) = ProgruzkaMainArray(i, 20)
        ProgruzkaMainArray(i, 143) = ProgruzkaMainArray(i, 21)
        ProgruzkaMainArray(i, 144) = ProgruzkaMainArray(i, 22)
        ProgruzkaMainArray(i, 145) = ProgruzkaMainArray(i, 23)
        ProgruzkaMainArray(i, 146) = ProgruzkaMainArray(i, 24)
        ProgruzkaMainArray(i, 147) = ProgruzkaMainArray(i, 25)
        ProgruzkaMainArray(i, 148) = ProgruzkaMainArray(i, 26)
        ProgruzkaMainArray(i, 149) = ProgruzkaMainArray(i, 27)
        ProgruzkaMainArray(i, 150) = ProgruzkaMainArray(i, 28)
        ProgruzkaMainArray(i, 151) = ProgruzkaMainArray(i, 29)
        ProgruzkaMainArray(i, 152) = ProgruzkaMainArray(i, 30)
        ProgruzkaMainArray(i, 153) = ProgruzkaMainArray(i, 31)
        ProgruzkaMainArray(i, 154) = ProgruzkaMainArray(i, 32)
        ProgruzkaMainArray(i, 155) = ProgruzkaMainArray(i, 33)
        ProgruzkaMainArray(i, 156) = ProgruzkaMainArray(i, 34)
        ProgruzkaMainArray(i, 157) = ProgruzkaMainArray(i, 35)
        ProgruzkaMainArray(i, 158) = ProgruzkaMainArray(i, 36)
        ProgruzkaMainArray(i, 159) = ProgruzkaMainArray(i, 37)
        ProgruzkaMainArray(i, 160) = ProgruzkaMainArray(i, 38)
        ProgruzkaMainArray(i, 161) = ProgruzkaMainArray(i, 40)
        ProgruzkaMainArray(i, 162) = ProgruzkaMainArray(i, 41)
        ProgruzkaMainArray(i, 163) = ProgruzkaMainArray(i, 42)
        ProgruzkaMainArray(i, 164) = ProgruzkaMainArray(i, 43)
        ProgruzkaMainArray(i, 165) = ProgruzkaMainArray(i, 44)
        ProgruzkaMainArray(i, 166) = ProgruzkaMainArray(i, 45)
        ProgruzkaMainArray(i, 167) = ProgruzkaMainArray(i, 46)
        ProgruzkaMainArray(i, 168) = ProgruzkaMainArray(i, 48)
        ProgruzkaMainArray(i, 169) = ProgruzkaMainArray(i, 49)
        ProgruzkaMainArray(i, 170) = ProgruzkaMainArray(i, 50)
        ProgruzkaMainArray(i, 171) = ProgruzkaMainArray(i, 51)
        ProgruzkaMainArray(i, 172) = ProgruzkaMainArray(i, 53)
        ProgruzkaMainArray(i, 173) = ProgruzkaMainArray(i, 57)
        ProgruzkaMainArray(i, 174) = ProgruzkaMainArray(i, 58)
        ProgruzkaMainArray(i, 175) = ProgruzkaMainArray(i, 59)
        ProgruzkaMainArray(i, 176) = ProgruzkaMainArray(i, 60)
        ProgruzkaMainArray(i, 177) = ProgruzkaMainArray(i, 61)
        ProgruzkaMainArray(i, 178) = ProgruzkaMainArray(i, 62)
        ProgruzkaMainArray(i, 179) = ProgruzkaMainArray(i, 63)
        ProgruzkaMainArray(i, 180) = ProgruzkaMainArray(i, 65)
        ProgruzkaMainArray(i, 181) = ProgruzkaMainArray(i, 67)
        ProgruzkaMainArray(i, 182) = ProgruzkaMainArray(i, 70)
        ProgruzkaMainArray(i, 183) = ProgruzkaMainArray(i, 71)
        ProgruzkaMainArray(i, 184) = ProgruzkaMainArray(i, 72)
        ProgruzkaMainArray(i, 185) = ProgruzkaMainArray(i, 73)
        ProgruzkaMainArray(i, 186) = ProgruzkaMainArray(i, 74)
        ProgruzkaMainArray(i, 187) = ProgruzkaMainArray(i, 75)
        ProgruzkaMainArray(i, 188) = ProgruzkaMainArray(i, 78)
        ProgruzkaMainArray(i, 189) = ProgruzkaMainArray(i, 79)
        ProgruzkaMainArray(i, 190) = ProgruzkaMainArray(i, 80)
        ProgruzkaMainArray(i, 191) = ProgruzkaMainArray(i, 81)
        ProgruzkaMainArray(i, 192) = ProgruzkaMainArray(i, 83)
        ProgruzkaMainArray(i, 193) = ProgruzkaMainArray(i, 84)
        ProgruzkaMainArray(i, 194) = ProgruzkaMainArray(i, 85)
        ProgruzkaMainArray(i, 195) = ProgruzkaMainArray(i, 86)
        ProgruzkaMainArray(i, 196) = ProgruzkaMainArray(i, 87)
        ProgruzkaMainArray(i, 197) = ProgruzkaMainArray(i, 88)
        ProgruzkaMainArray(i, 198) = ProgruzkaMainArray(i, 89)
        ProgruzkaMainArray(i, 199) = ProgruzkaMainArray(i, 90)
        ProgruzkaMainArray(i, 200) = ProgruzkaMainArray(i, 91)
        ProgruzkaMainArray(i, 201) = ProgruzkaMainArray(i, 92)
        ProgruzkaMainArray(i, 202) = ProgruzkaMainArray(i, 93)
        ProgruzkaMainArray(i, 203) = ProgruzkaMainArray(i, 95)
        ProgruzkaMainArray(i, 204) = ProgruzkaMainArray(i, 97)
        ProgruzkaMainArray(i, 205) = ProgruzkaMainArray(i, 99)
        ProgruzkaMainArray(i, 206) = ProgruzkaMainArray(i, 100)
        ProgruzkaMainArray(i, 207) = ProgruzkaMainArray(i, 101)
        ProgruzkaMainArray(i, 208) = ProgruzkaMainArray(i, 102)
        ProgruzkaMainArray(i, 209) = ProgruzkaMainArray(i, 103)
        ProgruzkaMainArray(i, 210) = ProgruzkaMainArray(i, 105)
        ProgruzkaMainArray(i, 211) = ProgruzkaMainArray(i, 106)
        ProgruzkaMainArray(i, 212) = ProgruzkaMainArray(i, 107)
        ProgruzkaMainArray(i, 213) = ProgruzkaMainArray(i, 109)
        ProgruzkaMainArray(i, 214) = ProgruzkaMainArray(i, 110)
        ProgruzkaMainArray(i, 215) = ProgruzkaMainArray(i, 111)
        ProgruzkaMainArray(i, 216) = ProgruzkaMainArray(i, 113)
        ProgruzkaMainArray(i, 217) = ProgruzkaMainArray(i, 116)
        ProgruzkaMainArray(i, 218) = ProgruzkaMainArray(i, 118)
        ProgruzkaMainArray(i, 219) = ProgruzkaMainArray(i, 119)
        ProgruzkaMainArray(i, 220) = ProgruzkaMainArray(i, 124)
        ProgruzkaMainArray(i, 221) = ProgruzkaMainArray(i, 126)
        ProgruzkaMainArray(i, 222) = ProgruzkaMainArray(i, 127)
    Next

    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:IA" & (TotalRowsTeplozharPromo + 1) * 2).Value = ProgruzkaMainArray
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:IA" & (TotalRowsTeplozharPromo + 1) * 2).RemoveDuplicates Columns:=3, Header:=xlYes
    Application.ScreenUpdating = True
    
    If UserForm1.Autosave = True Then
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Тепложар Промо")

    End If

ended:
End Sub


