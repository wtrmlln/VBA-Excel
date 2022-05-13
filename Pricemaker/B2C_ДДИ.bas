Attribute VB_Name = "B2C_ДДИ"
Option Private Module
Option Explicit

'Прогрузка DDI
Public Sub Прогрузка_DDI()
    
    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NcolumnProgruzka As Long
    Dim NCity As Long
    Dim r1 As Long
    Dim NRowProgruzka As Long
    Dim d As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a() As String
    Dim OptPrice As Long
    Dim RetailPrice As Long
    Dim MinTerminalTerminalPrice As Double
    Dim Postavshik As String
    Dim CitiesArray() As String
    Dim mvvod()
    Dim mvvodin() As String
    Dim NRowsArray()
    Dim str As String
    Dim smkat As Long
    Dim TotalRows As Long
    Dim PlatniiVes
    Dim TypeZabor As String
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim LgotniiGorod As Boolean
    Dim Tarif As Double
    Dim ErrorCount As Long
    Dim i As Long
    Dim tps As Double

    Dim Transitive As Boolean
    Transitive = True

    mvvodin = Split(UserForm1.TextBox1, vbCrLf)
    If (UBound(mvvodin)) = -1 Then
        GoTo ended
    End If

    ReDim mvvod(UBound(mvvodin))

    For i = 0 To UBound(mvvodin)
        mvvodin(i) = Application.Trim(mvvodin(i))
        If Len(mvvodin(i)) = 0 Or mvvodin(i) = vbNullString Then
            ReDim Preserve mvvod(UBound(mvvod) - 1)
        Else
            mvvod(i - (UBound(mvvodin) - UBound(mvvod))) = mvvodin(i)
        End If
    Next

    If DublicateFind(mvvod) = True Then
        GoTo ended
    End If

    If SaleFind(mvvod) = True Then
        GoTo ended
    End If

    ReDim NRowsArray(UBound(mvvod))
    
    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    smkat = 0
    NRowProgruzka = 2
    ErrorCount = 0
    
    For d = 0 To UBound(mvvod)
        
        str = mvvod(LBound(mvvod) + d)
        
        r1 = Proschet_Dict.Item(str)
        NRowsArray(d) = r1
        
        a = Split(ProschetArray(NRowsArray(d), 7), ";")

        ReDim mvvodin(UBound(a))
        For i = 0 To UBound(a)
            If a(i) = vbNullString Then
                ReDim Preserve mvvodin(UBound(mvvodin) - 1)
            Else
                mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
            End If
        Next

        If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
            MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(d), 17)
            GoTo ended
        End If

        TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows

optimization1:
        If d < UBound(mvvod) Then
            If Trim(StrConv(ProschetArray(r1 + 1, 17), vbUpperCase)) = mvvod(LBound(mvvod) + d + 1) Then
                r1 = r1 + 1
                d = d + 1
                NRowsArray(d) = r1

                a = Split(ProschetArray(NRowsArray(d), 7), ";")

                ReDim mvvodin(UBound(a))
                For i = 0 To UBound(a)
                    If a(i) = vbNullString Then
                        ErrorCount = ErrorCount + 1
                        ReDim Preserve mvvodin(UBound(mvvodin) - 1)
                    Else
                        mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
                    End If
                Next
                
                If ErrorCount > 0 Then
                    ActiveWorkbook.ActiveSheet.Cells(NRowsArray(i), 7).Value = Join(mvvodin, ";")
                    ProschetArray(NRowsArray(i), 7) = Join(mvvodin, ";")
                End If

                If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
                    MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(d), 17)
                    GoTo ended
                End If

                TotalRows = (UBound(mvvodin) + 1) / 5 + TotalRows

                GoTo optimization1
            End If
        End If
    Next

    d = 0

    If SaleFind(mvvod) = True Then
        GoTo ended
    End If

    ReDim ProgruzkaMainArray(1 To TotalRows + 1, 1 To 70)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С ДДИ")
    
    ReDim CitiesArray(20 To 74)
    CitiesArray = CreateCities(CitiesArray, "Б2С ДДИ")

Kod1C_Change:
    
    NRowProschet = NRowsArray(d)
    For i = 1 To UBound(ProschetArray, 2)
        ProschetArray(NRowProschet, i) = DeclareDatatypeProschet(i, ProschetArray(NRowProschet, i))
    Next i
    
    a = Split(ProschetArray(NRowProschet, 7), ";")

    ReDim mvvodin(UBound(a))

    For i = 0 To UBound(a)
        If a(i) = vbNullString Then
            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
        Else
            mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
        End If
    Next

    'Определение переменных из просчета:
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
        ZaborPrice = ZaborPriceProschet
        LgotniiGorod = False
    End If

    If d <> 0 And ProschetArray(NRowProschet, 1) = Postavshik Then
        GoTo Komponent_Change
    ElseIf d <> 0 Then
        Transitive = False
    End If


Komponent_Change:
    str = a(4 + smkat)

    TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps)
    If TypeZabor = "Err" Then GoTo ended
    
    Margin = fMargin(OptPriceProschet, Quantity, smkat, 0, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = a(4 + smkat)
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 60) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 61) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, False)
    ProgruzkaMainArray(NRowProgruzka, 62) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 63) = Volume
    ProgruzkaMainArray(NRowProgruzka, 64) = Weight
    ProgruzkaMainArray(NRowProgruzka, 65) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 66) = FillInfocolumnJU("СОТ/Торуда", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 67) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 68) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 69) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 70) = Artikul
    
    If IsError(OptPriceProschet) Then
        GoTo Change_Row_Progruzka
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Change_Row_Progruzka
    End If
    
    For NCity = 20 To UBound(CitiesArray, 1)

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
            GoTo zapis
        
        ElseIf Komplektuyeshee = "К2" Then
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
            GoTo zapis
    
        End If

        Price = OptPriceProschet * (Margin)

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
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96), -1)
            ElseIf OptPrice >= 1000 Then
                TotalPrice = Round_Up(((TotalPrice + JU) / 0.96), 0)
            Else
                TotalPrice = Round_Up(((TotalPrice + JU + 200) / 0.96), 0)
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

zapis:

        If UserForm1.CheckBox2 = True Then
            GoTo unrrc
        End If
    
        If NCity < 123 Then
            If TotalPrice < RetailPrice Then
                If Quantity <= 1 Then
                    TotalPrice = Round_Up(RetailPrice, -1)
                Else
                    TotalPrice = Round_Up(RetailPrice, 0)
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

    Call fMixer(ProgruzkaMainArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg

federal_orenburg:
Change_Row_Progruzka:

    Do While smkat <> ((UBound(mvvodin) + 1) - 5)
        smkat = smkat + 5
        NcolumnProgruzka = 5
        If ProgruzkaMainArray(NRowProgruzka, 2) <> vbNullString Then
            NRowProgruzka = NRowProgruzka + 1
        End If
        GoTo Komponent_Change
    Loop

    Postavshik = ProschetArray(NRowProschet, 1)

    If IsError(OptPriceProschet) Then
        GoTo Resetting
    End If

    If IsNumeric(OptPriceProschet) = False Or OptPriceProschet = 0 Then
        GoTo Resetting
    End If

    If ProgruzkaMainArray(NRowProgruzka, 2) = vbNullString Then
        NRowProgruzka = NRowProgruzka - 1
    End If

Resetting:
    Do While d <> UBound(NRowsArray)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        smkat = 0

        If ProgruzkaMainArray(NRowProgruzka, 2) <> vbNullString Then
            NRowProgruzka = NRowProgruzka + 1
        End If
        GoTo Kod1C_Change
    Loop

    If Transitive = False Then
        Postavshik = vbNullString
    End If
    
    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:BR" & NRowProgruzka).Value = ProgruzkaMainArray
    Application.ScreenUpdating = True

    If UserForm1.Autosave = True Then
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С ДДИ")
    End If

ended:
    Unload UserForm1

End Sub


