Attribute VB_Name = "B2C_Киска"
Option Private Module
Option Explicit

'Прогрузка Киска Поставщик счастья
Public Sub Прогрузка_Киска(ByRef KiskaArray, ByRef TotalRowsKiska)
    
    Dim ProgruzkaFilename As String
    Dim NCity As Long
    Dim d As Long
    Dim NRowProgruzka As Long
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim a
    Dim OptPrice
    Dim mvvodin()
    Dim CitiesArray() As String
    Dim mtarif(5, 250)
    Dim ProgruzkaSheetname As String
    Dim NcolumnProgruzka As Long
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim NRowTarifTerminal As Long
    Dim NRowTarifPickpoint As Long
    Dim Tarif As Double
    Dim Postavshik As String
    Dim mminimal() As Double

    Dim mpickpoint As Variant
    Dim PlatniiVes
    Dim TypeZabor
    Dim mpickpointdict
    Dim i As Long
    Dim tps As Double

    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    For i = UBound(KiskaArray, 2) To 0 Step -1
        If KiskaArray(0, i) = 0 Or KiskaArray(0, i) = vbNullString Then
            ReDim Preserve KiskaArray(3, UBound(KiskaArray, 2) - 1)
        End If
    Next
    
    ReDim ProgruzkaMainArray(1 To TotalRowsKiska + 1, 1 To 139)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С Киска")
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Киска")
    
    ReDim mminimal(LBound(CitiesArray) To UBound(CitiesArray))
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

    mpickpoint = Workbooks(ProschetWorkbookName).Worksheets("Pickpoint").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("Pickpoint").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set mpickpointdict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(mpickpoint, 1)

        If Trim(mpickpoint(i, 1)) <> vbNullString Then

            If mpickpointdict.exists(Trim(mpickpoint(i, 1))) = True Then
                MsgBox "Дубль в тарифах(Pickpoint) " & mpickpoint(i, 1)
                Exit Sub
            End If

            mpickpointdict.Add Trim(mpickpoint(i, 1)), i

        End If

    Next i

Kod1C_Change:
    
    NRowProschet = KiskaArray(2, d)
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
    
    'Замена ProschetArray(NRowsArray(d) на переменные:
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
    Else
        ZaborPrice = ZaborPriceProschet
    End If
    
    ReDim mminimal(LBound(CitiesArray) To UBound(CitiesArray))
    If Volume * 200 > 15 Or Weight > 15 Then

        TypeZabor = FillInfocolumnTypeZabor("1116", 200, RazgruzkaPrice, NRowProschet, True, tps)
        If TypeZabor = "Err" Then Exit Sub
    
        For NCity = LBound(CitiesArray) To UBound(CitiesArray)
    
            NRowTarifTerminal = DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity))
            mminimal(NCity) = TarifOMAXTArray(NRowTarifTerminal, 3)
            mtarif(0, NCity) = TarifOMAXTArray(NRowTarifTerminal, 4)
            mtarif(1, NCity) = TarifOMAXTArray(NRowTarifTerminal, 5)
            mtarif(2, NCity) = TarifOMAXTArray(NRowTarifTerminal, 6)
            mtarif(3, NCity) = TarifOMAXTArray(NRowTarifTerminal, 7)
            mtarif(4, NCity) = TarifOMAXTArray(NRowTarifTerminal, 8)
            mtarif(5, NCity) = TarifOMAXTArray(NRowTarifTerminal, 9)

        Next NCity

    Else
        TypeZabor = FillInfocolumnTypeZabor("1116", "Pickpoint", RazgruzkaPrice, NRowProschet, True, tps, 2)
        If TypeZabor = "Err" Then Exit Sub
    
        For NCity = LBound(CitiesArray) To UBound(CitiesArray)
    
            NRowTarifPickpoint = mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity))
            mminimal(NCity) = mpickpoint(NRowTarifPickpoint, 3)
            mtarif(0, NCity) = mpickpoint(NRowTarifPickpoint, 4)
            mtarif(1, NCity) = mpickpoint(NRowTarifPickpoint, 5)
            mtarif(2, NCity) = mpickpoint(NRowTarifPickpoint, 6)
            mtarif(3, NCity) = mpickpoint(NRowTarifPickpoint, 7)
            mtarif(4, NCity) = mpickpoint(NRowTarifPickpoint, 8)
            mtarif(5, NCity) = mpickpoint(NRowTarifPickpoint, 9)
    
        Next NCity

    End If
    
    Margin = fMargin(OptPriceProschet, Quantity, 0, 0, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = a(4)
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 129) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 130) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, False)
    ProgruzkaMainArray(NRowProgruzka, 131) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 132) = Volume
    ProgruzkaMainArray(NRowProgruzka, 133) = Weight
    ProgruzkaMainArray(NRowProgruzka, 134) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 135) = FillInfocolumnJU("Киска", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 136) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 137) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 138) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 139) = Artikul
       
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
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        Else
            Price = OptPriceProschet * Margin
        End If

        If VolumetricWeight <= 100 Then
            Tarif = mtarif(0, NCity)
        ElseIf VolumetricWeight <= 200 Then
            Tarif = mtarif(1, NCity)
        ElseIf VolumetricWeight <= 800 Then
            Tarif = mtarif(2, NCity)
        ElseIf VolumetricWeight <= 1500 Then
            Tarif = mtarif(3, NCity)
        ElseIf VolumetricWeight <= 3000 Then
            Tarif = mtarif(4, NCity)
        ElseIf VolumetricWeight > 3000 Then
            Tarif = mtarif(5, NCity)
        End If

        If NCity >= 123 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        Else
            PlatniiVes = VolumetricWeight
        End If

        If Volume * 200 > 15 Or Weight > 15 Then
            If PlatniiVes * _
            Tarif < mminimal(NCity) Then
                TerminalTerminalPrice = mminimal(NCity)
            Else
                TerminalTerminalPrice = Tarif * PlatniiVes
            End If
        Else
            If PlatniiVes <= 3 And NCity < 123 Then
                TerminalTerminalPrice = mminimal(NCity)
            ElseIf NCity < 123 Then
                TerminalTerminalPrice = mtarif(0, NCity) * (PlatniiVes - 3) + mminimal(NCity)
            Else
                If PlatniiVes * Tarif < mminimal(NCity) Then
                    TerminalTerminalPrice = mminimal(NCity)
                Else
                    TerminalTerminalPrice = Tarif * VolumetricWeight
                End If
            End If
        End If
        
        If NCity = 44 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If
        
        If Volume * 200 > 15 And NCity < 123 _
        Or Weight > 15 And NCity < 123 Then
            FullDostavkaPrice = TerminalTerminalPrice + ZaborPrice + RazgruzkaPrice + tps
        ElseIf NCity >= 123 Then
            FullDostavkaPrice = TerminalTerminalPrice + ZaborPriceProschet + RazgruzkaPrice + tps
        Else
            FullDostavkaPrice = TerminalTerminalPrice + RazgruzkaPrice + tps
        End If

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

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127
            If NcolumnProgruzka < 108 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
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

    
    Call fMixer(ProgruzkaMainArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg

federal_orenburg:

    ProgruzkaMainArray(NRowProgruzka, 128) = ProgruzkaMainArray(NRowProgruzka, 59)
    GoTo orenburg_infocolumn

    If Komplektuyeshee = "К" Then
        If OptPrice > 7500 Then
            ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(OptPriceProschet * 1.9, -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(OptPriceProschet * 1.9, 0)
        End If
    ElseIf Komplektuyeshee = "К1" Then
        If OptPrice > 7500 Then
            ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(OptPriceProschet * 1.6, -1)
        Else
            ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(OptPriceProschet * 1.6, 0)
        End If
    ElseIf OptPrice > 7500 Then
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(((OptPriceProschet * (Margin - 0.1) + RazgruzkaPrice) / 0.96), -1)
    Else
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(((OptPriceProschet * (Margin - 0.1) + RazgruzkaPrice) / 0.96), 0)
    End If

orenburg_infocolumn:

    If UserForm1.CheckBox2 = False And OptPrice > 7500 And ProgruzkaMainArray(NRowProgruzka, 128) < RetailPrice Then
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(RetailPrice, -1)
    ElseIf UserForm1.CheckBox2 = False And ProgruzkaMainArray(NRowProgruzka, 128) < RetailPrice Then
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(RetailPrice, 0)
    End If

Change_Row_Progruzka:
    Postavshik = ProschetArray(NRowProschet, 1)

    Do While d <> UBound(KiskaArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change
    Loop

'    NCity = 20
'    NColumnProgruzka = 5
'    d = 0

    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EI" & NRowProgruzka).Value = ProgruzkaMainArray
    Application.ScreenUpdating = True

    If UserForm1.Autosave = True Then
        
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Киска")

    End If

End Sub

'Прогрузка Промо Киска Поставщик счастья
Public Sub Прогрузка_Промо_Киска(ByRef KiskaPromoArray, ByRef TotalRowsKiskaPromo)

    Dim mtarif(5, 250)
    Dim CitiesArray() As String
    Dim FullDostavkaPrice
    Dim Price
    Dim TerminalTerminalPrice
    Dim TotalPrice
    Dim Margin
    Dim OptPrice
    Dim mminimal() As Double
    Dim Postavshik As String
    Dim mpickpoint As Variant
    Dim mpickpointdict
    Dim PlatniiVes
    Dim TypeZabor As String
    Dim mvvodin() As String
    Dim NRowProgruzka As Long
    
    Dim ProgruzkaFilename As String
    Dim ProgruzkaSheetname As String
    Dim NCity As Long
    Dim NcolumnProgruzka As Long
    Dim d As Long
    Dim a() As String
    Dim NRowProschet As Long
    Dim ObreshetkaPrice As Double
    Dim ZaborPrice As Double
    Dim Tarif As Double
    Dim i As Long
    Dim tps As Double
    
    Workbooks.Add
    ProgruzkaFilename = ActiveWorkbook.name
    ProgruzkaSheetname = ActiveSheet.name

    For i = UBound(KiskaPromoArray, 2) To 0 Step -1
        If KiskaPromoArray(0, i) = 0 Or KiskaPromoArray(0, i) = vbNullString Then
            ReDim Preserve KiskaPromoArray(3, UBound(KiskaPromoArray, 2) - 1)
        End If
    Next
    
    ReDim CitiesArray(20 To 142)
    CitiesArray = CreateCities(CitiesArray, "Б2С Киска")
    
    ReDim ProgruzkaMainArray(1 To TotalRowsKiskaPromo + 1, 1 To 139)
    ProgruzkaMainArray = CreateHeader(ProgruzkaMainArray, "Б2С Киска")
    
    NCity = 20
    NcolumnProgruzka = 5
    d = 0
    NRowProgruzka = 2

    mpickpoint = Workbooks(ProschetWorkbookName).Worksheets("Pickpoint").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("Pickpoint").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set mpickpointdict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(mpickpoint, 1)

        If Trim(mpickpoint(i, 1)) <> vbNullString Then

            If mpickpointdict.exists(Trim(mpickpoint(i, 1))) = True Then
                MsgBox "Дубль в тарифах(Pickpoint) " & mpickpoint(i, 1)
                Exit Sub
            End If

            mpickpointdict.Add Trim(mpickpoint(i, 1)), i

        End If

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

Kod1C_Change:
    
    NRowProschet = KiskaPromoArray(2, d)
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

    If DictLgotniiGorod.exists(GorodOtpravki) = True And ZaborPriceProschet <> 0 Then
        ZaborPrice = 0
    Else
        ZaborPrice = ZaborPriceProschet
    End If
    
    ReDim mminimal(LBound(CitiesArray) To UBound(CitiesArray))
    If Volume * 200 > 15 Or Weight > 15 Then

        TypeZabor = FillInfocolumnTypeZabor("1116", 0, RazgruzkaPrice, NRowProschet, True, tps)
        If TypeZabor = "Err" Then Exit Sub

        For NCity = LBound(CitiesArray) To UBound(CitiesArray)
            mtarif(0, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 4)
            mtarif(1, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 5)
            mtarif(2, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 6)
            mtarif(3, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 7)
            mtarif(4, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 8)
            mtarif(5, NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 9)
            mminimal(NCity) = TarifOMAXTArray(DictTarifOMAXT.Item(GorodOtpravki & CitiesArray(NCity)), 3)
        Next NCity

    Else

        TypeZabor = FillInfocolumnTypeZabor("1116", "Pickpoint", RazgruzkaPrice, NRowProschet, True, tps, 2)
        If TypeZabor = "Err" Then Exit Sub

        For NCity = LBound(CitiesArray) To UBound(CitiesArray)
            mtarif(0, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 4)
            mtarif(1, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 5)
            mtarif(2, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 6)
            mtarif(3, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 7)
            mtarif(4, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 8)
            mtarif(5, NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 9)
            mminimal(NCity) = mpickpoint(mpickpointdict.Item(GorodOtpravki & CitiesArray(NCity)), 3)
        Next NCity

    End If
    
    Margin = fMargin(OptPriceProschet, Quantity, 0, 1, a)

    ProgruzkaMainArray(NRowProgruzka, 1) = Artikul
    ProgruzkaMainArray(NRowProgruzka, 2) = "1116"
    ProgruzkaMainArray(NRowProgruzka, 3) = Kod1CProschet
    ProgruzkaMainArray(NRowProgruzka, 4) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 129) = GorodOtpravki
    ProgruzkaMainArray(NRowProgruzka, 130) = FillInfocolumnMargin(Komplektuyeshee, Margin, False, True)
    ProgruzkaMainArray(NRowProgruzka, 131) = TypeZabor
    ProgruzkaMainArray(NRowProgruzka, 132) = Volume
    ProgruzkaMainArray(NRowProgruzka, 133) = Weight
    ProgruzkaMainArray(NRowProgruzka, 134) = Quantity
    ProgruzkaMainArray(NRowProgruzka, 135) = FillInfocolumnJU("Киска", JU, Komplektuyeshee, Quantity, GorodOtpravki, ObreshetkaPrice)
    ProgruzkaMainArray(NRowProgruzka, 136) = FillInfocolumnRetailPrice(RetailPrice, RetailPriceInfo)
    ProgruzkaMainArray(NRowProgruzka, 137) = OptPriceProschet
    ProgruzkaMainArray(NRowProgruzka, 138) = ProschetArray(NRowProschet, 1)
    ProgruzkaMainArray(NRowProgruzka, 139) = Artikul
    
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
            Price = OptPriceProschet * Margin
        ElseIf NCity >= 123 Then
            Price = OptPriceProschet * (Margin + SNG_MARGIN)
        End If

        If VolumetricWeight <= 100 Then
            Tarif = mtarif(0, NCity)
        ElseIf VolumetricWeight <= 200 Then
            Tarif = mtarif(1, NCity)
        ElseIf VolumetricWeight <= 800 Then
            Tarif = mtarif(2, NCity)
        ElseIf VolumetricWeight <= 1500 Then
            Tarif = mtarif(3, NCity)
        ElseIf VolumetricWeight <= 3000 Then
            Tarif = mtarif(4, NCity)
        ElseIf VolumetricWeight > 3000 Then
            Tarif = mtarif(5, NCity)
        End If

        If NCity >= 123 Then
            If Volume * 250 > Weight Then
                PlatniiVes = Volume * 250
            Else
                PlatniiVes = Weight
            End If
        Else
            PlatniiVes = VolumetricWeight
        End If

        If Volume * 200 > 15 Or Weight > 15 Then
            If PlatniiVes * Tarif < mminimal(NCity) Then
                TerminalTerminalPrice = mminimal(NCity)
            Else
                TerminalTerminalPrice = Tarif * PlatniiVes
            End If
        Else

            If PlatniiVes <= 3 And NCity < 123 Then
                TerminalTerminalPrice = mminimal(NCity)
            ElseIf NCity < 123 Then
                TerminalTerminalPrice = mtarif(0, NCity) * (PlatniiVes - 3) + mminimal(NCity)
            Else
                If VolumetricWeight * Tarif < mminimal(NCity) Then
                    TerminalTerminalPrice = mminimal(NCity)
                Else
                    TerminalTerminalPrice = Tarif * VolumetricWeight
                End If
            End If
        End If
        
        If NCity = 44 Then
            TerminalTerminalPrice = TerminalTerminalPrice + 500
        End If
        
        If Volume * 200 > 15 And NCity < 123 Or Weight > 15 And NCity < 123 Then
            FullDostavkaPrice = TerminalTerminalPrice + ZaborPrice + RazgruzkaPrice + tps
        ElseIf NCity >= 123 Then
            FullDostavkaPrice = TerminalTerminalPrice + ZaborPriceProschet + RazgruzkaPrice + tps
        Else
            FullDostavkaPrice = TerminalTerminalPrice + RazgruzkaPrice + tps
        End If

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

        ProgruzkaMainArray(NRowProgruzka, NcolumnProgruzka) = TotalPrice
        NcolumnProgruzka = NcolumnProgruzka + 1

    Next

    If UserForm1.CheckBox1 = True Then
        GoTo federal_orenburg
    End If

    If Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Then

        For NcolumnProgruzka = 5 To 127
            
            If NcolumnProgruzka < 108 Then
                If ProgruzkaMainArray(NRowProgruzka, 4) > 7500 Then
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

    
    Call fMixer(ProgruzkaMainArray, UBound(CitiesArray), NRowProgruzka)
    GoTo federal_orenburg
federal_orenburg:

    ProgruzkaMainArray(NRowProgruzka, 128) = ProgruzkaMainArray(NRowProgruzka, 59)
    GoTo orenburg_infocolumn

orenburg_infocolumn:
    
    If UserForm1.CheckBox2 = False And OptPrice > 7500 And ProgruzkaMainArray(NRowProgruzka, 128) < RetailPrice Then
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(RetailPrice, -1)
    ElseIf UserForm1.CheckBox2 = False And ProgruzkaMainArray(NRowProgruzka, 128) < RetailPrice Then
        ProgruzkaMainArray(NRowProgruzka, 128) = Round_Up(RetailPrice, 0)
    End If

Change_Row_Progruzka:

    Postavshik = ProschetArray(NRowProschet, 1)
    
    Do While d <> UBound(KiskaPromoArray, 2)
        d = d + 1
        NCity = 20
        NcolumnProgruzka = 5
        NRowProgruzka = NRowProgruzka + 1
        GoTo Kod1C_Change
    Loop

    Application.ScreenUpdating = False
    Workbooks(ProgruzkaFilename).Worksheets(ProgruzkaSheetname).Range("A1:EI" & TotalRowsKiskaPromo + 1).Value = ProgruzkaMainArray
    Application.ScreenUpdating = True

    If UserForm1.Autosave = True Then
        Call Autosave(Postavshik, ProgruzkaFilename, "Б2С Киска Промо")
    End If

ended:
End Sub
