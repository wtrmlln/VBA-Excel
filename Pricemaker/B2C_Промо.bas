Attribute VB_Name = "B2C_Промо"
Option Private Module
Option Explicit

Public Sub Подготовка_Промо_Общее()

    Dim r1 As Long
    Dim NRowProgruzka As Long
    Dim d As Long
    Dim a() As String
    Dim mvvod()
    Dim NRowsArray()
    Dim mvvodin() As String
    Dim str As String
    Dim smkat As Long
    Dim nazenki
    Dim KomponentNumber As String
    Dim mut
    
    Dim SOTPromoArray()
    Dim PromoArray() As Variant
    Dim TeplozharPromoArray()
    Dim KompkreslaPromoArray()
    Dim KiskaPromoArray()
    Dim PMKPromoArray() As Variant
    
    Dim TotalRowsPromo As Long
    Dim promo_bound As Long
    
    Dim TriggerTorudaPromo As Long                   'tП
    Dim TotalRowsTeplozharPromo As Long                      '620П
    Dim TotalRowsKompkreslaPromo As Long                     'НП
    Dim TotalRowsKiskaPromo As Long                    '1196П
    Dim TotalRowsSOTPromo As Long                 '?
    Dim TotalRowsPMKPromo As Long
    Dim i As Long

    d = 0
    smkat = 0
    NRowProgruzka = 2
    kat = UserForm1.TextBox3

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
    ReDim PromoArray(4, 100, UBound(mvvod))

    For d = 0 To UBound(mvvod)
        mut = 0

        str = Trim(mvvod(LBound(mvvod) + d))
        r1 = Proschet_Dict.Item(str)

        nazenki = Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(r1, 7).Value

        If InStr(Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(r1, 7).Value, Trim(kat) & "t") <> 0 Then
            mut = 1
            nazenki = Replace(nazenki, Trim(kat) & "t", "!")
        End If

        nazenki = Replace(nazenki, Trim(kat), Trim(kat) & "П")

        Windows(ProschetWorkbookName).Activate
        Range(Cells(r1, 1), Cells(r1, 19)).Select
    
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
        End With
    
        Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(r1, 7).Value = nazenki

        If mut = 1 Then
            Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells(r1, 7).Value = Replace(nazenki, "!", Trim(kat) & "t")
        End If

        NRowsArray(d) = r1
    Next

    If InStr(kat, "П") = 0 Then
        kat = kat & "П"
    End If

    d = 0

    ProschetArray = Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Range("A1:T" & _
    Workbooks(ProschetWorkbookName).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

Kod1C_Change:

    a = Split(ProschetArray(NRowsArray(d), 7), ";")

    ReDim mvvodin(UBound(a))

    For i = 0 To UBound(a)
        If Len(a(i)) = 0 Then
            ReDim Preserve mvvodin(UBound(mvvodin) - 1)
        Else
            mvvodin(i - (UBound(a) - UBound(mvvodin))) = a(i)
        End If
    Next

    If (UBound(mvvodin) + 1) Mod 5 <> 0 Then
        MsgBox "Некорректная наценка/категория"
        GoTo ended
    End If


Komponent_Change:

    Do While a(4 + smkat) <> kat
        GoTo Change_Row_Progruzka
    Loop

    str = a(4 + smkat)
    KomponentNumber = Val(a(4 + smkat))

    If KomponentNumber <> str And str <> "Н" And InStr(str, "П") = 0 And InStr(str, "t") = 0 Then
        MsgBox "Некорректная наценка/категория " & ProschetArray(NRowsArray(d), 17) & " " & str
        GoTo ended
    End If

    If InStr(str, "tП") <> 0 Then
        TriggerTorudaPromo = TriggerTorudaPromo + 1
    End If

    If InStr(str, "П") <> 0 Then
        PromoArray(0, TotalRowsPromo, d) = NRowProgruzka - smkat / 5 + 2
        PromoArray(1, TotalRowsPromo, d) = UBound(mvvodin) + 1
        PromoArray(2, TotalRowsPromo, d) = NRowsArray(d)
        PromoArray(3, TotalRowsPromo, d) = smkat
        PromoArray(4, TotalRowsPromo, d) = a(4 + smkat)
        TotalRowsPromo = TotalRowsPromo + 1
        promo_bound = promo_bound + 1
        GoTo Change_Row_Progruzka
    End If

Change_Row_Progruzka:

    Do While smkat <> ((UBound(mvvodin) + 1) - 5)
        smkat = smkat + 5
        GoTo Komponent_Change
    Loop

    Do While d <> UBound(NRowsArray)
        d = d + 1
        smkat = 0
        TotalRowsPromo = 0
        GoTo Kod1C_Change
    Loop

    d = 0
    smkat = 0
    
    ReDim KiskaPromoArray(3, promo_bound)
    ReDim KompkreslaPromoArray(3, promo_bound)
    ReDim PMKPromoArray(4, promo_bound) As Variant
    ReDim TeplozharPromoArray(3, promo_bound)
    ReDim SOTPromoArray(4, promo_bound)
    
    If TriggerTorudaPromo <> 0 Then
        UserForm4.Show
    End If

    If promo_bound <> 0 Then
        If TriggerTorudaPromo <> promo_bound Then
            UserForm2.Show
        End If
        For i = 0 To UBound(mvvod)
            For TotalRowsPromo = 0 To UBound(PromoArray, 2)
                'Если компонент - "Киска", тогда запустится Прогрузка_Промо_Киска
                If PromoArray(4, TotalRowsPromo, i) = "1116П" Then
                    KiskaPromoArray(0, TotalRowsKiskaPromo) = PromoArray(0, TotalRowsPromo, i)
                    KiskaPromoArray(1, TotalRowsKiskaPromo) = PromoArray(1, TotalRowsPromo, i)
                    KiskaPromoArray(2, TotalRowsKiskaPromo) = PromoArray(2, TotalRowsPromo, i)
                    KiskaPromoArray(3, TotalRowsKiskaPromo) = PromoArray(3, TotalRowsPromo, i)
                    TotalRowsKiskaPromo = TotalRowsKiskaPromo + 1
                    'ReDim Preserve KiskaPromoArray(3, UBound(KiskaPromoArray, 2) + 1)
                    'g1 = g1 + 1
                    'Если компонент - Н, тогда запустится Прогрузка_Промо_Компкресла
                ElseIf PromoArray(4, TotalRowsPromo, i) = "НП" Then
                    KompkreslaPromoArray(0, TotalRowsKompkreslaPromo) = PromoArray(0, TotalRowsPromo, i)
                    KompkreslaPromoArray(1, TotalRowsKompkreslaPromo) = PromoArray(1, TotalRowsPromo, i)
                    KompkreslaPromoArray(2, TotalRowsKompkreslaPromo) = PromoArray(2, TotalRowsPromo, i)
                    KompkreslaPromoArray(3, TotalRowsKompkreslaPromo) = PromoArray(3, TotalRowsPromo, i)
                    TotalRowsKompkreslaPromo = TotalRowsKompkreslaPromo + 1
                    'ReDim Preserve KompkreslaPromoArray(3, UBound(KompkreslaPromoArray, 2) + 1)
                    'i1 = i1 + 1
                    'Если компонент - "ПМК", тогда запустится Прогрузка_Промо_ПМК
                ElseIf PromoArray(4, TotalRowsPromo, i) = "796П" Or PromoArray(4, TotalRowsPromo, i) = "1034П" Then
                    PMKPromoArray(0, TotalRowsPMKPromo) = PromoArray(0, TotalRowsPromo, i)
                    PMKPromoArray(1, TotalRowsPMKPromo) = PromoArray(1, TotalRowsPromo, i)
                    PMKPromoArray(2, TotalRowsPMKPromo) = PromoArray(2, TotalRowsPromo, i)
                    PMKPromoArray(3, TotalRowsPMKPromo) = PromoArray(3, TotalRowsPromo, i)
                    PMKPromoArray(4, TotalRowsPMKPromo) = PromoArray(4, TotalRowsPromo, i)
                    TotalRowsPMKPromo = TotalRowsPMKPromo + 1
                    'ReDim Preserve PMKPromoArray(4, UBound(PMKPromoArray, 2) + 1)
                    'pmk1 = pmk1 + 1
                    'Если компонент - "Тепложар", тогда запустится Прогрузка_Промо_Тепложар
                ElseIf PromoArray(4, TotalRowsPromo, i) = "620П" Or PromoArray(4, TotalRowsPromo, i) = "620tП" Then
                    TeplozharPromoArray(0, TotalRowsTeplozharPromo) = PromoArray(0, TotalRowsPromo, i)
                    TeplozharPromoArray(1, TotalRowsTeplozharPromo) = PromoArray(1, TotalRowsPromo, i)
                    TeplozharPromoArray(2, TotalRowsTeplozharPromo) = PromoArray(2, TotalRowsPromo, i)
                    TeplozharPromoArray(3, TotalRowsTeplozharPromo) = PromoArray(3, TotalRowsPromo, i)
                    TotalRowsTeplozharPromo = TotalRowsTeplozharPromo + 1
                    'ReDim Preserve TeplozharPromoArray(3, UBound(TeplozharPromoArray, 2) + 1)
                    'b1 = b1 + 1
                    'Если компонент - любой другой, тогда запустится Прогрузка_Промо_СОТ_Торуда
                ElseIf Len(PromoArray(4, TotalRowsPromo, i)) <> 0 Then
                    SOTPromoArray(0, TotalRowsSOTPromo) = PromoArray(0, TotalRowsPromo, i)
                    SOTPromoArray(1, TotalRowsSOTPromo) = PromoArray(1, TotalRowsPromo, i)
                    SOTPromoArray(2, TotalRowsSOTPromo) = PromoArray(2, TotalRowsPromo, i)
                    SOTPromoArray(3, TotalRowsSOTPromo) = PromoArray(3, TotalRowsPromo, i)
                    SOTPromoArray(4, TotalRowsSOTPromo) = PromoArray(4, TotalRowsPromo, i)
                    TotalRowsSOTPromo = TotalRowsSOTPromo + 1
                    'ReDim Preserve SOTPromoArray(4, UBound(PromoArray, 2) + 1)
                End If
            Next TotalRowsPromo
        Next i
    End If

    If TotalRowsPMKPromo <> 0 Then
        Call Прогрузка_Промо_ПМК(PMKPromoArray, TotalRowsPMKPromo)
        
    ElseIf TotalRowsTeplozharPromo <> 0 Then
        Call Прогрузка_Промо_Тепложар(TeplozharPromoArray, TotalRowsTeplozharPromo)
        
    ElseIf TotalRowsKompkreslaPromo <> 0 Then
        Call Прогрузка_Промо_Компкресла(KompkreslaPromoArray, TotalRowsKompkreslaPromo)
        
    ElseIf TotalRowsKiskaPromo <> 0 Then
        Call Прогрузка_Промо_Киска(KiskaPromoArray, TotalRowsKiskaPromo)

    ElseIf TotalRowsSOTPromo <> 0 Then
        Call Прогрузка_Промо_СОТ_Торуда(SOTPromoArray, TriggerTorudaPromo)
    End If

ended:
    Unload UserForm1

End Sub
