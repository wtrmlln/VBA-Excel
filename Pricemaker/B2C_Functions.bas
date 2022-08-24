Attribute VB_Name = "B2C_Functions"
Option Private Module
Option Explicit

'Расчет полных транспортных

'0.125 -> 0.3
'1.125 -> 1.3

Public Function fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, _
GorodOtpravki, JU, TypeZaborProschet, Quantity, _
RazgruzkaPrice, ZaborPriceProschet) As Double

    'Если город не льготный, тогда
    If DictLgotniiGorod.exists(GorodOtpravki) = False Then
        If NCity < 123 Or NCity >= 143 And NCity < 163 Then
            If JU <> 0 Then
                fDostavka = (TerminalTerminalPrice * 1.3) + ZaborPrice + ObreshetkaPrice / 2 + RazgruzkaPrice
            Else
                fDostavka = TerminalTerminalPrice + ZaborPrice + RazgruzkaPrice
            End If
        ElseIf NCity >= 123 And NCity < 143 Or NCity >= 163 Then
            If JU <> 0 Then
                fDostavka = (TerminalTerminalPrice * 1.3) + ZaborPriceProschet + ObreshetkaPrice / 2 + RazgruzkaPrice
            Else
                fDostavka = TerminalTerminalPrice + ZaborPriceProschet + RazgruzkaPrice
            End If
        End If
        'Если город льготный, тогда
    ElseIf DictLgotniiGorod.exists(GorodOtpravki) = True Then
        If NCity < 123 Or NCity >= 143 And NCity < 163 Then
            If JU <> 0 Then
                fDostavka = (TerminalTerminalPrice * 1.3) + ZaborPrice + ObreshetkaPrice / 2 + RazgruzkaPrice
            Else
                fDostavka = TerminalTerminalPrice + ZaborPrice + RazgruzkaPrice
            End If
        ElseIf NCity >= 123 And NCity < 143 Or NCity >= 163 Then
            If JU <> 0 Then
                fDostavka = (TerminalTerminalPrice * 1.3) + ZaborPriceProschet + ObreshetkaPrice / 2 + RazgruzkaPrice
            Else
                fDostavka = TerminalTerminalPrice + ZaborPriceProschet + RazgruzkaPrice
            End If
        End If
    End If
    
    fDostavka = fDostavka / Quantity
    
    'Надбавка за высокий сезон. 02.12.2021
    fDostavka = fDostavka * 1.12
    
End Function

Public Function fTarif(NRowProschet, CityName, Optional VolumetricWeight = 0, Optional IsKresla = 0) As Double
            
    'Если прогрузка не для компкресел или DDI тогда
    If IsKresla = 0 Then
        
        Dim NRowTarifTerminal As Long
        NRowTarifTerminal = DictTarifOMAXT.Item(ProschetArray(NRowProschet, 6) & CityName)
        
        'Если объемный вес = 0, тогда
        If VolumetricWeight = 0 Then
            'fTarif = из листа "Тариф(omaxT)" по ГородОтправления&ГородПолучения по минимальному значению
            fTarif = TarifOMAXTArray(NRowTarifTerminal, 3)
        Else
            'Если объемный вес <= 100,
            'fTarif = до 100
            If VolumetricWeight <= 100 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 4)
                'Если объемный вес <= 200,
                'fTarif = до 200
            ElseIf VolumetricWeight <= 200 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 5)
                'Если объемный вес <= 800,
                'fTarif = до 800
            ElseIf VolumetricWeight <= 800 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 6)
                'Если объемный вес <= 1500,
                'fTarif = до 1500
            ElseIf VolumetricWeight <= 1500 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 7)
                'Если объемный вес <= 3000,
                'fTarif = до 3000
            ElseIf VolumetricWeight <= 3000 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 8)
                'Если объемный вес > 3000,
                'fTarif = более 3000
            ElseIf VolumetricWeight > 3000 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 9)
            End If

        End If
    
        'Если Если прогрузка для компкресел или DDI, тогда
        'Условия аналогичны, fTarif берется из листа omaxД
    Else
        
        Dim NRowTarifDver As Long
        NRowTarifDver = DictTarifOMAXD.Item(ProschetArray(NRowProschet, 6) & CityName)
        
        If VolumetricWeight = 0 Then
            fTarif = TarifOMAXDArray(NRowTarifDver, 3)
        Else
            If VolumetricWeight <= 100 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 4)
            ElseIf VolumetricWeight <= 200 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 5)
            ElseIf VolumetricWeight <= 800 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 6)
            ElseIf VolumetricWeight <= 1500 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 7)
            ElseIf VolumetricWeight <= 3000 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 8)
            ElseIf VolumetricWeight > 3000 Then
                fTarif = TarifOMAXDArray(NRowTarifDver, 9)
            End If
        End If
    End If
End Function

'Расчет надбавки за тяжелую посылку
Public Function fTPS(Weight As Double, Quantity As Long) As Single

    'Если вес в просчете > 100, и Кол-во штук > 1, тогда
    If Weight > 100 Then
        If Quantity > 1 Then
            fTPS = Weight * 3 / Quantity
        Else
            fTPS = Weight * 3
        End If
    Else
        fTPS = 0
    End If

End Function



'Выбор наценки
Public Function fMargin(OptPriceProschet, Quantity, smesh, Promo, marmas) As Double
    'Margin = fMargin(smkat, NRowsArray(d), 0, a)

    'Если промо,
    'ОПТ < 10 тыс. - первая наценка
    'ОПТ < 20 тыс. - вторая наценка
    'ОПТ >= 200 тыс. - третья наценка
    'Если кол-во штук > 1:
    '(ОПТ * Кол-во штук) < 10 тыс. - первая наценка
    '(ОПТ * Кол-во штук) < 20 тыс. - вторая наценка
    '(ОПТ * Кол-во штук) >= 20 тыс. - третья наценка
    'Если не промо,
    'ОПТ < 15 тыс. - первая наценка
    'ОПТ < 50 тыс. - вторая наценка
    'ОПТ < 200 тыс. - третья наценка
    'ОПТ >= 200 тыс. - четвертая наценка
    'Если кол-во штук > 1: аналогично умножая на кол-во штук

    'smesh = smkat
    'stroka = NRowsArray(d)
    'promo = 0
    'marmas = a = Наценка
        
    'Если promo = 1, тогда
    
    If OptPriceProschet <= 0 Then
        fMargin = 0
        Exit Function
    End If
    
    If Promo = 1 Then
        
        'Если кол-во штук > 1, тогда
        If Quantity > 1 Then
            
            'Если ОПТ просчета * Кол-во штук < 10 тыс., тогда
            'fmargin = Первая наценка
            If OptPriceProschet * Quantity < 10000 Then
                fMargin = marmas(0)
            
                'Если ОПТ просчета * Кол-во штук < 20 тыс., тогда
                'fmargin = Вторая наценка
            ElseIf OptPriceProschet * Quantity < 20000 Then
                fMargin = marmas(1)
            
                'Если ОПТ просчета * Кол-во штук >= 200 тыс.,
                'fmargin = Третья наценка
            ElseIf OptPriceProschet * Quantity >= 20000 Then
                fMargin = marmas(2)
            End If
        Else
            'Если ОПТ просчета * Кол-во штук < 10 тыс., тогда
            'fmargin = Первая наценка
            If OptPriceProschet < 10000 Then
                fMargin = marmas(0)
            
                'Если ОПТ просчета * Кол-во штук < 20 тыс., тогда
                'fmargin = Вторая наценка
            ElseIf OptPriceProschet < 20000 Then
                fMargin = marmas(1)
            
                'Если ОПТ просчета * Кол-во штук >= 20 тыс.,
                'fmargin = Третья наценка
            ElseIf OptPriceProschet >= 20000 Then
                fMargin = marmas(2)
            End If
        End If
    Else
        'Если Кол-во штук > 1, тогда
        If Quantity > 1 Then
            
            'Если ОПТ просчета * Кол-во штук < 15 тыс., тогда
            'fmargin = Первая наценка
            If OptPriceProschet * Quantity < 15000 Then
                fMargin = marmas(0 + smesh)
            
                'Если ОПТ просчета * Кол-во штук < 50 тыс., тогда
                'fmargin = Вторая наценка
            ElseIf OptPriceProschet * Quantity < 50000 Then
                fMargin = marmas(1 + smesh)
            
                'Если ОПТ просчета * Кол-во штук < 200 тыс., тогда
                'fmargin = Третья наценка
            ElseIf OptPriceProschet * Quantity < 200000 Then
                fMargin = marmas(2 + smesh)
            
                'Если ОПТ просчета * Кол-во штук >= 200 тыс., тогда
                'fmargin = Четвертая наценка
            ElseIf OptPriceProschet * Quantity >= 200000 Then
                fMargin = marmas(3 + smesh)
            End If
        
            'В ином случае
        Else

            'Если ОПТ просчета < 15 тыс., тогда
            'fmargin = Первая наценка
            If OptPriceProschet < 15000 Then
                fMargin = marmas(0 + smesh)
            
                'Если ОПТ просчета < 50 тыс., тогда
                'fmargin = Вторая наценка
            ElseIf OptPriceProschet < 50000 Then
                fMargin = marmas(1 + smesh)
            
                'Если ОПТ просчета < 200 тыс., тогда
                'fmargin = Третья наценка
            ElseIf OptPriceProschet < 200000 Then
                fMargin = marmas(2 + smesh)
            
                'Если ОПТ просчета >= 200 тыс., тогда
                'fmargin = Четвертая наценка
            ElseIf OptPriceProschet >= 200000 Then
                fMargin = marmas(3 + smesh)
            
            End If
        End If
    End If
End Function
'Проверяет коды 1С из юзерформы на акцию
Public Function SaleFind(kody) As Boolean

    'Условие для вывода окна "Акция!":

    'А.
    '1. Акционная дата больше текущей,
    '2. Акционное РРЦ не 0 и не пустое значение
    '3. Акционное РРЦ не равно РРЦ просчета

    'Б.
    '1. Акционная дата больше текущей,
    '2. Акционный ОПТ не равен ОПТ просчета
    Dim i As Long
    Dim RetailPriceProschet As Long
    Dim RetailPriceAkcia As Long
    
    Dim OptPriceProschet As Long
    Dim OptPriceAkcia As Long
    
    Dim DateAkcia As Date
    
    For i = 0 To UBound(kody)
        
        If IsNumeric(ProschetArray(Proschet_Dict(Trim(kody(i))), 18)) = False Then
            RetailPriceProschet = 0
        Else
            RetailPriceProschet = CLng(ProschetArray(Proschet_Dict(Trim(kody(i))), 18))
        End If
        
        If IsNumeric(ProschetArray(Proschet_Dict(Trim(kody(i))), 19)) = False Then
            OptPriceProschet = 0
        Else
            OptPriceProschet = CLng(ProschetArray(Proschet_Dict(Trim(kody(i))), 19))
        End If
                
        'Если код 1С существует в листе "Акция", тогда
        If DictAkcia.exists(Trim(kody(i))) = True Then

            DateAkcia = CDate(AkciaArray(DictAkcia(Trim(kody(i))), 4))
            RetailPriceAkcia = AkciaArray(DictAkcia(Trim(kody(i))), 5)
            OptPriceAkcia = AkciaArray(DictAkcia(Trim(kody(i))), 6)

            If Date < DateAkcia Then
                If RetailPriceAkcia <> 0 Then
                    If RetailPriceAkcia <> RetailPriceProschet Then
                        MsgBox "Акция! " & kody(i)
                        SaleFind = True
                        Exit Function
                    End If
                ElseIf OptPriceAkcia <> OptPriceProschet Then
                    MsgBox "Акция! " & kody(i)
                    SaleFind = True
                    Exit Function
                End If
            End If
        End If
    Next i

End Function


