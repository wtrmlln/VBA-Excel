Attribute VB_Name = "B2C_Functions"
Option Private Module
Option Explicit

'������ ������ ������������

'0.125 -> 0.3
'1.125 -> 1.3

Public Function fDostavka(NCity, TerminalTerminalPrice, ZaborPrice, ObreshetkaPrice, _
GorodOtpravki, JU, TypeZaborProschet, Quantity, _
RazgruzkaPrice, ZaborPriceProschet) As Double

    '���� ����� �� ��������, �����
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
        '���� ����� ��������, �����
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
    
    '�������� �� ������� �����. 02.12.2021
    fDostavka = fDostavka * 1.12
    
End Function

Public Function fTarif(NRowProschet, CityName, Optional VolumetricWeight = 0, Optional IsKresla = 0) As Double
            
    '���� ��������� �� ��� ���������� ��� DDI �����
    If IsKresla = 0 Then
        
        Dim NRowTarifTerminal As Long
        NRowTarifTerminal = DictTarifOMAXT.Item(ProschetArray(NRowProschet, 6) & CityName)
        
        '���� �������� ��� = 0, �����
        If VolumetricWeight = 0 Then
            'fTarif = �� ����� "�����(omaxT)" �� ����������������&�������������� �� ������������ ��������
            fTarif = TarifOMAXTArray(NRowTarifTerminal, 3)
        Else
            '���� �������� ��� <= 100,
            'fTarif = �� 100
            If VolumetricWeight <= 100 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 4)
                '���� �������� ��� <= 200,
                'fTarif = �� 200
            ElseIf VolumetricWeight <= 200 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 5)
                '���� �������� ��� <= 800,
                'fTarif = �� 800
            ElseIf VolumetricWeight <= 800 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 6)
                '���� �������� ��� <= 1500,
                'fTarif = �� 1500
            ElseIf VolumetricWeight <= 1500 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 7)
                '���� �������� ��� <= 3000,
                'fTarif = �� 3000
            ElseIf VolumetricWeight <= 3000 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 8)
                '���� �������� ��� > 3000,
                'fTarif = ����� 3000
            ElseIf VolumetricWeight > 3000 Then
                fTarif = TarifOMAXTArray(NRowTarifTerminal, 9)
            End If

        End If
    
        '���� ���� ��������� ��� ���������� ��� DDI, �����
        '������� ����������, fTarif ������� �� ����� omax�
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

'������ �������� �� ������� �������
Public Function fTPS(Weight As Double, Quantity As Long) As Single

    '���� ��� � �������� > 100, � ���-�� ���� > 1, �����
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



'����� �������
Public Function fMargin(OptPriceProschet, Quantity, smesh, Promo, marmas) As Double
    'Margin = fMargin(smkat, NRowsArray(d), 0, a)

    '���� �����,
    '��� < 10 ���. - ������ �������
    '��� < 20 ���. - ������ �������
    '��� >= 200 ���. - ������ �������
    '���� ���-�� ���� > 1:
    '(��� * ���-�� ����) < 10 ���. - ������ �������
    '(��� * ���-�� ����) < 20 ���. - ������ �������
    '(��� * ���-�� ����) >= 20 ���. - ������ �������
    '���� �� �����,
    '��� < 15 ���. - ������ �������
    '��� < 50 ���. - ������ �������
    '��� < 200 ���. - ������ �������
    '��� >= 200 ���. - ��������� �������
    '���� ���-�� ���� > 1: ���������� ������� �� ���-�� ����

    'smesh = smkat
    'stroka = NRowsArray(d)
    'promo = 0
    'marmas = a = �������
        
    '���� promo = 1, �����
    
    If OptPriceProschet <= 0 Then
        fMargin = 0
        Exit Function
    End If
    
    If Promo = 1 Then
        
        '���� ���-�� ���� > 1, �����
        If Quantity > 1 Then
            
            '���� ��� �������� * ���-�� ���� < 10 ���., �����
            'fmargin = ������ �������
            If OptPriceProschet * Quantity < 10000 Then
                fMargin = marmas(0)
            
                '���� ��� �������� * ���-�� ���� < 20 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet * Quantity < 20000 Then
                fMargin = marmas(1)
            
                '���� ��� �������� * ���-�� ���� >= 200 ���.,
                'fmargin = ������ �������
            ElseIf OptPriceProschet * Quantity >= 20000 Then
                fMargin = marmas(2)
            End If
        Else
            '���� ��� �������� * ���-�� ���� < 10 ���., �����
            'fmargin = ������ �������
            If OptPriceProschet < 10000 Then
                fMargin = marmas(0)
            
                '���� ��� �������� * ���-�� ���� < 20 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet < 20000 Then
                fMargin = marmas(1)
            
                '���� ��� �������� * ���-�� ���� >= 20 ���.,
                'fmargin = ������ �������
            ElseIf OptPriceProschet >= 20000 Then
                fMargin = marmas(2)
            End If
        End If
    Else
        '���� ���-�� ���� > 1, �����
        If Quantity > 1 Then
            
            '���� ��� �������� * ���-�� ���� < 15 ���., �����
            'fmargin = ������ �������
            If OptPriceProschet * Quantity < 15000 Then
                fMargin = marmas(0 + smesh)
            
                '���� ��� �������� * ���-�� ���� < 50 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet * Quantity < 50000 Then
                fMargin = marmas(1 + smesh)
            
                '���� ��� �������� * ���-�� ���� < 200 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet * Quantity < 200000 Then
                fMargin = marmas(2 + smesh)
            
                '���� ��� �������� * ���-�� ���� >= 200 ���., �����
                'fmargin = ��������� �������
            ElseIf OptPriceProschet * Quantity >= 200000 Then
                fMargin = marmas(3 + smesh)
            End If
        
            '� ���� ������
        Else

            '���� ��� �������� < 15 ���., �����
            'fmargin = ������ �������
            If OptPriceProschet < 15000 Then
                fMargin = marmas(0 + smesh)
            
                '���� ��� �������� < 50 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet < 50000 Then
                fMargin = marmas(1 + smesh)
            
                '���� ��� �������� < 200 ���., �����
                'fmargin = ������ �������
            ElseIf OptPriceProschet < 200000 Then
                fMargin = marmas(2 + smesh)
            
                '���� ��� �������� >= 200 ���., �����
                'fmargin = ��������� �������
            ElseIf OptPriceProschet >= 200000 Then
                fMargin = marmas(3 + smesh)
            
            End If
        End If
    End If
End Function
'��������� ���� 1� �� ��������� �� �����
Public Function SaleFind(kody) As Boolean

    '������� ��� ������ ���� "�����!":

    '�.
    '1. ��������� ���� ������ �������,
    '2. ��������� ��� �� 0 � �� ������ ��������
    '3. ��������� ��� �� ����� ��� ��������

    '�.
    '1. ��������� ���� ������ �������,
    '2. ��������� ��� �� ����� ��� ��������
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
                
        '���� ��� 1� ���������� � ����� "�����", �����
        If DictAkcia.exists(Trim(kody(i))) = True Then

            DateAkcia = CDate(AkciaArray(DictAkcia(Trim(kody(i))), 4))
            RetailPriceAkcia = AkciaArray(DictAkcia(Trim(kody(i))), 5)
            OptPriceAkcia = AkciaArray(DictAkcia(Trim(kody(i))), 6)

            If Date < DateAkcia Then
                If RetailPriceAkcia <> 0 Then
                    If RetailPriceAkcia <> RetailPriceProschet Then
                        MsgBox "�����! " & kody(i)
                        SaleFind = True
                        Exit Function
                    End If
                ElseIf OptPriceAkcia <> OptPriceProschet Then
                    MsgBox "�����! " & kody(i)
                    SaleFind = True
                    Exit Function
                End If
            End If
        End If
    Next i

End Function


