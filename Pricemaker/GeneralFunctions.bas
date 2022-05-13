Attribute VB_Name = "GeneralFunctions"
Option Private Module
Option Explicit

'В активную книгу добавляет пустой лист в конец имеющихся
Public Sub CreateSheet(ByRef name As String)

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
    ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    ws.name = name
    
End Sub

Public Sub Autosave(ByRef Postavshik As String, ProgruzkaFilename As String, ByRef Site As String)

On Error GoTo ended
    
    Application.ScreenUpdating = False
    
    'B2С
    Dim NewFilename As String
    If Site = "Б2С СОТ" Then
        NewFilename = "Прогрузка"
    ElseIf Site = "Б2С Торуда" Then
        NewFilename = "Прогрузка Торуда"
    ElseIf Site = "Б2С СОТ Промо" Then
        NewFilename = "Прогрузка Промо"
    ElseIf Site = "Б2С Торуда Промо" Then
        NewFilename = "Прогрузка Торуда Промо"
    ElseIf Site = "Б2С Тепложар" Then
        NewFilename = "Прогрузка 620"
    ElseIf Site = "Б2С Тепложар Промо" Then
        NewFilename = "Прогрузка 620 Промо"
    ElseIf Site = "Б2С СОТ Снять с промо" Then
        NewFilename = "Прогрузка откл."
    ElseIf Site = "Б2С Торуда Снять с промо" Then
        NewFilename = "Прогрузка Торуда откл."
    ElseIf Site = "Б2С Компкресла" Then
        NewFilename = "Прогрузка Н"
    ElseIf Site = "Б2С Компкресла Промо" Then
        NewFilename = "Прогрузка Н Промо"
    ElseIf Site = "Б2С Киска" Then
        NewFilename = "Прогрузка Киска"
    ElseIf Site = "Б2С Киска Промо" Then
        NewFilename = "Прогрузка Киска Промо"
    ElseIf Site = "Б2С ДДИ" Then
        NewFilename = "Прогрузка DDI"
    ElseIf Site = "Б2С ПМК" Then
        NewFilename = "Прогрузка ПМК"
    ElseIf Site = "Б2С ПМК 1034" Then
        NewFilename = "Прогрузка ПМК 1034"
    ElseIf Site = "Б2С ПМК Промо" Then
        NewFilename = "Прогрузка ПМК Промо"
    ElseIf Site = "Б2С ПМК 1034 Промо" Then
        NewFilename = "Прогрузка ПМК 1034 Промо"
    'B2B
    ElseIf Site = "Б2Б ПМК" Then
        NewFilename = "Прогрузка B2B ПМК"
    ElseIf Site = "Б2Б ПМК 830/917" Then
        NewFilename = "Прогрузка B2B 830-917 ПМК"
    ElseIf Site = "Б2Б Торуда" Then
        NewFilename = "Прогрузка B2B Торуда"
    ElseIf Site = "Б2Б Кабинетоф" Then
        NewFilename = "Прогрузка B2B Кабинетоф"
    ElseIf Site = "Б2Б ПМК Ледосвет" Then
        NewFilename = "Прогрузка B2B ПМК Ледосвет"
    Else
        MsgBox "Неправильно указан аргумент Site в функции Autosave"
        Exit Sub
    End If
    
    Dim Number As Long
    Do While Dir(Environ("PROGRUZKA") & NewFilename & " " & Postavshik & " " & Date & " " & Number & ".csv") <> vbNullString
        Number = Number + 1
    Loop
    Workbooks(ProgruzkaFilename).SaveAs Filename:=Environ("PROGRUZKA") & NewFilename & " " & Postavshik & " " & Date & " " & Number & ".csv", FileFormat:=xlCSV
    
ended:
Application.ScreenUpdating = True
End Sub

Public Function Round_Up(ByVal input_number As Long, ByVal digits As Integer) As Long
    
    Dim result As Long
    
    result = VBA.Round(input_number)
    If digits = -1 Then
        Round_Up = result + (10 - result Mod 10)
    Else
        Round_Up = result
    End If
    
End Function

Public Function DublicateFind(UFKods_Array) As Boolean
    
    Dim ProverkaDict As Object
    Dim i As Long

    Set ProverkaDict = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(UFKods_Array)
    
        UFKods_Array(i) = Application.Trim(UFKods_Array(i))
        
        If Proschet_Dict.exists(UFKods_Array(i)) = False Then
            MsgBox "Отсутствует в просчете! " & UFKods_Array(i)
            DublicateFind = True
            Exit Function
        End If

        If ProverkaDict.exists(UFKods_Array(i)) = True Then
            MsgBox "Дубль! " & UFKods_Array(i)
            DublicateFind = True
            Exit Function
        Else
            ProverkaDict.Add UFKods_Array(i), i
        End If

        If UserForm1.OptionButton3 = True Then
            If InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "П") = 0 Then
                MsgBox UFKods_Array(i) & " Указанная категория не найдена!"
                DublicateFind = True
                Exit Function
            End If
        ElseIf UserForm1.OptionButton2 = True Then
            If InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat)) = 0 Then
                MsgBox UFKods_Array(i) & " Указанная категория не найдена!"
                DublicateFind = True
                Exit Function
            ElseIf InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "П") <> 0 Or _
            InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "ПП") <> 0 Then
                MsgBox UFKods_Array(i) & " Установлен как промо!"
                DublicateFind = True
                Exit Function
            End If
        End If

    Next i

End Function

Public Function FillInfocolumnRetailPrice(ByRef RetailPrice As Long, ByRef RetailPriceInfo As Long) As String

    If RetailPrice <> 0 And RetailPriceInfo <> 0 Then
        FillInfocolumnRetailPrice = RetailPrice & " (МРЦ), " & RetailPriceInfo & " (РРЦ)"
    ElseIf RetailPrice <> 0 Then
        FillInfocolumnRetailPrice = RetailPrice & " (МРЦ)"
    ElseIf RetailPriceInfo <> 0 Then
        FillInfocolumnRetailPrice = RetailPriceInfo & " (РРЦ)"
    End If

End Function

Public Function FillInfocolumnJU(ByRef Progruzka_Name, JU, ByRef Komplektuyeshee, _
ByRef Quantity, GorodOtpravki, ByRef ObreshetkaPrice) As String

If Progruzka_Name = "СОТ/Торуда" Then
    If JU = 0 Or Len(JU) = 0 Then
        FillInfocolumnJU = "нет"
    ElseIf Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Or Komplektuyeshee = "К2" Or Komplektuyeshee = "Т" Then
        FillInfocolumnJU = "нет"
    Else
        If Quantity > 1 Then
            If DictLgotniiGorod.exists(GorodOtpravki) = False Then
                FillInfocolumnJU = ObreshetkaPrice / 2 / Quantity & " +30% стоимости доставки DPD /" & Quantity
            Else
                FillInfocolumnJU = ObreshetkaPrice / 2 / Quantity & " +30% стоимости доставки DPD /" & Quantity
            End If
        ElseIf DictLgotniiGorod.exists(GorodOtpravki) = False Then
            FillInfocolumnJU = ObreshetkaPrice / 2 & " +30% стоимости доставки DPD"
        Else
            FillInfocolumnJU = ObreshetkaPrice / 2 & " +30% стоимости доставки DPD"
        End If
    End If
ElseIf Progruzka_Name = "Киска" Then
    If JU = 0 Or JU = vbNullString Then
        FillInfocolumnJU = "нет"
    ElseIf Komplektuyeshee = "К" Or Komplektuyeshee = "К1" Or Komplektuyeshee = "К2" Then
        FillInfocolumnJU = "нет"
    Else
        FillInfocolumnJU = "Для товаров данного магазина расчет обрешетки не предусмотрен!"
    End If
Else
    MsgBox "Неправильно указан аргумент Progruzka_Name функции FillInfocolumnJU"
    Stop
End If

End Function

Public Function FillInfocolumnMargin(Komplektuyeshee, Margin, Optional NeiskluchenieToruda = False, Optional Promo = False)

'Sot/Toruda nepromo
    
    If Margin = 0 Then
        FillInfocolumnMargin = vbNullString
        Exit Function
    End If
    
    If Promo = True Then
    
        If Komplektuyeshee = "Т" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: " & Margin & ", федеральная - " & Margin & " (ПРОМО цена!)"
        ElseIf Komplektuyeshee = "К2" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: 2, федеральная - 1.9" & " (ПРОМО цена!)"
        ElseIf NeiskluchenieToruda = True Then
            FillInfocolumnMargin = "1.35 Россия," & 1.35 + SNG_MARGIN & " в СНГ, федеральная - 1.25" & " (ПРОМО цена!)"
        ElseIf Komplektuyeshee = "К" Then
            FillInfocolumnMargin = "2 - Россия, федеральная - 1.9" & " (ПРОМО цена!)"
        ElseIf Komplektuyeshee = "К1" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: 1.7, федеральная - 1.6" & " (ПРОМО цена!)"
        Else
            FillInfocolumnMargin = Margin & " - Россия," & Margin + SNG_MARGIN & " в СНГ, федеральная - " & Margin - 0.1 & " (ПРОМО цена!)"
        End If
        
    ElseIf Promo = False Then
    
        If Komplektuyeshee = "Т" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: " & Margin & ", федеральная - " & Margin & " (НЕ ПРОМО цена)"
        ElseIf Komplektuyeshee = "К2" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: 2, федеральная - 1.9" & " (НЕ ПРОМО цена)"
        ElseIf NeiskluchenieToruda = True Then
            FillInfocolumnMargin = "1.5 Россия," & 1.5 + SNG_MARGIN & " в СНГ, федеральная - 1.4" & " (НЕ ПРОМО цена)"
        ElseIf Komplektuyeshee = "К" Then
            FillInfocolumnMargin = "2, федеральная - 1.9" & " (НЕ ПРОМО цена)"
        ElseIf Komplektuyeshee = "К1" Then
            FillInfocolumnMargin = "Наценка согласована с отделом закупа: 1.7, федеральная - 1.6" & " (НЕ ПРОМО цена)"
        Else
            FillInfocolumnMargin = Margin & " - Россия," & Margin + SNG_MARGIN & " в СНГ, федеральная - " & Margin - 0.1 & " (НЕ ПРОМО цена)"
        End If
        
    End If


End Function

'Расчет значения в столбце TypeZabor прогрузки
'TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, ProschetArray(NRowsArray(d), 12), NRowsArray(d), LgotniiGorod)
Public Function FillInfocolumnTypeZabor(Komponent, ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps, Optional spez As Integer = 0) As String
    
    
    
    'Если Кол-во штук > 1, тогда
    If ProschetArray(NRowProschet, 5) > 1 Then
        If LgotniiGorod = True Then
            If spez = 1 Then
                'Если город льготный, и ?spez = 1?
                'TypeZabor:
                'Тип забора в РФ: DPD
                'Тип забора в СНГ: Из просчета по столбцу "Тип забора"
                'Стоимость забора:
                'В России (Полный сервис) = 200 / Кол-во штук
                'В России (Эконом) = 0
                'В СНГ = Забор из просчета / Кол-во штук
                'Доп. издержки = Разгрузка из просчета / Кол-во штук
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем DPD(для отправок по России), тип забора в СНГ-  " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - в России(Полный сервис), 0 - в России(Эконом)) " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            ElseIf spez = 2 Then
                'Если город льготный, и ?spez = 2?
                'TypeZabor:
                'Тип забора в РФ: Pickpoint
                'Тип забора в СНГ: DPD
                'Стоимость забора:
                'В СНГ = Забор из просчета / Кол-во штук
                'Доп. издержки = Разгрузка из просчета / Кол-во штук
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем Pickpoint(для отправок по России), тип забора в СНГ-DPD, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            Else
                'Если город льготный, тогда
                'TypeZabor:
                'Тип забора в РФ: DPD
                'Тип забора в СНГ: Из просчета по столбцу "Тип забора"
                'Стоимость забора:
                'В России = 200 / Кол-во штук
                'В СНГ = Забор из просчета / Кол-во штук
                'Доп. издержки = Разгрузка из просчета / Кол-во штук
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем DPD(для отправок по России), тип забора в СНГ- " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - в России, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            End If
        ElseIf LgotniiGorod = False Then
            'Если город нельготный, тогда
            'TypeZabor:
            'Тип забора в РФ: Из просчета по столбцу "Тип забора"
            'Стоимость забора:
            'В России = 200 / Кол-во штук
            'В СНГ = Забор из просчета / Кол-во штук
            'Доп. издержки = Разгрузка из просчета / Кол-во штук
            'Надбавка за тяжелую посылку = tps
            FillInfocolumnTypeZabor = ProschetArray(NRowProschet, 8) & " " & CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - в России, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
            " - в СНГ, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
            tps & " - надбавка за тяжелую посылку"
        End If
    Else
        If LgotniiGorod = True Then
            If spez = 1 Then
                'Если город льготный и ?spez = 1?, тогда
                'TypeZabor:
                'Тип забора в РФ: DPD
                'Тип забора в СНГ: Из просчета по столбцу "Тип забора"
                'Стоимость забора:
                'В России (Полный сервис) = 200
                'В России (Эконом) = 0
                'В СНГ = Забор из просчета
                'Доп. издержки = Разгрузка из просчета
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем DPD(для отправок по России), тип забора в СНГ-  " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) & " - в России(Полный сервис), 0 - в России(Эконом)) " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            ElseIf spez = 2 Then
                'Если город льготный, и ?spez = 2?
                'TypeZabor:
                'Тип забора в РФ: Pickpoint
                'Тип забора в СНГ: DPD
                'Стоимость забора:
                'В СНГ = Забор из просчета
                'Доп. издержки = Разгрузка из просчета
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем Pickpoint(для отправок по России), тип забора в СНГ-DPD, " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            Else
                'Если город льготный, тогда
                'TypeZabor:
                'Тип забора в РФ: DPD
                'Тип забора в СНГ: Из просчета по столбцу "Тип забора"
                'Стоимость забора:
                'В России = 200
                'В СНГ = Забор из просчета
                'Доп. издержки = Разгрузка из просчета
                'Надбавка за тяжелую посылку = tps
                FillInfocolumnTypeZabor = "Забираем DPD(для отправок по России), тип забора в СНГ- " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) & " - в России, " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - в СНГ, " & CLng(RazgruzkaPrice) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
                tps & " - надбавка за тяжелую посылку"
            End If
        Else
            'Если город нельготный, тогда
            'TypeZabor:
            'Тип забора в РФ: Из просчета по столбцу "Тип забора"
            'Стоимость забора:
            'В России = 200
            'В СНГ = Забор из просчета
            'Доп. издержки = Разгрузка из просчета
            'Надбавка за тяжелую посылку = tps
            FillInfocolumnTypeZabor = ProschetArray(NRowProschet, 8) & " " & CLng(ZaborPrice) & " - в России, " & CLng(ProschetArray(NRowProschet, 13)) & _
            " - в СНГ, " & CLng(RazgruzkaPrice) & " - Доп. издержки (Разгрузка, особенности забора, доп. упаковка и др.), " & _
            tps & " - надбавка за тяжелую посылку"
        End If
    End If

End Function

Public Function DeclareDatatypeProschet(ByVal NColumn, Value)

'Поставщик
If NColumn = 1 Then
    DeclareDatatypeProschet = CStr(Value)
'Артикул
ElseIf NColumn = 2 Then
    DeclareDatatypeProschet = CStr(Value)
'Категория - не используется в просчете цены
ElseIf NColumn = 3 Then
    DeclareDatatypeProschet = vbNullString
'Количество мест - не используется в просчете цены
ElseIf NColumn = 4 Then
    DeclareDatatypeProschet = vbNullString
'Количество штук
ElseIf NColumn = 5 Then
    If Value <> 0 Or Len(Value) <> 0 Then
        If Value > 1 Then
            DeclareDatatypeProschet = CLng(Value)
        Else
            DeclareDatatypeProschet = CLng(1)
        End If
    Else
        DeclareDatatypeProschet = CLng(1)
    End If
'Город отправки
ElseIf NColumn = 6 Then
    DeclareDatatypeProschet = CStr(Value)
'Наценки
ElseIf NColumn = 7 Then
    DeclareDatatypeProschet = CStr(Value)
'Тип забора
ElseIf NColumn = 8 Then
    DeclareDatatypeProschet = CStr(Value)
'Объем
ElseIf NColumn = 9 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Вес
ElseIf NColumn = 10 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Объемный вес
ElseIf NColumn = 11 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Разгрузка
ElseIf NColumn = 12 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Забор
ElseIf NColumn = 13 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Жесткая упаковка
ElseIf NColumn = 14 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CDbl(0)
    End If
'Объем с ЖУ - не используется в просчете цены
ElseIf NColumn = 15 Then
    DeclareDatatypeProschet = vbNullString
'Формула (Комплектующее)
ElseIf NColumn = 16 Then
    DeclareDatatypeProschet = CStr(Value)
'Код 1С
ElseIf NColumn = 17 Then
    DeclareDatatypeProschet = CStr(Value)
'РРЦ
ElseIf NColumn = 18 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'Оптовая цена
ElseIf NColumn = 19 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'РРЦ (ИНФО)
ElseIf NColumn = 20 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
Else
    MsgBox "Неправильно передан аргумент NColumn функции DeclareDatatypeProschet)"
    Stop
End If

End Function

Public Function GetPriceStep(ByRef Sitename, ByRef NCol_Progruzka As Long, ByRef PriceValue As Long) As Long

    If Sitename = "Б2С СОТ/Торуда" Then
        If PriceValue > 7500 Then
            If NCol_Progruzka >= 120 And NCol_Progruzka < 126 Then
                If PriceValue > 150000 Then
                    GetPriceStep = 10
                Else
                    GetPriceStep = 1
                End If
            Else
                GetPriceStep = 10
            End If
        Else
            If NCol_Progruzka >= 108 And NCol_Progruzka < 120 Or NCol_Progruzka >= 148 And NCol_Progruzka <= 155 Then
                GetPriceStep = 10
            Else
                GetPriceStep = 1
            End If
        End If
        
    ElseIf Sitename = "Б2Б ПМК" Then
        If NCol_Progruzka < 92 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        ElseIf NCol_Progruzka >= 92 And NCol_Progruzka < 99 Then
            GetPriceStep = 10

        ElseIf NCol_Progruzka = 99 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        ElseIf NCol_Progruzka = 100 Then
            GetPriceStep = 1
        ElseIf NCol_Progruzka = 101 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        ElseIf NCol_Progruzka = 102 Then
            GetPriceStep = 10
        End If

    ElseIf Sitename = "Б2Б Торуда" Then
        If NCol_Progruzka < 128 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        Else
            GetPriceStep = 10
        End If

    ElseIf Sitename = "Б2Б Кабинетоф" Then
        If NCol_Progruzka < 96 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        ElseIf NCol_Progruzka >= 96 And NCol_Progruzka < 104 Then
            GetPriceStep = 10
        Else
            GetPriceStep = 1
        End If
    
    ElseIf Sitename = "Б2Б ПМК Ледосвет" Then
        If NCol_Progruzka < 108 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
    
        ElseIf NCol_Progruzka >= 108 And NCol_Progruzka < 115 Then
            GetPriceStep = 10
    
        ElseIf NCol_Progruzka > 115 And NCol_Progruzka < 126 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
    
        ElseIf NCol_Progruzka = 126 Then
            GetPriceStep = 1
    
        ElseIf NCol_Progruzka = 127 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        End If
        
    Else
        MsgBox "Неправильно указан аргумент Sitename функции GetPriceStep"
        Stop
    End If

End Function
