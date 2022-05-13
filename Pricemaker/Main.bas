Attribute VB_Name = "Main"
Option Explicit

Public AkciaArray As Variant                            'Массив листа "Акция" от A1:F до последней заполненной ячейки
Public DictAkcia As Object                                   'Ключ - Код 1С из листа "Акция", Элемент - номер строки
Public ProschetArray As Variant                      'Активный лист от А1:T До последней заполненной ячейки + 1
Public Proschet_Dict As Object                             'Ключ - Код 1С из листа "Просчет цен", элемент - номер строки
Public ProgruzkaMainArray As Variant                     'Активная книга. Активный лист. От A1:T до последней заполненной ячейки + 1
Public ProschetWorkbookName As String
Public kat As String                             'Текст из юзерформы "Наценки"

Public TarifOMAXTArray As Variant                     'Лист(omaxT) от С1:K до последней заполненной ячейки + 1
Public DictTarifOMAXT As Object                            'Ключ - Лист (omaxT), Столбец NColumnProgruzka, элемент - номер строки
Public TarifOMAXDArray As Variant                     'Лист(omaxД) от С1:K до последней заполненной ячейки + 1
Public DictTarifOMAXD As Object                           'Ключ - Лист(omaxД) ,Столбец NColumnProgruzka, элемент - номер строки (Диапазон NColumnProgruzka:K)
Public LgotniiGorod As Integer                          '0 - если город не льготный, 1 - если льготный, используется в FillInfocolumnTypeZabor
Public DictLgotniiGorod As Object                                'Ключ - Наименование города из листа "Льготные", Элемент - номер строки
Public SpecPostavshikArray() As Variant                                   'Спец.поставщики(A1:A) и последняя заполненная ячейка + 1
Public DictSpecPostavshik As Object                                     'Ключ - первый столбец листа "Спецпоставщики", элемент - номер строки
Public DictIsklucheniyaToruda As Object                                     'Ключ - Код 1С из листа "Исключения(toruda<=1000), Элемент - номер строки


Public wks As Object
Public KURS_TENGE As Double
Public KURS_BELRUB As Double
Public KURS_SOM As Double
Public KURS_DRAM As Double
Public SNG_MARGIN As Double                                     'Всегда равен значению из просчета(P7)

Sub Main()
    
    Dim i As Long
    
    ProschetWorkbookName = ActiveWorkbook.name
    
    If ProschetWorkbookName = "Просчет цен B2B.xlsb" Then
        Call Прогрузка_Б2Б
        Exit Sub
    End If

    Set DictSpecPostavshik = CreateObject("Scripting.Dictionary")

    'Добавляет в словарь DictSpecPostavshik: Ключ - наименование спец.поставщика, элемент - номер строки
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "Спец.-поставщики(1)" Then
            'SpecPostavshikArray = Спец.поставщики(A1:A) и последняя заполненная ячейка + 1
            SpecPostavshikArray = Workbooks(ProschetWorkbookName).Worksheets("Спец.-поставщики(1)").Range("A1:A" & _
            Workbooks(ProschetWorkbookName).Worksheets("Спец.-поставщики(1)").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
            For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("Спец.-поставщики(1)").Cells.SpecialCells(xlCellTypeLastCell).Row
                DictSpecPostavshik.Add CStr(SpecPostavshikArray(i, 1)), i
            Next i
        End If
    Next wks

    'ProschetArray = Активная книга. Активный лист. От A1:T до последней заполненной ячейки + 1
    ProschetArray = Workbooks(ProschetWorkbookName).ActiveSheet.Range("A1:T" & _
    Workbooks(ProschetWorkbookName).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

    'Set schetdict = CreateObject("Scripting.Dictionary")
    Set Proschet_Dict = CreateObject("Scripting.Dictionary")

    'Проверка кодов 1С на дубли в просчете
    'И добавление в словарь Proschet_Dict: Ключ - Код 1С из листа "Просчет цен", Элемент - № строки
    For i = 1 To UBound(ProschetArray, 1)
        If Trim(ProschetArray(i, 17)) <> vbNullString Then
            If Proschet_Dict.exists(Trim(ProschetArray(i, 17))) = True Then
                MsgBox "Дубль в просчете! " & ProschetArray(i, 17)
                Exit Sub
            End If
            Proschet_Dict.Add Trim(ProschetArray(i, 17)), i
        End If
    Next i

    'Проверка на дубли в листе omaxT по столбцу NColumnProgruzka
    'И добавление в словарь DictTarifOMAXT: Ключ - Город&Город из листа "Тариф(оmaxT)", Элемент - номер строки
    'TarifOMAXTArray = Лист(omaxT) от С1:K до последней заполненной ячейки + 1
    TarifOMAXTArray = Workbooks(ProschetWorkbookName).Worksheets("Тариф(omaxТ)").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("Тариф(omaxТ)").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

    Set DictTarifOMAXT = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(TarifOMAXTArray, 1)
        If Trim(TarifOMAXTArray(i, 1)) <> vbNullString Then
            If DictTarifOMAXT.exists(Trim(TarifOMAXTArray(i, 1))) = True Then
                MsgBox "Дубль в тарифах! " & TarifOMAXTArray(i, 1)
                Exit Sub
            End If
            DictTarifOMAXT.Add Trim(TarifOMAXTArray(i, 1)), i
        End If
    Next i

    'Проверка на дубли в листе omaxД по столбцу С
    'И добавление в словарь DictTarifOMAXD: Ключ - Город&Город из листа "Тариф(оmaxД)", Элемент - номер строки
    'TarifOMAXDArray = Лист(omaxД) от С1:K до последней заполненной ячейки + 1
    TarifOMAXDArray = Workbooks(ProschetWorkbookName).Worksheets("Тариф(omaxД)").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("Тариф(omaxД)").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    Set DictTarifOMAXD = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(TarifOMAXDArray, 1)
        If Trim(TarifOMAXDArray(i, 1)) <> vbNullString Then
            If DictTarifOMAXD.exists(Trim(TarifOMAXDArray(i, 1))) = True Then
                MsgBox "Дубль в тарифах(до двери)! " & TarifOMAXDArray(i, 1)
                Exit Sub
            End If
            DictTarifOMAXD.Add Trim(TarifOMAXDArray(i, 1)), i
        End If
    Next i

    'Проверяет каждый лист по названию
    'Если есть лист "Исключения(toruda<=1000)" - то переходит к start3
    'Если нет - создает лист "Исключения(toruda<=1000)" и переходит к start3
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "Исключения(toruda<=1000)" Then
            GoTo start3
        End If
    Next wks

    Call CreateSheet("Исключения(toruda<=1000)")

start3:
    'Добавляет в словарь DictIsklucheniyaToruda: Ключ - Код 1С из листа "Исключения(toruda<=1000), Элемент - номер строки
    Set DictIsklucheniyaToruda = CreateObject("Scripting.Dictionary")
    For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("Исключения(toruda<=1000)").Cells.SpecialCells(xlCellTypeLastCell).Row
        If Trim(Workbooks(ProschetWorkbookName).Worksheets("Исключения(toruda<=1000)").Cells(i, 1).Value) <> vbNullString Then
            DictIsklucheniyaToruda.Add Trim(Workbooks(ProschetWorkbookName).Worksheets("Исключения(toruda<=1000)").Cells(i, 1).Value), i
        End If
    Next i

    'Проверяет каждый лист по названию
    'Если есть лист "Льготные" - то переходит к start2
    'Если нет - создает лист "Льготные" и переходит к start2
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "Льготные" Then
            GoTo start2
        End If
    Next wks

    Call CreateSheet("Льготные")

    Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(1, 1).Value = "Москва"
    Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(2, 1).Value = "Ижевск"
    Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(3, 1).Value = "Ульяновск"
    Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(4, 1).Value = "Санкт-Петербург"
    Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(5, 1).Value = "Екатеринбург"

    'Добавляет в словарь DictLgotniiGorod: Ключ - Наименование города из листа "Льготные", Элемент - номер строки
start2:
    Set DictLgotniiGorod = CreateObject("Scripting.Dictionary")
    For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells.SpecialCells(xlCellTypeLastCell).Row
        If Trim(Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(i, 1).Value) <> vbNullString Then
            DictLgotniiGorod.Add Trim(Workbooks(ProschetWorkbookName).Worksheets("Льготные").Cells(i, 1).Value), i
        End If
    Next i

    'Проверяет каждый лист по названию
    'Если есть лист "Акция" - то переходит к start1
    'Если нет - создает лист "Акция" и переходит к start1
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "Акция" Then
            GoTo start1
        End If
    Next wks

    Call CreateSheet("Акция")
start1:
    'AkciaArray = Массив листа "Акция" от A1:F до последней заполненной ячейки
    AkciaArray = Workbooks(ProschetWorkbookName).Worksheets("Акция").Range("A1:F" & _
    Workbooks(ProschetWorkbookName).Worksheets("Акция").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    'Добавляет в словарь DictAkcia: Ключ - Код 1С из листа "Акция", Элемент - номер строки
    Set DictAkcia = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(AkciaArray, 1)
        If Trim(AkciaArray(i, 1)) <> vbNullString Then
            DictAkcia.Add Trim(AkciaArray(i, 1)), i
        End If
    Next i

    'Константы из просчета:
    If Workbooks(ProschetWorkbookName).name = "Просчет цен DDI.xlsb" Then
        SNG_MARGIN = 0
        KURS_TENGE = 1
        KURS_BELRUB = 1
        KURS_SOM = 1
        KURS_DRAM = 1
    Else
        SNG_MARGIN = ProschetArray(7, 16)
        KURS_TENGE = ProschetArray(1, 16)
        KURS_BELRUB = ProschetArray(2, 16)
        KURS_SOM = ProschetArray(3, 16)
        KURS_DRAM = ProschetArray(4, 16)
    End If

    UserForm1.Caption = Left(ThisWorkbook.name, 19)
    UserForm1.Show
    
    If UserForm1.OptionButton4 = True Then
        Call Salemaker
        ProschetArray = Workbooks(ProschetWorkbookName).ActiveSheet.Range("A1:T" & _
        Workbooks(ProschetWorkbookName).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
        Call Прогрузка_СОТ_Торуда

'        'Окно "Выгрузить категории"
'    ElseIf UserForm1.Выгрузить_категории = True Then
'        Call Выгрузить_категории

        'Окно "Замена наценок"
    ElseIf UserForm1.Замена_наценок = True Then
        Call Замена_наценок

        'Окно "Выгрузить промо"
    ElseIf UserForm1.Выгрузить_компоненты = True Then
        Call Выгрузить_компоненты
    
    ElseIf UserForm1.Добавить_компонент = True Or UserForm1.Удалить_компонент = True Then
        Call Добавить_удалить_компонент
    
        'Если просчет называется Просчет цен DDI.xlsb
    ElseIf Workbooks(ProschetWorkbookName).name = "Просчет цен DDI.xlsb" Then
        Call Прогрузка_DDI

        'Если выбран Default
    ElseIf UserForm1.OptionButton1 = True Or UserForm1.OptionButton3 = True Then
        Call Прогрузка_СОТ_Торуда

        'Если выбрано "Промо" и в окне "Кат-я" пусто - выводится окно "Некорректная категория"
        'и запускается Main
    ElseIf UserForm1.OptionButton2 = True And Trim(UserForm1.TextBox3) = vbNullString Then
        MsgBox "Некорректная категория"
        Call Main

        'Если выбрано "Промо"
    ElseIf UserForm1.OptionButton2 = True Then
        Call Подготовка_Промо_Общее

        'Если выбрано "Снять с промо" и в окно "Кат-я" пуста -выводится окно "Некорректная категория"
        'и запускается Main
    ElseIf UserForm1.OptionButton3 = True And Trim(UserForm1.TextBox3) = vbNullString Then
        MsgBox "Некорректная категория"
        Call Main

        'Если выбрано "Снять с промо"
'    ElseIf UserForm1.OptionButton3 = True Then
'        Call Прогрузка_Промо_Снять
    End If
    
End Sub

