Attribute VB_Name = "Module1"
Public raspreddict, mraspred As Variant, mproschetsotP As Variant, mproschetsotA As Variant, mproschetsotK As Variant, mproschetsotN As Variant, path As String, names As Variant, dict1, dict2

Private Function fNcolumn(source, namecol, Optional par As Integer = 1)

    For i = 1 To UBound(source, 2)

        For i1 = 1 To UBound(source, 1)

            If Trim(source(i1, i)) = namecol Then
                fNcolumn = i
                Exit Function
            End If

        Next i1

    Next i

    fNcolumn = "none"

    If par = 1 Then MsgBox "Столбец " & namecol & " не найден"

End Function

Private Function Create_File_Ostatki(abook, asheet)

    Workbooks.Add
    abook = ActiveWorkbook.name
    asheet = ActiveSheet.name
    Workbooks(abook).Worksheets(asheet).Cells(1, 1) = "kod 1s"
    Workbooks(abook).Worksheets(asheet).Cells(1, 2) = "ostatki"

End Function

Private Function Create_File_Ostatki_KK(abook1, asheet1)

    Workbooks.Add
    abook1 = ActiveWorkbook.name
    asheet1 = ActiveSheet.name
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 1) = "Category"
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 2) = "kod 1s"
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 3) = "ostatki"

End Function

Private Sub Concatenation()

    Workbooks.Add
    ConcatBookName = ActiveWorkbook.name
    Dim NamesArray()
    Dim NamesArrayKK()

    NOstatki = 0
    For Each wbk In Application.Workbooks
        If InStr(wbk.FullName, Date) <> 0 And InStr(wbk.FullName, ".csv") <> 0 And InStr(wbk.FullName, Environ("OSTATKI")) Then
            If InStr(StrConv(wbk.name, vbLowerCase), "компкресла") <> 0 Then
                
                NOstatkiKK = NOstatkiKK + 1
                ReDim Preserve NamesArrayKK(1 To NOstatkiKK)
                NamesArrayKK(NOstatkiKK) = wbk.name
            
                If NOstatkiKK = 1 Then
                
                    ConcatArrayKK = Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet. _
                    Range("A1:C" & Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
                
                ElseIf NOstatkiKK > 1 Then
                
                    If ConcatBookNameKK = sEmpty Then
                        
                        Workbooks.Add
                        ConcatBookNameKK = ActiveWorkbook.name
                        Workbooks(ConcatBookNameKK).ActiveSheet.Range("A1:C" & UBound(ConcatArrayKK, 1)).Value = ConcatArrayKK
                        ConcatArrayKK = Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet.Range("B2:C" & Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet. _
                        Cells.SpecialCells(xlCellTypeLastCell).Row).Value
                
                        Workbooks(ConcatBookNameKK).ActiveSheet.Range("B" & Workbooks(ConcatBookNameKK).ActiveSheet. _
                        Cells.SpecialCells(xlCellTypeLastCell).Row + 1 & ":C" & Workbooks(ConcatBookNameKK).ActiveSheet. _
                        Cells.SpecialCells(xlCellTypeLastCell).Row + UBound(ConcatArrayKK, 1)).Value = ConcatArrayKK
                    
                        GoTo nextKK
                        
                    End If
                
                    ConcatArrayKK = Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet.Range("B2:C" & Workbooks(NamesArrayKK(NOstatkiKK)).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row).Value
                
                    Workbooks(ConcatBookNameKK).ActiveSheet.Range("B" & Workbooks(ConcatBookNameKK).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row + 1 & ":C" & Workbooks(ConcatBookNameKK).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row + UBound(ConcatArrayKK, 1)).Value = ConcatArrayKK
nextKK:
                End If
        
            ElseIf InStr(StrConv(wbk.name, vbLowerCase), "остатки") <> 0 Then
                
                NOstatki = NOstatki + 1
                ReDim Preserve NamesArray(1 To NOstatki)
                NamesArray(NOstatki) = wbk.name
            
                If Workbooks(ConcatBookName).ActiveSheet.Cells(1, 1).Value = "" Then
                    
                    ConcatArray = Workbooks(NamesArray(NOstatki)).ActiveSheet. _
                    Range("A1:B" & Workbooks(NamesArray(NOstatki)).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
                    
                    Workbooks(ConcatBookName).ActiveSheet.Range("A1:B" & UBound(ConcatArray, 1)).Value = ConcatArray
                    
                Else
                    
                    ConcatArray = Workbooks(NamesArray(NOstatki)).ActiveSheet.Range("A2:B" & Workbooks(NamesArray(NOstatki)).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row).Value
                    
                    Workbooks(ConcatBookName).ActiveSheet.Range("A" & Workbooks(ConcatBookName).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row + 1 & ":B" & Workbooks(ConcatBookName).ActiveSheet. _
                    Cells.SpecialCells(xlCellTypeLastCell).Row + UBound(ConcatArray, 1)).Value = ConcatArray
                    
                End If
            End If
        End If
    
    Next

    If NOstatkiKK > 1 Then
        
        Do While Dir(Environ("OSTATKI") & NamesArrayKK(1) & "+concat компкресла " & Date & " " & nof0 & ".csv") <> ""
            nof0 = nof0 + 1
        Loop
        
        Workbooks(ConcatBookNameKK).Worksheets("Лист1").Cells(1, 1).Value = "category"
        Workbooks(ConcatBookNameKK).SaveAs Filename:=Environ("OSTATKI") & NamesArrayKK(1) & "+concat " & Date & " " & nof0 & ".csv", FileFormat:=xlCSV

    End If

    If NOstatki > 1 Then

        Do While Dir(Environ("OSTATKI") & NamesArray(1) & "+concat " & Date & " " & nof1 & ".csv") <> ""
            nof1 = nof1 + 1
        Loop
        
        Workbooks(ConcatBookName).SaveAs Filename:=Environ("OSTATKI") & NamesArray(1) & "+concat " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
    
    End If

End Sub

Public Sub Остатки()

    Set dict2 = CreateObject("Scripting.Dictionary")
    dict2.Add "Н103876", 1
    dict2.Add "Н103877", 1
    dict2.Add "Н103878", 1

    If Environ("UserName") = "argakov" Then
        UserForm1.MultiPage1.Value = 0
    ElseIf Environ("UserName") = "bazhenova" Then
        UserForm1.MultiPage1.Value = 1
    ElseIf Environ("UserName") = "gizzatulin" Then
        UserForm1.MultiPage1.Value = 2
    ElseIf Environ("UserName") = "churakov" Then
        UserForm1.MultiPage1.Value = 3
    End If
    
    If UserForm1.Autosave = True Then
        UserForm1.Concatenation.Enabled = True
    End If
    
    UserForm1.Caption = Left(ThisWorkbook.name, 23)
    UserForm1.Show

    If UserForm1.CheckBox1 = True Then
        Call ООвелон
    End If

    If UserForm1.CheckBox2 = True Then
        Call ОАвангард
    End If

    If UserForm1.CheckBox3 = True Then
        Call ОБелаяГвардия
    End If

    If UserForm1.CheckBox4 = True Then
        Call ОБелфорт
    End If

    'If UserForm1.CheckBox5 = True Then
    '    Call ОКиска
    'End If

    If UserForm1.CheckBox6 = True Then
        Call ОТэтчер
    End If

    If UserForm1.CheckBox7 = True Then
        Call ОRealChair
    End If

    If UserForm1.CheckBox9 = True Then
        Call ОWoodVille
    End If

    If UserForm1.CheckBox10 = True Then
        Call ОЦветМебели
    End If

    If UserForm1.CheckBox11 = True Then
        Call ОРэдБлэк
    End If

    If UserForm1.CheckBox13 = True Then
        Call ОБеон
    End If

    If UserForm1.CheckBox14 = True Then
        Call ОКМС
    End If

    If UserForm1.CheckBox15 = True Then
        Call ОДриада
    End If

    If UserForm1.CheckBox17 = True Then
        Call ОRomana
    End If

    If UserForm1.CheckBox18 = True Then
        Call ОKampfer
    End If

    If UserForm1.CheckBox19 = True Then
        Call ОНорденГрупп
    End If

    If UserForm1.CheckBox21 = True Then
        Call О100Печей
    End If

    If UserForm1.CheckBox22 = True Then
        Call ОEverproof1
    End If

    If UserForm1.CheckBox23 = True Then
        Call ОGardenWay
    End If

    If UserForm1.CheckBox24 = True Then
        Call Оormatek
    End If

    If UserForm1.CheckBox26 = True Then
        Call ОМебельдлядома
    End If

    If UserForm1.CheckBox27 = True Then
        Call ОЛибао
    End If

    If UserForm1.Тайпит = True Then
        Call ОТайпит
    End If

    If UserForm1.CheckBox30 = True Then
        Call ОDekomo
    End If

    If UserForm1.mebff = True Then
        Call ОМебфф
    End If

    If UserForm1.Logomebel = True Then
        Call ОЛогомебель
    End If

    If UserForm1.Русклимат = True Then
        Call ОРусклимат
    End If

    If UserForm1.Avanti = True Then
        Call ОAvanti
    End If

    If UserForm1.Sititek = True Then
        Call ОSititek
    End If

    If UserForm1.Aquaplast = True Then
        Call ОAquaplast
    End If

    If UserForm1.Amfitness = True Then
        Call ОAmfitness
    End If

    If UserForm1.ZSO = True Then
        Call ОZSO
    End If

    If UserForm1.rtoys = True Then
        Call Оrtoys
    End If

    If UserForm1.Family = True Then
        Call ОFamily
    End If

    If UserForm1.Atemi = True Then
        Call ОAtemi
    End If

    If UserForm1.Sparta = True Then
        Call ОSparta
    End If

    If UserForm1.sporthouse = True Then
        Call Оsporthouse
    End If

    If UserForm1.Afina = True Then
        Call ОAfina
    End If

    If UserForm1.Hudora = True Then
        Call ОHudora
    End If

    If UserForm1.ESF = True Then
        Call ОESF
    End If

    If UserForm1.Kettler = True Then
        Call Оkettler
    End If

    If UserForm1.FunDesk = True Then
        Call Оfundesk
    End If

    If UserForm1.Ruskresla = True Then
        Call ОRuskresla
    End If

    If UserForm1.mealux = True Then
        Call Оmealux
    End If

    If UserForm1.GoodChair = True Then
        Call ОGoodChair
    End If

    If UserForm1.Эколайн = True Then
        Call ОЭколайн
    End If

    If UserForm1.АР_Импекс = True Then
        Call ОАРИмпекс
    End If

    If UserForm1.Импекс = True Then
        Call ОИмпекс
    End If

    If UserForm1.ISit = True Then
        Call ОIsit
    End If

    If UserForm1.Officio = True Then
        Call ОOfficio
    End If

    If UserForm1.RHome = True Then
        Call ОRhome
    End If

    If UserForm1.Автовентури = True Then
        Call OАвтовентури
    End If

    If UserForm1.Hasstings = True Then
        Call ОHasttings
    End If

    If UserForm1.Riva = True Then
        Call ОRiva
    End If

    If UserForm1.Идея = True Then
        Call ОИдея
    End If

    If UserForm1.Триумф = True Then
        Call ОТриумф
    End If

    If UserForm1.Радуга = True Then
        Call ОРадуга
    End If

    If UserForm1.Экодизайн = True Then
        Call ОЭкодизайн
    End If

    If UserForm1.Фалто = True Then
        Call ОФалто
    End If

    If UserForm1.Unixline = True Then
        Call ОUnixline
    End If

    If UserForm1.Intex = True Then
        Call ОIntex
    End If

    If UserForm1.Промет = True Then
        Call ОПромет
    End If

    If UserForm1.AllBestGames = True Then
        Call ОAllBestGames
    End If

    If UserForm1.CharBroil = True Then
        Call ОCharBroil
    End If

    If UserForm1.Литком = True Then
        Call ОЛитком
    End If

    If UserForm1.Еврокамин = True Then
        Call ОЕврокамин
    End If

    If UserForm1.Swollen = True Then
        Call ОSwollen
    End If

    If UserForm1.Прагматика = True Then
        Call ОПрагматика
    End If

    If UserForm1.ТочкаОпоры = True Then
        Call ОТочкаОпоры
    End If

    If UserForm1.FoxGamer = True Then
        Call ОFoxGamer
    End If

    If UserForm1.Sheffilton = True Then
        Call ОSheffilton
    End If

    If UserForm1.МойДвор = True Then
        Call ОМойДвор
    End If

    If UserForm1.VictoryFit = True Then
        Call ОVictoryFit
    End If
    
    If UserForm1.Luxfire = True Then
        Call ОLuxfire
    End If
    
    If UserForm1.ВПК = True Then
        Call ОВПК
    End If
    
    If UserForm1.Ultragym = True Then
        Call ОUltraGym
    End If
    
    If UserForm1.Ramart = True Then
        Call ОRamart
    End If
    
    If UserForm1.Мебельторг = True Then
        Call ОМебельторг
    End If
    
    If UserForm1.Gardeck = True Then
        Call ОGardeck
    End If
    
    If UserForm1.Парнас = True Then
        Call ОПарнас
    End If
    
    If UserForm1.Городки77 = True Then
        Call ОГородки77
    End If

    If UserForm1.Concatenation = True Then
        Call Concatenation
    End If

    Unload UserForm1
End Sub

Private Sub ООвелон()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Овелон(Flexter)\Остатки\остатки связь.xlsx"
    msvyaz = Workbooks("остатки связь.xlsx").Worksheets("Лист1").Range("A1:B" & _
    Workbooks("остатки связь.xlsx").Worksheets("Лист1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.Workbooks.Open "https://dsk-formula.ru/files/uploads/price/DSK-Formula-price-OPT3.xls", ReadOnly
    
    mprice = ActiveWorkbook.Worksheets("Прайс-лист").Range("D2:G" & _
    ActiveWorkbook.Worksheets("Прайс-лист").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        If IsNumeric(mprice(i, 1)) = True Then
            mprice(i, 1) = CLng(mprice(i, 1))
        End If
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискОвелон(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=" шт.", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Овелон " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("остатки связь.xlsx").Close False
    Workbooks("DSK-Formula-price-OPT3.xls").Close False

End Sub

Private Sub ПоискОвелон(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
    
    Dim r As Long
    
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror1
            r = r + 1
        Loop
        
        mostatki(i) = mprice(r, 4)
    
        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If
        
        GoTo ended
        
    Else
        GoTo onerror1
    End If

onerror1:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub
'
'Private Sub ОКиска()
'
'    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка кодов Кисонька.xlsx"
'    Dim msvyaz As Variant, mprice As Variant, mostatki As Variant, _
'    abook1 As String, asheet1 As String, _
'    a As String, i As Long
'
'    msvyaz = Workbooks("Привязка кодов Кисонька.xlsx").Worksheets("Лист1").Range("A1:C" & _
'    Workbooks("Привязка кодов Кисонька.xlsx").Worksheets("Лист1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'    Workbooks.Add
'    With ActiveSheet.QueryTables.Add(Connection:= _
'    "TEXT;http://stripmag.ru/datafeed/insales_full_cp1251.csv", Destination:= _
'    Range("$A$1"))
'        .name = "insales_full_cp1251"
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .TextFilePromptOnRefresh = False
'        .TextFilePlatform = 1251
'        .TextFileStartRow = 1
'        .TextFileParseType = xlDelimited
'        .TextFileTextQualifier = xlTextQualifierDoubleQuote
'        .TextFileConsecutiveDelimiter = False
'        .TextFileTabDelimiter = False
'        .TextFileSemicolonDelimiter = True
'        .TextFileCommaDelimiter = False
'        .TextFileSpaceDelimiter = False
'        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'        1, 1, 1, 1, 1)
'        .TextFileTrailingMinusNumbers = True
'        .Refresh BackgroundQuery:=False
'    End With
'
'    a = ActiveWorkbook.name
'    mprice = ActiveWorkbook.ActiveSheet.Range("B2:L" & _
'    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'    Call Create_File_Ostatki_KK(abook1, asheet1)
'
'    ReDim mostatki(UBound(msvyaz, 1))
'
'    For i = 2 To UBound(msvyaz, 1)
'        Call ПоискКиска(i, msvyaz, mprice, abook1, asheet1, mostatki)
'    Next
'
'    If UserForm1.Autosave = True Then
'        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Киска " & Date & ".csv", FileFormat:=xlCSV
'    End If
'
'    Workbooks("Привязка кодов Кисонька.xlsx").Close False
'    Workbooks(a).Close False
'
'End Sub
'
'Private Sub ПоискКиска(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook1, ByRef asheet1, ByRef mostatki)
'    Dim r As Long
'    r = 1
'
'    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
'        Do While msvyaz(i, 2) <> mprice(r, 1)
'            On Error GoTo onerror2
'            r = r + 1
'        Loop
'        mostatki(i) = mprice(r, 11)
'
'        If mostatki(i) = 0 Or mostatki(i) = "" Then
'            mostatki(i) = 30000
'        End If
'        GoTo ended
'    Else
'        GoTo onerror2
'    End If
'
'onerror2:
'    mostatki(i) = 30000
'
'ended:
'    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
'    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = "ero" & msvyaz(i, 2)
'    Workbooks(abook1).Worksheets(asheet1).Cells(i, 1) = msvyaz(i, 3)
'End Sub

Private Sub ОАвангард()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Авангард\остатки\связь.xlsx"
    msvyaz = Workbooks("связь.xlsx").Worksheets("Лист1").Range("B1:D" & _
    Workbooks("связь.xlsx").Worksheets("Лист1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Авангард"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B4:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискАвангард(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Авангард " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискАвангард(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 5)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОБелфорт()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Белфорт\Остатки\ост. связь.xlsx"
    msvyaz = Workbooks("ост. связь.xlsx").Worksheets("Лист1").Range("A1:B" & _
    Workbooks("ост. связь.xlsx").Worksheets("Лист1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.DisplayAlerts = False
    G = "https://belfortkamin.ru/test/Stok_" & Format(Date, "dd.mm.yyyy") & ".xlsx"
    On Error Resume Next
    Application.Workbooks.Open G
    Application.DisplayAlerts = True

    If ActiveWorkbook.name <> G Then
        MsgBox "Выберите остатки Белфорт"
        a = Application.GetOpenFilename
        If a = False Then
            MsgBox "Файл не выбран"
            GoTo ended:
        End If
        Workbooks.Open Filename:=a
    End If

    a = ActiveWorkbook.name


    mprice = ActiveWorkbook.ActiveSheet.Range("C2:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискБелфорт(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Белфорт " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False

ended:
    Workbooks("ост. связь.xlsx").Close False
    Exit Sub
End Sub

Private Sub ПоискБелфорт(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = CLng(Replace(mprice(r, 6), ".00", ""))

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОТэтчер()

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Тетчер.xlsx"
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, book1 As String, _
    a, b As Long, mostatki As Variant, abook1 As String, asheet1 As String

    msvyaz = Workbooks("Привязка Тетчер.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Тетчер.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value


    MsgBox "Выберите остатки Тетчер"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A6:L" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискТэтчер(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Тэтчер " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Тэтчер компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Тетчер.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискТэтчер(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)
    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 12)


        If mostatki(i) = 0 Or mostatki(i) = "" Or InStr(mostatki(i), "Ожидаемая") <> 0 Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "Более 10" Then
            mostatki(i) = 10
        ElseIf mostatki(i) = "Ожидается" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "Спрашивайте у менеджера" Then
            mostatki(i) = 0
        ElseIf mostatki(i) = "Нет в наличии" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)


End Sub

Private Sub ОRealChair()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка RealChair.xlsx"
    msvyaz = Workbooks("Привязка RealChair.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка RealChair.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Real Chair"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B2:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRealChair(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки RealChair " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки RealChair компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка RealChair.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискRealChair(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)
    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 3)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОБелаяГвардия()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, _
    i As Long, a, b As String, mostatki As Variant, r As Integer, ii As Double, ili As Double

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Белая гвардия\Остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "Выберите остатки Белая Гвардия"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    b = ActiveSheet.name

    'Подсчет последней строки в файле для остатков
    LastRow = Workbooks(a).ActiveSheet.Cells.SpecialCells(xlLastCell).Row


    Call Create_File_Ostatki(abook, asheet)

    'Замена цветовых значений на числовые в 5 колонке
    For i = 1 To LastRow
        If Workbooks(a).Worksheets(b).Cells(i, 3).Interior.Color = 16777215 Then
            Workbooks(a).Worksheets(b).Cells(i, 4).Value = 10
        ElseIf Workbooks(a).Worksheets(b).Cells(i, 3).Interior.Color = 13434879 Then
            Workbooks(a).Worksheets(b).Cells(i, 4).Value = 3
        ElseIf Workbooks(a).Worksheets(b).Cells(i, 3).Interior.Color = 13421823 Then
            Workbooks(a).Worksheets(b).Cells(i, 4).Value = 1
        End If
    Next i

    mprice = Workbooks(a).Worksheets(b).Range("A1:E" & (LastRow + 1)).Value

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискБелаяГвардияИ(i, msvyaz, mostatki, r, ii, ili, mprice)
    Next
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискБелаяГвардияИли(i, msvyaz, abook, asheet, mostatki, r, ii, ili, mprice)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Белая Гвардия " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискБелаяГвардияИ(ByRef i, ByRef msvyaz, ByRef mostatki, ByRef r, ByRef ii, ByRef ili, ByRef mprice)

    ii = 0
    r = 1

    'msvyaz 2 = Для "UUID"
    'msvyaz 8 = Для "Код 1С"
    If msvyaz(i, 8) <> 0 And msvyaz(i, 8) <> "" Then
        Do While msvyaz(i, 8) <> mprice(r, 3)
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        If mprice(r, 4) = 10 Then
            ii = 10
        ElseIf mprice(r, 4) = 3 Then
            ii = 3
        ElseIf mprice(r, 4) = 1 Then
            ii = 1
        End If
    
        GoTo ended1

    Else
        GoTo ended1
    End If

onerror2:
    ii = 0.01

ended1:
    mostatki(i, 1) = ii

End Sub

Private Sub ПоискБелаяГвардияИли(ByRef i, ByRef msvyaz, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef r, ByRef ii, ByRef ili, ByRef mprice)

    ili = 0

    r = 1

    'msvyaz 3 = Для "UUID"
    'msvyaz 9 = Для "Код 1С"
    If msvyaz(i, 9) <> 0 And msvyaz(i, 9) <> "" Then
        Do While msvyaz(i, 9) <> mprice(r, 3)
            On Error GoTo res
            r = r + 1
        Loop
        If mprice(r, 4) = 10 And ili < 10 Then
            ili = 10
        ElseIf mprice(r, 4) = 3 And ili < 3 Then
            ili = 3
        ElseIf mprice(r, 4) = 1 And ili < 1 Then
            ili = 1
        End If
    
        'msvyaz 4 = Для "UUID"
        'msvyaz 10 = Для "Код 1С"
        If msvyaz(i, 10) <> 0 And msvyaz(i, 10) <> "" Then
            GoTo onerror1
        Else: GoTo ended
        End If
    Else
        GoTo ended
    End If

res:     Resume onerror1

onerror1:
    r = 1

    'msvyaz 4 = Для "UUID"
    'msvyaz 10 = Для "Код 1С"
    If msvyaz(i, 10) <> 0 And msvyaz(i, 10) <> "" Then
        Do While msvyaz(i, 10) <> mprice(r, 3)
            On Error GoTo res1
            r = r + 1
        Loop
        If mprice(r, 4) = 10 And ili < 10 Then
            ili = 10
        ElseIf mprice(r, 4) = 3 And ili < 3 Then
            ili = 3
        ElseIf mprice(r, 4) = 1 And ili < 1 Then
            ili = 1
        End If
    
        'msvyaz 5 = Для "UUID"
        'msvyaz 11 = Для "Код 1С"
        If msvyaz(i, 11) <> 0 And msvyaz(i, 11) <> "" Then
            GoTo onerror2
        Else: GoTo ended
        End If
    Else
        GoTo onerrorx
    End If

res1:     Resume onerror2

onerror2:
    r = 1

    'msvyaz 5 = Для "UUID"
    'msvyaz 11 = Для "Код 1С"
    If msvyaz(i, 11) <> 0 And msvyaz(i, 11) <> "" Then
        Do While msvyaz(i, 11) <> mprice(r, 3)
            On Error GoTo res2
            r = r + 1
        Loop
        If mprice(r, 4) = 10 And ili < 10 Then
            ili = 10
        ElseIf mprice(r, 4) = 3 And ili < 3 Then
            ili = 3
        ElseIf mprice(r, 4) = 1 And ili < 1 Then
            ili = 1
        End If
    
        'msvyaz 6 = Для "UUID"
        'msvyaz 12 = Для "Код 1С"
        If msvyaz(i, 12) <> 0 And msvyaz(i, 12) <> "" Then
            GoTo onerror3
        Else: GoTo ended
        End If
    Else
        GoTo onerrorx
    End If

res2:     Resume onerror3

onerror3:
    r = 1

    'msvyaz 6 = Для "UUID"
    'msvyaz 12 = Для "Код 1С"
    If msvyaz(i, 12) <> 0 And msvyaz(i, 12) <> "" Then
        Do While msvyaz(i, 12) <> mprice(r, 3)
            On Error GoTo onerrorx
            r = r + 1
        Loop
        If mprice(r, 4) = 10 And ili < 10 Then
            ili = 10
        ElseIf mprice(r, 4) = 3 And ili < 3 Then
            ili = 3
        ElseIf mprice(r, 4) = 1 And ili < 1 Then
            ili = 1
        End If
    Else
        GoTo onerrorx
    End If

onerrorx:
    If ili = 0 Then
        ili = 0.01
    End If

ended:
    mostatki(i, 2) = ili

    If mostatki(i, 1) = 0.01 Or mostatki(i, 2) = 0.01 Then
        mostatki(i, 2) = 30000
    ElseIf mostatki(i, 2) < mostatki(i, 1) And mostatki(i, 2) <> 0 Then
        mostatki(i, 2) = mostatki(i, 2)
    ElseIf mostatki(i, 2) > mostatki(i, 1) And mostatki(i, 1) <> 0 Then
        mostatki(i, 2) = mostatki(i, 1)
    ElseIf mostatki(i, 2) < mostatki(i, 1) Then
        mostatki(i, 2) = mostatki(i, 1)
    ElseIf mostatki(i, 2) > mostatki(i, 1) Then
        mostatki(i, 2) = mostatki(i, 2)
    End If

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i, 2)

End Sub

Private Sub ОWoodVille()
    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Woodville.xlsx"
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, _
    a, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки WoodVille"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.OpenText Filename:=a, DataType:=xlDelimited, Local:=True
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:AG" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    ReDim mostatki(UBound(msvyaz, 1))

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Остатки")
    If ncolumn1 = "none" Then GoTo ended

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискWoodVille(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки WoodVille " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки WoodVille Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Woodville.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискWoodVille(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОЦветМебели()
    Application.ScreenUpdating = False

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Цвет мебель\Таблица СОТб.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.Workbooks.Open "https://cvet-mebel.ru/files/uploads/stock/stock.xls"

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЦветМебели(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Цвет Мебели " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Цвет Мебели Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТб.xlsx").Close False
    Workbooks("stock.xls").Close False
    Application.ScreenUpdating = True

End Sub

Private Sub ПоискЦветМебели(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 4)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)


End Sub

Private Sub ОРэдБлэк()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Red-black(о)\Таблица  СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("D1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name
    ActiveWorkbook.XmlImport URL:= _
    "http://www.red-black.ru/pricelist_api.xml?org=4152ce29-8154-11e9-8be6-00155d396f01" _
    , ImportMap:=Nothing, Overwrite:=True, Destination:=Range("$A$1")

    mprice = ActiveWorkbook.ActiveSheet.Range("H1:O" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискРэдБлэк(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Red Black " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Red Black Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица  СОТ.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискРэдБлэк(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)
    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 8)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub


Private Sub ОБеон()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, _
    mostatki As Variant, mprice1 As Variant, mprice2 As Variant, mprice4 As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Беон\Таблица СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    MsgBox "Выберите остатки Беон"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    
    mprice = Workbooks(a).ActiveSheet.Range("A1:J" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискБеон(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="под заказ ", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="под заказ ", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Беон " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Беон Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТ.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискБеон(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long

    r = 1
    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 2)
            On Error GoTo onerror1
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 5)

        If mostatki(i) = 0 Or mostatki(i) = "" _
        Or Application.Trim(StrConv(mostatki(i), vbLowerCase)) = "под заказ" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror1
    End If

onerror1:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОКМС()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка КМС.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки КМС"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("F4:J" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        If IsNumeric(mprice(i, 1)) = True Then
            mprice(i, 1) = CStr(mprice(i, 1))
        End If
    Next

    For i = 1 To UBound(msvyaz, 1)
        If IsNumeric(msvyaz(i, 2)) = True Then
            msvyaz(i, 2) = CStr(msvyaz(i, 2))
        End If
    Next


    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискКМС(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки КМС " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False
ended:
    Workbooks("Привязка КМС.xlsx").Close False

End Sub

Private Sub ПоискКМС(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 5)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)


End Sub


Private Sub ОДриада()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Дриада\Остатк Дриада.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name

    ActiveWorkbook.XmlImport URL:= _
    "https://driada-sport.ru/images/blog/2/XML_prise_s.xml", ImportMap:=Nothing, _
    Overwrite:=True, Destination:=Range("$A$1")

    mprice = ActiveWorkbook.ActiveSheet.Range("E1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    For i = 1 To UBound(mprice, 1)
        If IsNumeric(mprice(i, 1)) = True Then
            mprice(i, 1) = CSng(mprice(i, 1))
        End If
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискДриада(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="более 10", Replacement:="21", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
      
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Дриада " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Дриада.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискДриада(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then

        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 5)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "более 10" Then
            mostatki(i) = 21
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОRomana()


    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Romana\Остатк Romana.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "Выберите остатки Romana"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    Workbooks.Add
    abook1 = ActiveWorkbook.name
    asheet1 = ActiveSheet.name

    Workbooks(abook1).Worksheets(asheet1).Cells(1, 1) = "sklad"
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 2) = "kod 1s"


    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRomana(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook1).Worksheets(asheet1).Sort.SortFields.Add Key _
    :=Range("A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortTextAsNumbers
    With Workbooks(abook1).Worksheets(asheet1).Sort
        .SetRange Range("A2:B" & _
        Workbooks(abook1).Worksheets(asheet1).Cells.SpecialCells(xlCellTypeLastCell).Row)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Workbooks(abook1).Worksheets(asheet1).Columns("A:B").Select
    ActiveSheet.Range("$A$1:$B$" & _
    Workbooks(abook1).Worksheets(asheet1).Cells.SpecialCells(xlCellTypeLastCell).Row).RemoveDuplicates Columns:=2, Header:= _
    xlYes

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Romana " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Склады Romana " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Romana.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискRomana(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)

    Dim r As Long, r1 As Integer

    r1 = 54
    r = 1


    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 4)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    r1 = 54
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    If r1 = 0 Then
        Exit Sub
    End If

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 1) = r1
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)


End Sub

Private Sub ОKampfer()

    Application.ScreenUpdating = False
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Set dict1 = CreateObject("Scripting.Dictionary")

    dict1.Add "Н101977", 1
    dict1.Add "Н101978", 1
    dict1.Add "Н101979", 1
    dict1.Add "Н101980", 1
    dict1.Add "Н101981", 1
    dict1.Add "Н101982", 1
    dict1.Add "Н101983", 1
    dict1.Add "Н101984", 1
    dict1.Add "Н101985", 1
    dict1.Add "Н101986", 1
    dict1.Add "К042181", 1
    dict1.Add "К042187", 1
    dict1.Add "К042253", 1
    dict1.Add "К042254", 1
    dict1.Add "К049962", 1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Kampfer\Остатк ILGC.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Open Filename:="https://centr-opt.com/pricelist/price_ilgc-group-dilerskaya-dropshipping-RRC.xls"
   
    mprice = ActiveWorkbook.ActiveSheet.Range("C1:N" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    Workbooks.Add
    abook1 = ActiveWorkbook.name
    asheet1 = ActiveSheet.name
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 1) = "sklad"
    Workbooks(abook1).Worksheets(asheet1).Cells(1, 2) = "kod 1s"

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискKampfer(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next


    Workbooks(abook).Worksheets(asheet).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=">20", Replacement:="21", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Kampfer " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Склады Kampfer " & Date & ".csv", FileFormat:=xlCSV

    End If

    Workbooks("Остатк ILGC.xlsx").Close False
    Workbooks("price_ilgc-group-dilerskaya-dropshipping-RRC.xls").Close False

    Application.ScreenUpdating = True
End Sub

Private Sub ПоискKampfer(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 12)

        If mostatki(i) = 0 And dict1.exists(msvyaz(i, 2)) = True Or mostatki(i) = "" And dict1.exists(msvyaz(i, 2)) = True Then
            mostatki(i) = 10000
        ElseIf mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    If dict1.exists(msvyaz(i, 2)) = True Then
        mostatki(i) = 10000
    Else
        mostatki(i) = 30000
    End If

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 1) = msvyaz(i, 3)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОНорденГрупп()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a _
    , abook1 As String, asheet1 As String, ncolumn


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Норден групп.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "Выберите остатки Норден Групп"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B1:L" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Основной склад")
    If ncolumn = "none" Then GoTo ended


    For i = 2 To UBound(msvyaz, 1)
        Call ПоискНорденГрупп(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn)
    Next

    Workbooks(abook).Worksheets(asheet).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="-", Replacement:=30000, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Worksheets(asheet1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="-", Replacement:=30000, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Норден Групп " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Норден Групп компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Норден групп.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub


Private Sub ПоискНорденГрупп(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki _
, ByRef abook1, ByRef asheet1, ByRef ncolumn)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, ncolumn)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub О100Печей()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\100 печей\Остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки 100 печей"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("F1:R" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call Поиск100Печей(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки 100Печей " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub


Private Sub Поиск100Печей(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 11) + mprice(r, 12)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОEverproof1()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Everproof\Таблица КК.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Everproof MSK"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("C1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended
    ncolumn1 = fNcolumn(mprice, "Остаток")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискEverproof1(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Everproof MSK " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Everproof MSK Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица КК.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискEverproof1(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)


End Sub

Private Sub ОGardenWay()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Garden way\остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Garden Way"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискGardenWay(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=">15", Replacement:="20", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="<15", Replacement:="15", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="<5", Replacement:="5", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Garden Way " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискGardenWay(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 8)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)


End Sub

Private Sub Оormatek()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Ormatek\Остатк Ormatek.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Ormatek"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискOrmatek(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Ormatek " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Ormatek.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискOrmatek(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 2)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОМебельдлядома()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Мебель для дома\Остатк Мебель для дома.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Мебель для дома"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("C1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискМебельдлядома(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Мебель для дома " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Мебель для дома Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Мебель для дома.xlsx").Close False
    Workbooks(a).Close False

ended:


End Sub

Private Sub ПоискМебельдлядома(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 4)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОЛибао()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Либао\Таблица СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Либао"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЛибао(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Либао " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Либао Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТ.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискЛибао(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, sr1, sr2

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        sr1 = msvyaz(i, 1)
        sr2 = mprice(r, 1)
        Do While Trim(sr1) <> Trim(sr2)
            On Error GoTo onerror2
            r = r + 1
            sr2 = mprice(r, 1)
        Loop
        mostatki(i) = mprice(r, 2)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОТайпит()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Тайпит\СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Тайпит"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискТайпит(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Тайпит " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Тайпит Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("СОТ.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискТайпит(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, _
ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        If mprice(r, 5) < 0 Then
            mostatki(i) = 0
        Else
            mostatki(i) = mprice(r, 5)
        End If

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОDekomo()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, dict

    Application.Workbooks.Open "\\AKABINET\Doc\СпортОтдыхТуризм\Для Николая\SOT резерв\актуальный резерв\Просчет цен SOT(Люстры 938).xlsb"
    msvyaz = Workbooks("Просчет цен SOT(Люстры 938).xlsb").Worksheets("Просчет цен").Range("Q10:S" & _
    Workbooks("Просчет цен SOT(Люстры 938).xlsb").Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

    Application.Workbooks.Open "http://www.dekomo.ru/get_content/v4/excel.php?ost=1&brand_type[]=30&token=94fbe3099b8d207301b891b9ea7dc667"
    Columns("B:B").Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Range("B1:DI" & ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).RemoveDuplicates Columns:=1, Header:= _
    xlYes

    mprice = ActiveWorkbook.ActiveSheet.Range("B2:D" & _
    Cells(Rows.Count, 2).End(xlUp).Row).Value
    
    
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To Cells(Rows.Count, 2).End(xlUp).Row - 1
        dict.Add mprice(i, 1), mprice(i, 3)
    Next i

    nstock = ActiveWorkbook.name
    ReDim mostatki(UBound(msvyaz, 1))

    Workbooks.Add
    abook = ActiveWorkbook.name
    asheet = ActiveSheet.name
    Workbooks(abook).Worksheets(asheet).Cells(1, 1) = "Category"
    Workbooks(abook).Worksheets(asheet).Cells(1, 2) = "kod 1s"
    Workbooks(abook).Worksheets(asheet).Cells(1, 3) = "ostatki"
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискDekomo(i, msvyaz, mprice, abook, asheet, mostatki, dict)
    Next

    Workbooks(abook).Worksheets(asheet).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="Под заказ", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Dekomo " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Просчет цен SOT(Люстры 938).xlsb").Close False
    Workbooks(nstock).Close False

End Sub

Private Sub ПоискDekomo(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef dict)
    Dim r As Long
    r = 1
    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 3) <> 0 Then
        mostatki(i) = dict(msvyaz(i, 1))
    Else
        GoTo ended
    End If


    If mostatki(i) <= 0 Or mostatki(i) = "" Then
        mostatki(i) = 30000
    End If

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 3) = mostatki(i)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = 938


End Sub

Private Sub ОМебфф()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Меб-ФФ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Меб-ФФ"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    'C1:F
    mprice = ActiveWorkbook.ActiveSheet.Range("B1:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    nstock = ActiveWorkbook.name

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    Workbooks.Add
    abook2 = ActiveWorkbook.name
    asheet2 = ActiveSheet.name
    Workbooks(abook2).Worksheets(asheet1).Cells(1, 1) = "sklad"
    Workbooks(abook2).Worksheets(asheet1).Cells(1, 2) = "kod 1s"

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискМебфф(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, abook2, asheet2)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=">", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
      

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки МебФФ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки МебФФ Компкресла " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook2).SaveAs Filename:=Environ("OSTATKI") & "Склады МебФФ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Меб-ФФ.xlsx").Close False
    Workbooks(nstock).Close False

ended:

End Sub

Private Sub ПоискМебфф(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef abook2, ByRef asheet2)

    Dim r As Long
    r = 1

    If Trim(msvyaz(i, 2)) <> 0 And Trim(msvyaz(i, 2)) <> "" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 7)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = IIf(Trim(msvyaz(i, 1)) <> "Н103354", mostatki(i), 21)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = IIf(Trim(msvyaz(i, 1)) <> "Н103354", mostatki(i), 21)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook2).Worksheets(asheet2).Cells(i, 1) = msvyaz(i, 3)
    Workbooks(abook2).Worksheets(asheet2).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОOLSS()

    Dim r, r1, i, mostatki As Variant, abook, asheet, razm1 As Integer, razm2 As Integer

    If Dir(path & "\остаткиOLSS.xls") = "" Then
        Workbooks.Add
        ActiveWorkbook.ActiveSheet.Cells(1, 1).Value = "kod 1s"
        ActiveWorkbook.SaveAs Filename:=path & "\остаткиOLSS.xls", FileFormat:=56
    Else
        Application.Workbooks.Open path & "\остаткиOLSS.xls"
    End If

    abook = ActiveWorkbook.name
    asheet = ActiveWorkbook.ActiveSheet.name

    mostatki = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    razm1 = ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    razm2 = razm1

    r = 1
    Do While mproschetsotK(r, 1) <> "OLSS"
        r = r + 1
    Loop
    r1 = r

    Do While mproschetsotK(r, 1) = "OLSS"
        r = r + 1
    Loop

    For i = r1 To r - 1
        Call ПоискOLSS(i, mostatki, razm2, abook, asheet)
    Next i

    If razm2 <> razm1 Then
        Workbooks.Add
        ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Kod 1S"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "Ostatok"
        For i = 1 To razm2 - razm1
            ActiveWorkbook.ActiveSheet.Cells(i + 1, 1).Value = Workbooks(abook).Worksheets(asheet).Cells(razm1 + i, 1).Value
            ActiveWorkbook.ActiveSheet.Cells(i + 1, 2).Value = 999
        Next i
        If UserForm1.Autosave = True Then
            ActiveWorkbook.SaveAs Filename:=Environ("OSTATKI") & "Остатки OLSS " & Date & ".csv", FileFormat:=xlCSV
        End If
    End If

    If razm2 <> razm1 Then
        Workbooks.Add
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "Kod 1S"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "Ostatok"
        For i = 1 To razm2 - razm1
            ActiveWorkbook.ActiveSheet.Cells(i + 1, 2).Value = Workbooks(abook).Worksheets(asheet).Cells(razm1 + i, 1).Value
            ActiveWorkbook.ActiveSheet.Cells(i + 1, 3).Value = 999
        Next i
        If UserForm1.Autosave = True Then
            ActiveWorkbook.SaveAs Filename:=Environ("OSTATKI") & "Остатки OLSS Кресла " & Date & ".csv", FileFormat:=xlCSV
        End If
    End If


    Workbooks(abook).Close False

End Sub

Private Sub ПоискOLSS(ByRef i, ByRef mostatki, ByRef razm2, ByRef abook, ByRef asheet)
    Dim r2
    r2 = 1

    If mproschetsotK(i, 17) <> 0 And mproschetsotK(i, 17) <> "" Then

        Do While mostatki(r2, 1) <> mproschetsotK(i, 17)
            On Error GoTo onerror1
            r2 = r2 + 1
        Loop
        GoTo ended:

    End If

    Exit Sub

onerror1:
    razm2 = razm2 + 1
    Workbooks(abook).Worksheets(asheet).Cells(razm2, 1).Value = mproschetsotK(i, 17)

ended:

End Sub

Private Sub ОЛогомебель()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, _
    mostatki As Variant, a, abook1 As String, asheet1 As String, b, mprice1 As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Логомебель\Таблица СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Логомебель СПБ"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("F1:P" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Логомебель Москва"

    b = Application.GetOpenFilename
    If b = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=b

    b = ActiveWorkbook.name
    
    mprice1 = ActiveWorkbook.ActiveSheet.Range("F1:P" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЛогомебель(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЛогомебель1(i, msvyaz, mprice1, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Логомебель " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Логомебель Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТ.xlsx").Close False
    Workbooks(a).Close False
    Workbooks(b).Close False

ended:

End Sub

Private Sub ПоискЛогомебель(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        If mprice(r, 11) < 0 Then
            mostatki(i) = 0
        Else
            mostatki(i) = mprice(r, 11)
        End If

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        If InStr(mostatki(i), ">") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), ">", ""))
            mostatki(i) = CLng(mostatki(i))
        ElseIf InStr(mostatki(i), "<") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), "<", ""))
            mostatki(i) = CLng(mostatki(i))
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 3)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 3)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ПоискЛогомебель1(ByRef i, ByRef msvyaz, ByRef mprice1, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice1(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        If mprice1(r, 11) < 0 Then
            mprice1(r, 11) = 0
        Else
            mostatki(i) = IIf(mostatki(i) = 30000, CLng(mprice1(r, 11)), CLng(mostatki(i) + mprice1(r, 11)))
        End If

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = IIf(mostatki(i) = 30000, "30000", CLng(mostatki(i)))

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 3)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 3)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОРусклимат()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, _
    mostatki As Variant, mostatkiili As Variant, a

    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Русклимат привязка.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Русклимат - Киржач"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("D1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        mprice(i, 1) = Application.Trim(mprice(i, 1))
    Next i

    For i = 1 To UBound(msvyaz, 1)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
        msvyaz(i, 3) = Application.Trim(msvyaz(i, 3))
    Next i

    ReDim mostatki(UBound(msvyaz, 1))
    ReDim mostatkiili(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискРусклимат(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Русклимат " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Русклимат привязка.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискРусклимат(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    Dim PortalRow As Long
    Dim OchagRow As Long
    r = 1
    PortalRow = 1
    OchagRow = 1

    'Каминокомплекты
    If msvyaz(i, 2) <> "-" And msvyaz(i, 3) <> "-" Then
        Do While msvyaz(i, 2) <> mprice(PortalRow, 1)
            On Error GoTo onerror
            PortalRow = PortalRow + 1
        Loop
    
        Do While msvyaz(i, 3) <> mprice(OchagRow, 1)
            On Error GoTo onerror
            OchagRow = OchagRow + 1
        Loop
    
        If mprice(PortalRow, 5) <> sEmpty Or mprice(PortalRow, 5) <> 0 _
        Or mprice(OchagRow, 5) <> sEmpty Or mprice(OchagRow, 5) <> 0 Then
            If mprice(PortalRow, 5) > mprice(OchagRow, 5) Then
                mostatki(i) = mprice(OchagRow, 5)
            Else
                mostatki(i) = mprice(PortalRow, 5)
            End If
        Else
            mostatki(i) = 30000
        End If
        GoTo ended

        'Порталы
    ElseIf msvyaz(i, 2) <> "-" Then
        Do While msvyaz(i, 2) <> mprice(PortalRow, 1)
            On Error GoTo onerror
            PortalRow = PortalRow + 1
        Loop
    
        If mprice(PortalRow, 5) <> sEmpty Or mprice(PortalRow, 5) <> 0 Then
            mostatki(i) = mprice(PortalRow, 5)
        Else
            mostatki(i) = 30000
        End If
        GoTo ended
        'Очаги
    ElseIf msvyaz(i, 3) <> "-" Then
        Do While msvyaz(i, 3) <> mprice(OchagRow, 1)
            On Error GoTo onerror
            OchagRow = OchagRow + 1
        Loop
    
        If mprice(OchagRow, 5) <> sEmpty Or mprice(OchagRow, 5) <> 0 Then
            mostatki(i) = mprice(OchagRow, 5)
        Else
            mostatki(i) = 30000
        End If
        GoTo ended
        'Пустая привязка
    Else
        mostatki(i) = 30000
    End If

onerror:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = IIf(mostatki(i) <= 0, 30000, mostatki(i))

End Sub

Private Sub ОAvanti()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Avanti\Таблица СОТ.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Avanti"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B2:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    nstock = ActiveWorkbook.name

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAvanti(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="В наличии", Replacement:="1", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
       
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="В наличии", Replacement:="1", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
       

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Avanti " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Avanti Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТ.xlsx").Close False
    Workbooks(nstock).Close False

ended:

End Sub

Private Sub ПоискAvanti(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long
    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 2)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОSititek()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Сититек\Остатк Сититек.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="TDSheet", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Excel.Workbook(Web.Contents(""https://www.sititek.ru/dealer/Price.xls""), null, true)," & Chr(13) & "" & Chr(10) & "    TDSheet1 = Источник{[Name=""TDSheet""]}[Data]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    TDSheet1"
    '    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TDSheet" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [TDSheet]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "TDSheet"
        .Refresh BackgroundQuery:=False
    End With
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискSititek(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="30+", Replacement:="31", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="30+", Replacement:="31", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Sititek " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Sititek компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Сититек.xlsx").Close False
    Workbooks(nstock).Close False
End Sub

Private Sub ПоискSititek(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 2)
            On Error GoTo onerror2
            r = r + 1
        Loop

        If InStr(mprice(r, 1), "-") Then
            a = Split(mprice(r, 1), "-")
            mostatki(i) = a(0)
        Else
            mostatki(i) = mprice(r, 1)
        End If


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОAquaplast()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Аквапласт\сот1.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Aquaplast"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAquaplast(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Aquaplast " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("сот1.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискAquaplast(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 5)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОUnixline()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\UNIX\Остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="catalog", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://unixfit.ru/bitrix/catalog_export/catalog.php""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Измененный тип""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""url""," & _
    " type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""vendorCode"", type text}, {""name"", type text}, {""Attribute:id"", Int64.Typ" & _
    "e}, {""Attribute:available"", type logical}, {""picture"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=catalog" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [catalog]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "catalog"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:J" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
 
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискUnixline(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Unixline " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(nstock).Close False
    Workbooks("связь.xlsx").Close False

End Sub

Private Sub ПоискUnixline(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 5))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 9)
        If mostatki(i) = False Or StrConv(mostatki(i), vbLowerCase) = "false" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = True Or StrConv(mostatki(i), vbLowerCase) = "true" Then
            mostatki(i) = 5
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub


Private Sub ОAmfitness()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\American fitness\Остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Amfitness"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    b = ActiveSheet.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:G" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ncolumn = fNcolumn(mprice, "Наименование")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Спб")
    If ncolumn1 = "none" Then GoTo ended

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAmfitness(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Amfitness " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискAmfitness(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b, ByRef ncolumn, _
ByRef ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop


        If Trim(StrConv(mprice(r, ncolumn1), vbUpperCase)) = "ЕСТЬ" Then
            mostatki(i) = 10
        ElseIf IsNumeric(mprice(r, ncolumn1)) = False Then
            mostatki(i) = 0
        Else
            mostatki(i) = mprice(r, ncolumn1)
        End If

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОZSO()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\ЗСО\остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="full", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Csv.Document(Web.Contents(""https://www.zavodsporta.ru/stockexport/index/full""),[Delimiter="";"",Encoding=1251])," & Chr(13) & "" & Chr(10) & "    #""Повышенные заголовки"" = Table.PromoteHeaders(Источник)," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(#""Повышенные заголовки"",{{""Артикул"", type text}, {""Категория"", type text}, {""Цена"", Int64.Type}, {""Цена-Р" & _
    "РЦ"", type text}, {""Наименование"", type text}, {""Остаток"", Int64.Type}, {""Изображения"", type text}, {""Описание"", type text}, {""Атрибуты"", type text}, {""Производитель"", type text}, {""URL"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип"""
    '    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=full" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [full]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "full"
        .Refresh BackgroundQuery:=False
    End With
    '    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискZSO(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки ZSO " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискZSO(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 6)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 10000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub Оrtoys()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\R-toys\Остатки\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="api_export_3", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://r-toys.ru/bitrix/catalog_export/api_export_3.xml""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Измененный тип""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""ur" & _
    "l"", type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""name"", type text}, {""description"", type text}, {""adult"", type text}," & _
    " {""weight"", type text}, {""Attribute:id"", Int64.Type}, {""Attribute:available"", type logical}, {""age"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент barcode"" = Table.ExpandTableColumn(#""Измененный тип2"", ""barcode"", {""Element:Text""}, {""barcode.Element:Text""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент barcode"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=api_export_3" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [api_export_3]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "api_export_3"
        .Refresh BackgroundQuery:=False
    End With
    
    mprice = ActiveWorkbook.ActiveSheet.Range("H1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRtoys(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Rtoys " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискRtoys(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 4)

        If mostatki(i) = True Then
            mostatki(i) = 10
        ElseIf mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub



Private Sub ОFamily()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Фэмили\Остатк Фэмили.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Family"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискFamily(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Family " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Фэмили.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискFamily(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 4)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub


Private Sub ОAtemi()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка АТЕМИ.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Atemi"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    On Error Resume Next
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    On Error GoTo 0
 
 
    mprice = ActiveWorkbook.ActiveSheet.Range("G1:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAtemi(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Atemi " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка АТЕМИ.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискAtemi(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 2))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 1)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОSparta()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\Спарта\остатки\связь Спарта.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Спарта"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискSparta(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Спарта " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь Спарта.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискSparta(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 5)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub Оsporthouse()

    Application.ScreenUpdating = False

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\sporthouse\связь.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="yml (2)", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://sportpro.ru/yml""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop"" = Table.ExpandTableColumn(#""Измененный тип"", ""shop"", {""offers""}, {""shop.offers""})," & Chr(13) & "" & Chr(10) & "    #""shop offers"" = #""Развернутый элемент shop""{0}[shop.of" & _
    "fers]," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент offer"" = Table.ExpandTableColumn(#""shop offers"", ""offer"", {""Attribute:available"", ""vendorCode""}, {""offer.Attribute:available"", ""offer.vendorCode""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент offer"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""yml (2)""" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [yml (2)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "yml__2"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call Поискsporthouse(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Sporthouse " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("связь.xlsx").Close False
    Workbooks(nstock).Close False

    Application.ScreenUpdating = True

End Sub

Private Sub Поискsporthouse(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b)

    Dim r As Long

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, 2))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 1)

        If mostatki(i) = 0 Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)


End Sub

Private Sub ОAfina()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Афина.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.OpenText Filename:="http://afinalux.ru/clients/import.csv", DataType:=xlDelimited, Local:=True

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("D1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAfina(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Afina " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Афина.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискAfina(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 8)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 0
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 0

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОHudora()
    Application.ScreenUpdating = False


    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Hudora\Остатк Hudora.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="Запрос1", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://hudoraworld.ru/wholesale.xml""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop"" = Table.ExpandTableColumn(#""Измененный тип"", ""shop"", {""offers""}, {""shop.offers""})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers"" = Table.E" & _
    "xpandTableColumn(#""Развернутый элемент shop"", ""shop.offers"", {""offer""}, {""shop.offers.offer""})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers.offer"" = Table.ExpandTableColumn(#""Развернутый элемент shop.offers"", ""shop.offers.offer"", {""vendorCode"", ""Attribute:available""}, {""shop.offers.offer.vendorCode"", ""shop.offers.offer.Attribute:available""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "  " & _
    "  #""Развернутый элемент shop.offers.offer"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Запрос1" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Запрос1]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "Запрос1"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
 

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискHudora(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Hudora " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Hudora.xlsx").Close False
    Workbooks(nstock).Close False
    Application.ScreenUpdating = True

End Sub

Private Sub ПоискHudora(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b)

    Dim r As Long

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = IIf(mprice(r, 2) = "true", 5, 30000)

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)


End Sub


Private Sub ОESF()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка ЕСФ.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки ESF"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Свободный остаток")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискESF(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки ESF " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки ESF компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка ЕСФ.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискESF(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "в наличии" Then
            mostatki(i) = 50
        ElseIf IsDate(mostatki(i)) = True Then
            mostatki(i) = 30000
        End If

        If InStr(StrConv(mostatki(i), vbLowerCase), "более") <> 0 Then
            mostatki(i) = Application.Trim(Replace(StrConv(mostatki(i), vbLowerCase), "более", vbNullString))
            mostatki(i) = CLng(mostatki(i))
        End If

        If InStr(mostatki(i), "шт") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), "шт", vbNullString))
        End If

        If InStr(mostatki(i), "нал") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), "нал", vbNullString))
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
End Sub
Private Sub Оkettler()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\kettler\остатки\связь kettler.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.DisplayAlerts = False
    On Error GoTo ended
    Application.Workbooks.Open "http://www.kettlermebel.ru/export"
    nstock = ActiveWorkbook.name
    Application.DisplayAlerts = True

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:AJ" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    ncolumn = fNcolumn(mprice, "/shop/offers/offer/vendorCode")
    ncolumn1 = fNcolumn(mprice, "/shop/offers/offer/quantum")

    For i = 2 To UBound(msvyaz, 1)
        Call Поискkettler(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b, ncolumn, ncolumn1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки kettler " & Date & ".csv", FileFormat:=xlCSV
    End If



ended:
Workbooks("связь kettler.xlsx").Close False
If nstock <> vbNullString Then
    Workbooks(nstock).Close False
End If
End Sub

Private Sub Поискkettler(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b, ByRef ncolumn, _
ByRef ncolumn1)


    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) = 0 Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub Оfundesk()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\FunDesk\связь\остатки Fundesk.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Fundesk"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Наличие")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискFundesk(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Fundesk " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Fundesk компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("остатки Fundesk.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискFundesk(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)


End Sub

Private Sub ОRuskresla()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Русские кресла\СОТ1.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Русские кресла"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    Workbooks.Add
    abook2 = ActiveWorkbook.name
    asheet2 = ActiveSheet.name
    Workbooks(abook2).Worksheets(asheet2).Cells(1, 1) = "sklad"
    Workbooks(abook2).Worksheets(asheet2).Cells(1, 2) = "kod 1s"


    ncolumn = fNcolumn(mprice, "Номенклатура.Код")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Свободный остаток", 2)
    If ncolumn1 = "none" Then
        ncolumn1 = fNcolumn(mprice, "Количество")
        If ncolumn1 = "none" Then GoTo ended
    End If

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRuskresla(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, abook2, asheet2)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Русские кресла " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Русские кресла компкресла " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook2).SaveAs Filename:=Environ("OSTATKI") & "Склад Русские кресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("СОТ1.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискRuskresla(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1 _
, ByRef abook2, ByRef asheet2)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub Оmealux()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Mealux\Остатк Mealux.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    abook2 = ActiveWorkbook.name

    MsgBox "Выберите остатки Mealux"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    For Each wks In Application.Workbooks(a).Worksheets
        wks.Activate
        On Error Resume Next
        Workbooks(a).ActiveSheet.ShowAllData
        On Error GoTo 0
        Range("A1:AE" & ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
        Selection.Copy
    
        Workbooks(abook2).Activate
    
        ActiveWorkbook.ActiveSheet.Cells(ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).Select
    
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
        Application.Workbooks(a).Activate
    
    Next wks
 
    mprice = Workbooks(abook2).ActiveSheet.Range("A1:AE" & _
    Workbooks(abook2).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Наличие")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call Поискmealux(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки mealux " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки mealux компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Остатк Mealux.xlsx").Close False
    Workbooks(a).Close False
    Workbooks(abook2).Close False

ended:

End Sub

Private Sub Поискmealux(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)



        If mostatki(i) = "Есть" Then
            mostatki(i) = 21
        ElseIf mostatki(i) = "Уточняйте" Then
            mostatki(i) = 0
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОGoodChair()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Хорошие кресла(o)\Связь.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Хорошие кресла"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = Workbooks(a).ActiveSheet.Range("B5:E" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Склад / Номенклатура")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Наличие")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискGoodchair(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="нет", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="нет", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Хорошие кресла " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Хорошие кресла компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Связь.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискGoodchair(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 0

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОЭколайн()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Эколайн\SOT.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Эколайн"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value


    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Свободно")
    If ncolumn = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЭколайн(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Эколайн " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Эколайн компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("SOT.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискЭколайн(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = StrConv(CStr(mprice(r, ncolumn)), vbLowerCase)
        
        If mostatki(i) = "достаточно" Then
            mostatki(i) = 20
        ElseIf mostatki(i) = "очень много" Then
            mostatki(i) = 20
        ElseIf mostatki(i) = "мало" Then
            mostatki(i) = 5
        ElseIf mostatki(i) = "нет" Then
            mostatki(i) = 30000
        Else
            mostatki(i) = 30000
        End If

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОАРИмпекс()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1 As String

    Dim naimen() As Variant
    Dim naimenconcat() As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\А.Р.Импекс\Связь.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки А.Р.Импекс"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 

    naimen() = Workbooks(a).ActiveSheet.Range("A1:B" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim naimenconcat(UBound(naimen), LBound(naimen))
    For i = 1 To UBound(naimen)
        j = 1
        naimenconcat(i, j - 1) = naimen(i, j) & " " & naimen(i, j + 1)
    Next i

    mprice = Workbooks(a).ActiveSheet.Range("A1:E" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    If fNcolumn(mprice, "Наименование товара") <> 1 And fNcolumn(mprice, "Модификация") <> 2 Then
        GoTo ended
    End If

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn1 = fNcolumn(mprice, "Свободно")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискАРИмпекс(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, naimenconcat, ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=" ", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки А.Р.Импекс " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки А.Р.Импекс компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Связь.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискАРИмпекс(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef naimenconcat, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(naimenconcat(r, 0))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОИмпекс()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, asheet2 As String

    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Импекс.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Импекс"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:H" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn1 = fNcolumn(mprice, "Доступно")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискИмпекс(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Импекс " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Импекс компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Импекс.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискИмпекс(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОIsit()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, _
    mostatki As Variant, mprice1 As Variant, mprice2 As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\i-sit\Таблица СОТ+КК остатки.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки I-sit"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice1 = Workbooks(a).Worksheets("COMFORT SEATING").Range("C3:F" & _
    ActiveWorkbook.Worksheets("COMFORT SEATING").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)


    For i = 2 To UBound(msvyaz, 1)
        Call ПоискIsit(i, msvyaz, mprice, abook, asheet, mostatki, mprice1, mprice2, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="нет в наличии", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="нет в наличии", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="в наличии", Replacement:="30001", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="в наличии", Replacement:="30001", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки I-sit " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки I-sit Компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица СОТ+КК остатки.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискIsit(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef mprice1, ByRef mprice2, _
ByRef abook1, ByRef asheet1)

    Dim r As Long
    
    r = 1
    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        
        Do While msvyaz(i, 1) <> mprice1(r, 1)
            On Error GoTo onerrorx
            r = r + 1
        Loop
        
        mostatki(i) = mprice1(r, 4)
        
        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If
        GoTo ended
        
    Else
        GoTo onerrorx
    End If
    
onerrorx:
    mostatki(i) = 30000
    
ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)

End Sub

Private Sub ОOfficio()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Officio\Привязка Officio.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Officio"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("D1:G" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)


    ncolumn = fNcolumn(mprice, "Наименование")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Доступно")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискOfficio(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="-", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Officio " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Officio компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Привязка Officio.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискOfficio(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОRhome()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка R-HOME.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки R-HOME"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:C" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Код")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "К продаже")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRHome(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="-", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки R-Home " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки R-Home компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Привязка R-HOME.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискRHome(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        If mprice(r, ncolumn1) > 0 Then
            mostatki(i) = mprice(r, ncolumn1)
        Else
            GoTo onerror2
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub OАвтовентури()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Автовентури.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.Workbooks.Open "http://www.autoventuri.ru/upload/prices/price.xlsx"
    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("Y13:AI" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice)
        If IsNumeric(mprice(i, 1)) = True Then
            mprice(i, 1) = CStr(mprice(i, 1))
        End If
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискАвтовентури(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Автовентури " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False
    Workbooks("Привязка Автовентури.xlsx").Close False
End Sub

Private Sub ПоискАвтовентури(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then

        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        mostatki(i) = mprice(r, 11)

        If mostatki(i) = 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = ">20" Then
            mostatki(i) = 21
        End If
        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОHasttings()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Hasting.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Hasttings"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:L" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискHasttings(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Hasttings " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("Привязка Hasting.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискHasttings(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then

        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        mostatki(i) = mprice(r, 12)
    
        If mostatki(i) = "Нет" Or mostatki(i) = "нет" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "Есть" Or mostatki(i) = "есть" Then
            mostatki(i) = 999
        End If
    
        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОRiva()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Кирилл\Кресла Riva\Привязка Riva.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Riva"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A3:R" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "Номенклатура.Код")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Москва Склад")
    If ncolumn1 = "none" Then
    End If

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRiva(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, abook2, asheet2)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Riva " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Riva компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Riva.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискRiva(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1 _
, ByRef abook2, ByRef asheet2)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)


        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 2)

End Sub

Private Sub ОИдея()
    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Идея.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Идея"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    Range("B1:E" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).MergeCells = False

    mprice = Workbooks(a).ActiveSheet.Range("B1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 2 To UBound(mprice)
        If mprice(i, 1) = sEmpty Then
            mprice(i, 1) = mprice(i - 1, 1)
            If mprice(i, 2) = sEmpty Then
                mprice(i, 2) = mprice(i - 1, 2)
            End If
        End If
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    ncolumn = fNcolumn(mprice, "Комплект")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Плаcтик")
    If ncolumn1 = "none" Then GoTo ended

    ncolumn2 = fNcolumn(mprice, "Подушки")
    If ncolumn2 = "none" Then GoTo ended

    ncolumn3 = fNcolumn(mprice, "Остаток")
    If ncolumn3 = "none" Then GoTo ended

    ReDim mconcat(UBound(mprice))
    For i = 1 To UBound(mconcat)
        If mprice(i, ncolumn1) <> sEmpty Then
            mconcat(i) = Trim(mprice(i, ncolumn)) & Trim(mprice(i, ncolumn1)) & Trim(mprice(i, ncolumn2))
            mconcat(i) = Trim(mconcat(i))
        Else
            mconcat(i) = mprice(i, ncolumn) & mprice(i, ncolumn2)
            mconcat(i) = Trim(mconcat(i))
        End If
    Next i

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискИдея(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, ncolumn2, ncolumn3, mconcat)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Идея " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Идея.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub ПоискИдея(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1 _
, ByRef ncolumn2, ByRef ncolumn3, ByRef mconcat)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mconcat(r)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn3)
    
        If Trim(StrConv(mostatki(i), vbLowerCase)) = "нет в наличии" Then
            mostatki(i) = 30000
        ElseIf Trim(StrConv(mostatki(i), vbLowerCase)) = "в наличии" Then
            mostatki(i) = 30001
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОТриумф()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Алексей Аргаков\Остатки\Привязка\Привязка Триумф.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="TDSheet (2)", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Excel.Workbook(Web.Contents(""ftp://178.248.71.47/OSTATKI_SPORT.xls""), null, true)," & Chr(13) & "" & Chr(10) & "    TDSheet1 = Источник{[Name=""TDSheet""]}[Data]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    TDSheet1"
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""TDSheet (2)""" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [TDSheet (2)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "TDSheet__2"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    For i = 1 To UBound(mprice, 2)
        
        For j = 1 To UBound(mprice, 1)
            mprice(j, i) = Application.Trim(mprice(j, i))
        Next j
        
        If mprice(1, i) = "Артикул" Or mprice(2, i) = "Артикул" Or mprice(3, i) = "Артикул" _
        Or mprice(1, i) = "Артикул наш" Or mprice(2, i) = "Артикул наш" Or mprice(3, i) = "Артикул наш" Then
            ncolumn = i
        End If
    
        If mprice(2, i) & " " & mprice(3, i) = "Высоковск Свободный остаток" Then
            ncolumn1 = i
        End If
        
        If mprice(2, i) & " " & mprice(3, i) = "Офис Свободный остаток" Then
            ncolumn2 = i
        End If
    
    Next i
    
    If ncolumn = 0 Then
        MsgBox "Столбец Артикул не найден"
        GoTo ended
    End If
    
    If ncolumn1 = 0 Then
        MsgBox "Столбец Высоковск свободный остаток не найден"
        GoTo ended
    End If
    
    If ncolumn2 = 0 Then
        MsgBox "Столбец Офис свободный остаток не найден"
        GoTo ended
    End If
    
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискТриумф(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1, ncolumn2)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Триумф " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Триумф.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub ПоискТриумф(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1, ByRef ncolumn2)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = CLng(mprice(r, ncolumn1) + mprice(r, ncolumn2))

        If mostatki(i) > 50 Then
            mostatki(i) = 999
        ElseIf mostatki(i) <= 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub
Private Sub ОРадуга()

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Радуга.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Радуга"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    'Снимает общий доступ к книге, если он присутствует
    If ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.ExclusiveAccess
        Application.DisplayAlerts = True
    End If

    Range("C1:F" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).MergeCells = False

    mprice = Workbooks(a).ActiveSheet.Range("C1:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(1, i) = Application.Trim(mprice(1, i))
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    ncolumn = fNcolumn(mprice, "Описание")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Цвет")
    If ncolumn1 = "none" Then GoTo ended

    ncolumn2 = fNcolumn(mprice, "Остатки СПб")
    If ncolumn2 = "none" Then GoTo ended

    ncolumn3 = fNcolumn(mprice, "Остатки Екб")
    If ncolumn3 = "none" Then GoTo ended

    For i = 2 To UBound(mprice)
        If mprice(i, ncolumn) = sEmpty Then
            mprice(i, ncolumn) = mprice(i - 1, ncolumn)
        End If
    Next i


    ReDim mconcat(UBound(mprice))

    For i = 1 To UBound(mconcat)
        mconcat(i) = Trim(mprice(i, ncolumn)) & " " & Trim(mprice(i, ncolumn1))
    
        If InStr(mconcat(i), Chr(160)) <> 0 Then
            mconcat(i) = Replace(mconcat(i), Chr(160), " ")
        End If
    
        mconcat(i) = Application.WorksheetFunction.Trim(mconcat(i))
    Next i

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискРадуга(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, ncolumn2, ncolumn3, mconcat)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Радуга " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Радуга.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub ПоискРадуга(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1 _
, ByRef ncolumn2, ByRef ncolumn3, ByRef mconcat)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mconcat(r)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = (mprice(r, ncolumn2) + mprice(r, ncolumn3))
    
        If mostatki(i) = 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОЭкодизайн()

    Dim r As Long

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Экодизайн.xlsx"

    stock_name = ActiveWorkbook.name

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A2:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Экодизайн"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("B1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(1, i) = Application.Trim(mprice(1, i))
    Next i

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ReDim mostatki(1 To UBound(msvyaz, 1))

    For i = 1 To UBound(msvyaz, 1)
        r = 1
        msvyaz(i, 3) = Application.Trim(msvyaz(i, 3))
        If msvyaz(i, 3) <> 0 And msvyaz(i, 3) <> "" And msvyaz(i, 3) <> "-" Then
            '3, 5, 7 - #msvyaz_komplekts
            '9, 10, 11 - #kolvo_shtuk
            If InStr(msvyaz(i, 1), "Комплект") <> 0 Then
            
                msvyaz(i, 1) = Replace(msvyaz(i, 1), "Комплект", "")
            
                'Рассчитывает количество частей для комплекта
                For j = 6 To 8
                    If msvyaz(i, j) <> 0 Then
                        parts_count = parts_count + 1
                    End If
                Next j
            
                ReDim parts(parts_count)
            
                For j = 0 To UBound(parts)
                    r = 1
                    Do While msvyaz(i, j + 3) <> mprice(r, 1) And r < UBound(mprice)
                        r = r + 1
                    Loop
                
                    parts(j) = (mprice(r, 6) + mprice(r, 10))
                
                    'Считает количество комплектов
                    If parts(j) = 0 Or parts(j) = sEmpty Then
                        max_part = 0
                        Exit For
                    ElseIf parts(j) <> 0 Then
                        parts(j) = parts(j) \ msvyaz(i, j + 6)
                        If parts(j) = 0 Or parts(j) = sEmpty Then
                            max_part = 0
                            Exit For
                        ElseIf parts(j) > max_part Then
                            max_part = parts(j)
                        End If
                    End If
            
                Next j
            
                If max_part = 0 Then
                    mostatki(i) = 30000
                Else
                    mostatki(i) = max_part
                End If
        
            Else

                Do While msvyaz(i, 3) <> mprice(r, 1)
                    r = r + 1
                    If r = UBound(mprice) Then
                        Exit Do
                    End If
                Loop
            
                'mprice(r, 6) - свободный остаток склад Красноярск
                'mprice(r, 10) - свободный остаток склад МСК
                mostatki(i) = (mprice(r, 6) + mprice(r, 10))
            
                If mostatki(i) = 0 Or mostatki(i) = sEmpty Then
                    mostatki(i) = 30000
                End If
            End If
        Else
            mostatki(i) = 30000
        End If
    Next

    Workbooks(abook).Worksheets(asheet).Range("A2").Resize(UBound(mostatki), 1).Value = msvyaz
    Workbooks(abook1).Worksheets(asheet1).Range("B2").Resize(UBound(mostatki), 1).Value = msvyaz

    Workbooks(abook).Worksheets(asheet).Range("B2").Resize(UBound(mostatki), 1).Value = WorksheetFunction.Transpose(mostatki)
    Workbooks(abook1).Worksheets(asheet1).Range("C2").Resize(UBound(mostatki), 1).Value = WorksheetFunction.Transpose(mostatki)

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Экодизайн " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Экодизайн компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка экодизайн.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ОФалто()

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Фалто\Остатки Фалто.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    MsgBox "Выберите остатки Фалто"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("B1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(1, i) = Application.Trim(mprice(1, i))
    Next i

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ReDim mostatki(UBound(msvyaz, 1))

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискФалто(i, msvyaz, mprice, abook, asheet, abook1, asheet1, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Фалто " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Фалто компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатки Фалто.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub
Private Sub ПоискФалто(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, _
ByRef abook1, ByRef asheet1, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = (mprice(r, 3))
    
        If mostatki(i) = 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОIntex()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Intex\Остатк Intex.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz)
        If IsNumeric(msvyaz(i, 1)) = True Then
            msvyaz(i, 1) = CStr(msvyaz(i, 1))
        End If
    Next i

    MsgBox "Выберите остатки Intex"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("D1:N" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискIntex(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Intex " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Остатк Intex.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискIntex(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ncolumn, ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While Trim(msvyaz(i, 1)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        mostatki(i) = Application.Trim((mprice(r, 11)))
        mostatki(i) = Replace(mostatki(i), " ", vbNullString)
    
        If mostatki(i) = "<5" Then
            mostatki(i) = 3
        ElseIf mostatki(i) = ">10" Then
            mostatki(i) = 10
        ElseIf mostatki(i) = ">100" Then
            mostatki(i) = 999
        ElseIf mostatki(i) = ">20" Then
            mostatki(i) = 999
        ElseIf mostatki(i) = ">200" Then
            mostatki(i) = 999
        ElseIf mostatki(i) = ">5" Then
            mostatki(i) = 30001
        ElseIf mostatki(i) = ">50" Then
            mostatki(i) = 999
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 2)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОПромет()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Промет.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Промет"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B2:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)


    For i = 2 To UBound(msvyaz, 1)
        Call ПоискПромет(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Промет " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Промет.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискПромет(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = CLng((mprice(r, 3)))
    
        If mostatki(i) < 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОAllBestGames()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка AllBestGames.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки AllBestGames"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    For Each Sh In Workbooks(a).Sheets
        If Application.Trim(StrConv(Sh.name, vbLowerCase)) = "товары для детей" Then
            SheetName = Sh.name
        End If
    Next Sh

    If SheetName = sEmpty Then
        MsgBox "Лист 'Товары для детей' в прайсе не найден"
        GoTo ended:
    End If

    mprice = Workbooks(a).Sheets(SheetName).Range("A3:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        mprice(i, 1) = Application.Trim(mprice(i, 1))
        mprice(i, 3) = Application.Trim(mprice(i, 3))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискAllBestGames(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки AllBestGames " & Date & ".csv", FileFormat:=xlCSV
    End If


    Workbooks("Привязка AllBestGames.xlsx").Close False
    Workbooks(a).Close False
ended:
End Sub

Private Sub ПоискAllBestGames(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 3)
    
        If mostatki(i) = "в наличии" Then
            mostatki(i) = 30001
        ElseIf mostatki(i) = "нет в наличии" Then
            mostatki(i) = 30000
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОCharBroil()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка CharBroil.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="yandex_896364", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://dealer.clmbs.ru/bitrix/catalog_export/yandex_896364.php?1623166096""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Измененный тип""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company""," & _
    " type text}, {""url"", type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    offer"
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=yandex_896364" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [yandex_896364]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "yandex_896364"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ncolumn = fNcolumn(mprice, "Attribute:id")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "quanity")
    If ncolumn1 = "none" Then GoTo ended

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискCharBroil(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки CharBroil " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Charbroil.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub ПоискCharBroil(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)
    
        If mostatki(i) <= 0 Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОЛитком()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Литком.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Литком"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B1:G" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        mprice(i, 1) = Application.Trim(mprice(i, 1))
        mprice(i, 2) = Application.Trim(mprice(i, 2))
    Next

    For i = 1 To UBound(msvyaz, 1)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЛитком(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Литком " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Литком.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискЛитком(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 2)
    
        If mostatki(i) <= 0 Or mostatki(i) = vbNullString Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОЕврокамин()

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Еврокамин.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Еврокамин"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    fileostatki_name = ActiveWorkbook.name
    FirstSheet = ActiveWorkbook.Sheets(1).name

    For Each Sh In ActiveWorkbook.Sheets
    
        ActiveSheetName = Sh.name
    
        If ActiveSheetName = FirstSheet Then
            ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName). _
            Range("A5:C" & ActiveWorkbook.Sheets(ActiveSheetName). _
            Cells.SpecialCells(xlCellTypeLastCell).Row).Value
        
            For i = 1 To UBound(ConcatArray, 1)
                If Application.Trim(ConcatArray(i, 1)) = "Москва" Then
                    ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName).Range("A5:C" & (i + 3)).Value
                    Exit For
                End If
            Next
          
            ReDim mprice(1, 1 To UBound(ConcatArray, 1))
            For j = 1 To UBound(mprice, 2)
                mprice(0, j) = Application.Trim(ConcatArray(j, 2))
                mprice(1, j) = Application.Trim(ConcatArray(j, 3))
            Next
        
            GoTo continue
        End If
    
        ActiveWorkbook.Sheets(ActiveSheetName). _
        Range("A5:C" & ActiveWorkbook.Sheets(ActiveSheetName). _
        Cells.SpecialCells(xlCellTypeLastCell).Row).MergeCells = False
    
        ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName). _
        Range("A5:C" & ActiveWorkbook.Sheets(ActiveSheetName). _
        Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
        For i = 1 To UBound(ConcatArray, 1)
            If Application.Trim(ConcatArray(i, 1)) = "Москва" Then
                ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName).Range("A5:C" & (i + 3)).Value
                Exit For
            End If
        Next
    
        StartPos = UBound(mprice, 2)
        ReDim Preserve mprice(1, 1 To (UBound(ConcatArray, 1) + UBound(mprice, 2)))
        i = 1
        For j = StartPos To UBound(mprice, 2)
            If i > UBound(ConcatArray, 1) Then Exit For
            If Application.Trim(ConcatArray(i, 1)) <> sEmpty Or Application.Trim(ConcatArray(i, 1)) <> vbNullString Then
                mprice(0, j) = Application.Trim(ConcatArray(i, 2))
                mprice(1, j) = Application.Trim(ConcatArray(i, 3))
            End If
            i = i + 1
        Next j
    
continue:
    Next Sh

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискЕврокамин(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Еврокамин " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Еврокамин.xlsx").Close False
    Workbooks(fileostatki_name).Close False


End Sub

Private Sub ПоискЕврокамин(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(0, r)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(1, r)
       
        If Trim(StrConv(mostatki(i), vbLowerCase)) = "нет в наличии" Then
            mostatki(i) = 30000
        ElseIf Trim(StrConv(mostatki(i), vbLowerCase)) = "в наличии" Then
            mostatki(i) = 30001
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОSwollen()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Swollen.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Swollen"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        mprice(i, 1) = Application.Trim(mprice(i, 1))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискSwollen(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки SWOLLEN " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Swollen.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub ПоискSwollen(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 6)
    
        If mostatki(i) >= 20 Then
            mostatki(i) = 999
        ElseIf mostatki(i) <= 0 Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "Более 10" Then
            mostatki(i) = 30001
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)


End Sub

Private Sub ОПрагматика()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Прагматик.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "Выберите остатки Прагматик"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 3) = Application.Trim(mprice(i, 3))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискПрагматика(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Прагматик " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Прагматик компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Прагматик.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub ПоискПрагматика(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 3)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 4)
    
        If mostatki(i) < 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОТочкаОпоры()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Точка опоры.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "Выберите остатки Точка Опоры"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B4:G" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    'Наименование - C(2)
    'Остатки - E(4)

    For i = 1 To UBound(mprice, 2)
        mprice(i, 2) = Application.Trim(mprice(i, 2))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискТочкаОпоры(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Точка Опоры " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка Точка опоры.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискТочкаОпоры(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 2)
    
        If mostatki(i) < 0 Or mostatki(i) = sEmpty Then
            mostatki(i) = 30000
        ElseIf mostatki(i) > 20 Then
            mostatki(i) = 999
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОFoxGamer()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка FoxGamer.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "Выберите остатки FoxGamer"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ncolumn = fNcolumn(mprice, "Артикул")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Наличие")
    If ncolumn1 = "none" Then GoTo ended

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискFoxGamer(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки FoxGamer " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки FoxGamer компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("Привязка FoxGamer.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub ПоискFoxGamer(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)
    
        If StrConv(CStr(mostatki(i)), vbLowerCase) = "да" Then
            mostatki(i) = 5
        ElseIf StrConv(CStr(mostatki(i)), vbLowerCase) = "нет" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОSheffilton()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Sheffilton.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Sheffilton"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "UID (числовой формат, без нулей)")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "Остатки, шт")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискSheffilton(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Sheffilton " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки Sheffilton компкресла " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Sheffilton.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ПоискSheffilton(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОМойДвор()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка МойДвор.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки МойДвор"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 1) = Application.Trim(mprice(i, 1))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискМойДвор(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки МойДвор " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка МойДвор.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискМойДвор(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 2))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 6)

        If mostatki(i) <= 0 Or mostatki(i) = "" Then
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОVictoryFit()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Victory Fit.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name
    ActiveWorkbook.Queries.Add name:="feed-yml-0", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""http://victoryfit.ru/wp-content/uploads/feed-yml-0.xml""))," & Chr(13) & "" & Chr(10) & "    shop = Источник{0}[shop]," & Chr(13) & "" & Chr(10) & "    offers = Источник{0}[shop]{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    offer"
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=feed-yml-0" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [feed-yml-0]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "feed_yml_0"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    ncolumn = fNcolumn(mprice, "vendorCode")
    If ncolumn = "none" Then GoTo ended

    For i = 1 To UBound(mprice, 2)
        mprice(i, ncolumn) = Application.Trim(mprice(i, ncolumn))
    Next

    ncolumn1 = fNcolumn(mprice, "Attribute:available")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискVictoryFit(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Victory Fit " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Victory Fit.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискVictoryFit(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        If StrConv(mprice(r, ncolumn1), vbLowerCase) = "true" Then
            mostatki(i) = 30001
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОLuxfire()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Luxfire.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name
    ActiveWorkbook.Queries.Add name:="export_luxfire", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://luxfire.ru/bitrix/catalog_export/export_luxfire.xml""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetimezone}})," & Chr(13) & "" & Chr(10) & "    shop = #""Измененный тип""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}" & _
        ", {""url"", type text}, {""platform"", type text}, {""version"", type date}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""vendor"", type text}, {""model"", type t" & _
        "ext}, {""description"", type text}, {""Attribute:id"", Int64.Type}, {""Attribute:type"", type text}, {""Attribute:available"", type logical}, {""dimensions"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=export_luxfire" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [export_luxfire]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "export_luxfire"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    ncolumn = fNcolumn(mprice, "model")
    If ncolumn = "none" Then GoTo ended

    For i = 1 To UBound(mprice, 2)
        mprice(i, ncolumn) = Application.Trim(mprice(i, ncolumn))
    Next

    ncolumn1 = fNcolumn(mprice, "Attribute:available")
    If ncolumn1 = "none" Then GoTo ended
    
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискLuxfire(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Luxfire " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Luxfire.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискLuxfire(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        If StrConv(mprice(r, ncolumn1), vbLowerCase) = True Then
            mostatki(i) = 30001
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОВПК()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка ВПК.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки ВПК"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    Application.ScreenUpdating = False
    Workbooks(a).ActiveSheet.Cells.ClearOutline
    Application.ScreenUpdating = True
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 5) = Application.Trim(mprice(i, 5))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискВПК(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки ВПК " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки ВПК компкресла" & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка ВПК.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискВПК(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 5))
            On Error GoTo onerror2
            r = r + 1
        Loop
        If IsNumeric(mprice(r, 4)) = True And mprice(r, 4) <> vbNullString Then
            If mprice(r, 4) <= 20 Then
                mostatki(i) = mprice(r, 4)
            ElseIf mprice(r, 4) > 20 And mprice(r, 4) < 100 Then
                mostatki(i) = 999
            ElseIf mprice(r, 4) > 100 Then
                mostatki(i) = 30001
            Else
                mostatki(i) = 30000
            End If
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = mostatki(i)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОUltraGym()
    
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка UltraGym.xlsx"
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Workbooks.Add
    a = ActiveWorkbook.name
    
    ActiveWorkbook.Queries.Add name:="market", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Xml.Tables(Web.Contents(""https://sportholding.ru/index.php?route=extension/payment/yandex_money/market""))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Changed Type""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type " & _
        "text}, {""url"", type text}, {""platform"", type text}, {""version"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(offer,{{""categoryId"", Int64.Type}, {""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""description"", type text}, {""vendor" & _
        """, type text}, {""model"", type text}, {""weight"", Int64.Type}, {""dimensions"", type text}, {""market-sku"", Int64.Type}, {""available"", Int64.Type}, {""Attribute:id"", Int64.Type}, {""Attribute:type"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=market" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [market]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "market"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:O" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    'Поиск номеров столбца с наименованиями
    ncolumn = fNcolumn(mprice, "Attribute:id", 1)
    If ncolumn = 0 Then GoTo ended
    
    ncolumn1 = fNcolumn(mprice, "available", 1)
    If ncolumn1 = 0 Then GoTo ended
    
    ReDim mostatki(UBound(msvyaz, 1))
    
    Call Create_File_Ostatki(abook, asheet)
    
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискUltraGym(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next
    
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки UltraGym " & Date & ".csv", FileFormat:=xlCSV
    End If
    
ended:
    Workbooks("Привязка UltraGym.xlsx").Close False
    Workbooks(a).Close False
    
End Sub

Private Sub ПоискUltraGym(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)
    
    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        If IsNumeric(mprice(r, ncolumn1)) = True Then
            If mprice(r, ncolumn1) > 0 Then
                mostatki(i) = mprice(r, ncolumn1)
            Else
                mostatki(i) = 30000
            End If
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    
End Sub

Private Sub ОRamart()

    Dim msvyaz As Variant, mprice() As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1
    Dim mstock() As Long
    Dim StockBookName As String
    Dim StockSheetName As String

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка RAMART DESIGN.xlsx"
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Ramart Design"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("C9:L" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    ReDim mostatki(UBound(msvyaz, 1))
    ReDim mstock(UBound(msvyaz, 1))
    
    Call Create_File_Ostatki(abook, asheet)
    Workbooks.Add
    StockBookName = ActiveWorkbook.name
    StockSheetName = ActiveSheet.name
    Workbooks(StockBookName).Worksheets(StockSheetName).Cells(1, 1).Value = "Sklad"
    Workbooks(StockBookName).Worksheets(StockSheetName).Cells(1, 2).Value = "Kod1C"
    
    For i = 2 To UBound(msvyaz, 1)
        Call ПоискRamart(i, msvyaz, mprice, abook, asheet, mostatki, mstock, StockBookName, StockSheetName)
    Next
    
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Ramart " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(StockBookName).SaveAs Filename:=Environ("OSTATKI") & "Склады Ramart " & Date & ".csv", FileFormat:=xlCSV
    End If
    
ended:
    Workbooks("Привязка RAMART DESIGN.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ПоискRamart(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef mstock, _
                        ByRef StockBookName, ByRef StockSheetName)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 10))
            On Error GoTo onerror2
            r = r + 1
        Loop
        
        If IsNumeric(mprice(r, 7)) = True Then
            If mprice(r, 7) > 0 Then
                mostatki(i) = mprice(r, 7)
                mstock(i) = 9
            Else
                mostatki(i) = 10000
                mstock(i) = 76
            End If
        Else
            mostatki(i) = 10000
            mstock(i) = 76
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 10000
    mstock(i) = 76

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)
    Workbooks(StockBookName).Worksheets(StockSheetName).Cells(i, 1) = mstock(i)
    Workbooks(StockBookName).Worksheets(StockSheetName).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub ОМебельторг()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Мебельторг\Таблица остатк.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Мебельторг"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    Application.ScreenUpdating = False
    Workbooks(a).ActiveSheet.Cells.ClearOutline
    Application.ScreenUpdating = True
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:J" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 7) = Application.Trim(mprice(i, 7))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискМебельторг(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Мебельторг " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Таблица остатк.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискМебельторг(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 7))
            On Error GoTo onerror2
            r = r + 1
        Loop
        If IsNumeric(mprice(r, 9)) = True And mprice(r, 9) <> vbNullString Then
            mostatki(i) = mprice(r, 9)
        Else
            mostatki(i) = 30000
        End If

        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОGardeck()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Gardeck.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Gardeck"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:AG" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 1) = CStr(Application.Trim(mprice(i, 1)))
    Next
    
    For i = 1 To UBound(msvyaz, 1)
        msvyaz(i, 2) = CStr(msvyaz(i, 2))
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискGardeck(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Gardeck " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Gardeck.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискGardeck(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop
        
        If IsNumeric(mprice(r, 16)) = True And mprice(r, 16) <> vbNullString Then
            If mprice(r, 16) <= 0 Then
                mostatki(i) = 30000
            Else
                mostatki(i) = CLng(mprice(r, 16))
            End If
        Else
            mostatki(i) = 30000
        End If
        
        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОПарнас()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Парнас.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Парнас"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 1) = CStr(Application.Trim(mprice(i, 1)))
    Next
    
    For i = 1 To UBound(msvyaz, 1)
        msvyaz(i, 2) = CStr(Application.Trim(msvyaz(i, 2)))
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискПарнас(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Парнас " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Парнас.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискПарнас(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 1))
            On Error GoTo onerror2
            r = r + 1
        Loop
        If IsNumeric(mprice(r, 4)) = True And mprice(r, 4) <> vbNullString Then
            If mprice(r, 4) <= 0 Then
                mostatki(i) = 30000
            Else
                mostatki(i) = CLng(mprice(r, 4))
            End If
        Else
            mostatki(i) = 30000
        End If
        
        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub

Private Sub ОГородки77()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Городки77.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "Выберите остатки Городки77"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 2)
        mprice(i, 1) = CStr(Application.Trim(mprice(i, 1)))
    Next
    
    For i = 1 To UBound(msvyaz, 1)
        msvyaz(i, 2) = CStr(Application.Trim(msvyaz(i, 2)))
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ПоискГородки77(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "Остатки Городки77 " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("Привязка Городки77.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ПоискГородки77(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> Trim(mprice(r, 2))
            On Error GoTo onerror2
            r = r + 1
        Loop
        
        If IsNumeric(mprice(r, 5)) = True And mprice(r, 5) <> vbNullString Then
            mostatki(i) = CLng(mprice(r, 5))
        Else
            mostatki(i) = 30000
        End If
        
        GoTo ended
    Else
        GoTo onerror2
    End If

onerror2:
    mostatki(i) = 30000

ended:

    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = mostatki(i)

End Sub
