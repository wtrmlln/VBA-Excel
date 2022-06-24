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

    If par = 1 Then MsgBox "������� " & namecol & " �� ������"

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
            If InStr(StrConv(wbk.name, vbLowerCase), "����������") <> 0 Then
                
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
        
            ElseIf InStr(StrConv(wbk.name, vbLowerCase), "�������") <> 0 Then
                
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
        
        Do While Dir(Environ("OSTATKI") & NamesArrayKK(1) & "+concat ���������� " & Date & " " & nof0 & ".csv") <> ""
            nof0 = nof0 + 1
        Loop
        
        Workbooks(ConcatBookNameKK).Worksheets("����1").Cells(1, 1).Value = "category"
        Workbooks(ConcatBookNameKK).SaveAs Filename:=Environ("OSTATKI") & NamesArrayKK(1) & "+concat " & Date & " " & nof0 & ".csv", FileFormat:=xlCSV

    End If

    If NOstatki > 1 Then

        Do While Dir(Environ("OSTATKI") & NamesArray(1) & "+concat " & Date & " " & nof1 & ".csv") <> ""
            nof1 = nof1 + 1
        Loop
        
        Workbooks(ConcatBookName).SaveAs Filename:=Environ("OSTATKI") & NamesArray(1) & "+concat " & Date & " " & nof1 & ".csv", FileFormat:=xlCSV
    
    End If

End Sub

Public Sub �������()

    Set dict2 = CreateObject("Scripting.Dictionary")
    dict2.Add "�103876", 1
    dict2.Add "�103877", 1
    dict2.Add "�103878", 1

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
        Call �������
    End If

    If UserForm1.CheckBox2 = True Then
        Call ���������
    End If

    If UserForm1.CheckBox3 = True Then
        Call �������������
    End If

    If UserForm1.CheckBox4 = True Then
        Call ��������
    End If

    'If UserForm1.CheckBox5 = True Then
    '    Call ������
    'End If

    If UserForm1.CheckBox6 = True Then
        Call �������
    End If

    If UserForm1.CheckBox7 = True Then
        Call �RealChair
    End If

    If UserForm1.CheckBox9 = True Then
        Call �WoodVille
    End If

    If UserForm1.CheckBox10 = True Then
        Call �����������
    End If

    If UserForm1.CheckBox11 = True Then
        Call ��������
    End If

    If UserForm1.CheckBox13 = True Then
        Call �����
    End If

    If UserForm1.CheckBox14 = True Then
        Call ����
    End If

    If UserForm1.CheckBox15 = True Then
        Call �������
    End If

    If UserForm1.CheckBox17 = True Then
        Call �Romana
    End If

    If UserForm1.CheckBox18 = True Then
        Call �Kampfer
    End If

    If UserForm1.CheckBox19 = True Then
        Call ������������
    End If

    If UserForm1.CheckBox21 = True Then
        Call �100�����
    End If

    If UserForm1.CheckBox22 = True Then
        Call �Everproof1
    End If

    If UserForm1.CheckBox23 = True Then
        Call �GardenWay
    End If

    If UserForm1.CheckBox24 = True Then
        Call �ormatek
    End If

    If UserForm1.CheckBox26 = True Then
        Call ��������������
    End If

    If UserForm1.CheckBox27 = True Then
        Call ������
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.CheckBox30 = True Then
        Call �Dekomo
    End If

    If UserForm1.mebff = True Then
        Call ������
    End If

    If UserForm1.Logomebel = True Then
        Call �����������
    End If

    If UserForm1.��������� = True Then
        Call ����������
    End If

    If UserForm1.Avanti = True Then
        Call �Avanti
    End If

    If UserForm1.Sititek = True Then
        Call �Sititek
    End If

    If UserForm1.Aquaplast = True Then
        Call �Aquaplast
    End If

    If UserForm1.Amfitness = True Then
        Call �Amfitness
    End If

    If UserForm1.ZSO = True Then
        Call �ZSO
    End If

    If UserForm1.rtoys = True Then
        Call �rtoys
    End If

    If UserForm1.Family = True Then
        Call �Family
    End If

    If UserForm1.Atemi = True Then
        Call �Atemi
    End If

    If UserForm1.Sparta = True Then
        Call �Sparta
    End If

    If UserForm1.sporthouse = True Then
        Call �sporthouse
    End If

    If UserForm1.Afina = True Then
        Call �Afina
    End If

    If UserForm1.Hudora = True Then
        Call �Hudora
    End If

    If UserForm1.ESF = True Then
        Call �ESF
    End If

    If UserForm1.Kettler = True Then
        Call �kettler
    End If

    If UserForm1.FunDesk = True Then
        Call �fundesk
    End If

    If UserForm1.Ruskresla = True Then
        Call �Ruskresla
    End If

    If UserForm1.mealux = True Then
        Call �mealux
    End If

    If UserForm1.GoodChair = True Then
        Call �GoodChair
    End If

    If UserForm1.������� = True Then
        Call ��������
    End If

    If UserForm1.��_������ = True Then
        Call ���������
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.ISit = True Then
        Call �Isit
    End If

    If UserForm1.Officio = True Then
        Call �Officio
    End If

    If UserForm1.RHome = True Then
        Call �Rhome
    End If

    If UserForm1.����������� = True Then
        Call O�����������
    End If

    If UserForm1.Hasstings = True Then
        Call �Hasttings
    End If

    If UserForm1.Riva = True Then
        Call �Riva
    End If

    If UserForm1.���� = True Then
        Call �����
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.��������� = True Then
        Call ����������
    End If

    If UserForm1.����� = True Then
        Call ������
    End If

    If UserForm1.Unixline = True Then
        Call �Unixline
    End If

    If UserForm1.Intex = True Then
        Call �Intex
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.AllBestGames = True Then
        Call �AllBestGames
    End If

    If UserForm1.CharBroil = True Then
        Call �CharBroil
    End If

    If UserForm1.������ = True Then
        Call �������
    End If

    If UserForm1.��������� = True Then
        Call ����������
    End If

    If UserForm1.Swollen = True Then
        Call �Swollen
    End If

    If UserForm1.���������� = True Then
        Call �����������
    End If

    If UserForm1.���������� = True Then
        Call �����������
    End If

    If UserForm1.FoxGamer = True Then
        Call �FoxGamer
    End If

    If UserForm1.Sheffilton = True Then
        Call �Sheffilton
    End If

    If UserForm1.������� = True Then
        Call ��������
    End If

    If UserForm1.VictoryFit = True Then
        Call �VictoryFit
    End If
    
    If UserForm1.Luxfire = True Then
        Call �Luxfire
    End If
    
    If UserForm1.��� = True Then
        Call ����
    End If
    
    If UserForm1.Ultragym = True Then
        Call �UltraGym
    End If
    
    If UserForm1.Ramart = True Then
        Call �Ramart
    End If
    
    If UserForm1.���������� = True Then
        Call �����������
    End If
    
    If UserForm1.Gardeck = True Then
        Call �Gardeck
    End If
    
    If UserForm1.������ = True Then
        Call �������
    End If
    
    If UserForm1.�������77 = True Then
        Call ��������77
    End If

    If UserForm1.Concatenation = True Then
        Call Concatenation
    End If

    Unload UserForm1
End Sub

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\������(Flexter)\�������\������� �����.xlsx"
    msvyaz = Workbooks("������� �����.xlsx").Worksheets("����1").Range("A1:B" & _
    Workbooks("������� �����.xlsx").Worksheets("����1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.Workbooks.Open "https://dsk-formula.ru/files/uploads/price/DSK-Formula-price-OPT3.xls", ReadOnly
    
    mprice = ActiveWorkbook.Worksheets("�����-����").Range("D2:G" & _
    ActiveWorkbook.Worksheets("�����-����").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(mprice, 1)
        If IsNumeric(mprice(i, 1)) = True Then
            mprice(i, 1) = CLng(mprice(i, 1))
        End If
    Next i

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=" ��.", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� �����.xlsx").Close False
    Workbooks("DSK-Formula-price-OPT3.xls").Close False

End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
    
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
'Private Sub ������()
'
'    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ����� ��������.xlsx"
'    Dim msvyaz As Variant, mprice As Variant, mostatki As Variant, _
'    abook1 As String, asheet1 As String, _
'    a As String, i As Long
'
'    msvyaz = Workbooks("�������� ����� ��������.xlsx").Worksheets("����1").Range("A1:C" & _
'    Workbooks("�������� ����� ��������.xlsx").Worksheets("����1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
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
'        Call ����������(i, msvyaz, mprice, abook1, asheet1, mostatki)
'    Next
'
'    If UserForm1.Autosave = True Then
'        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ����� " & Date & ".csv", FileFormat:=xlCSV
'    End If
'
'    Workbooks("�������� ����� ��������.xlsx").Close False
'    Workbooks(a).Close False
'
'End Sub
'
'Private Sub ����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook1, ByRef asheet1, ByRef mostatki)
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

Private Sub ���������()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\��������\�������\�����.xlsx"
    msvyaz = Workbooks("�����.xlsx").Worksheets("����1").Range("B1:D" & _
    Workbooks("�����.xlsx").Worksheets("����1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ��������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B4:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� �������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
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

Private Sub ��������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\�������\�������\���. �����.xlsx"
    msvyaz = Workbooks("���. �����.xlsx").Worksheets("����1").Range("A1:B" & _
    Workbooks("���. �����.xlsx").Worksheets("����1").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.DisplayAlerts = False
    G = "https://belfortkamin.ru/test/Stok_" & Format(Date, "dd.mm.yyyy") & ".xlsx"
    On Error Resume Next
    Application.Workbooks.Open G
    Application.DisplayAlerts = True

    If ActiveWorkbook.name <> G Then
        MsgBox "�������� ������� �������"
        a = Application.GetOpenFilename
        If a = False Then
            MsgBox "���� �� ������"
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
        Call ������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False

ended:
    Workbooks("���. �����.xlsx").Close False
    Exit Sub
End Sub

Private Sub ������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)
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

Private Sub �������()

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, book1 As String, _
    a, b As Long, mostatki As Variant, abook1 As String, asheet1 As String

    msvyaz = Workbooks("�������� ������.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("�������� ������.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value


    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)
    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
        mostatki(i) = mprice(r, 12)


        If mostatki(i) = 0 Or mostatki(i) = "" Or InStr(mostatki(i), "���������") <> 0 Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "����� 10" Then
            mostatki(i) = 10
        ElseIf mostatki(i) = "���������" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "����������� � ���������" Then
            mostatki(i) = 0
        ElseIf mostatki(i) = "��� � �������" Then
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

Private Sub �RealChair()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� RealChair.xlsx"
    msvyaz = Workbooks("�������� RealChair.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("�������� RealChair.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Real Chair"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����RealChair(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� RealChair " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� RealChair ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� RealChair.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����RealChair(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub �������������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, _
    i As Long, a, b As String, mostatki As Variant, r As Integer, ii As Double, ili As Double

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\����� �������\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "�������� ������� ����� �������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    b = ActiveSheet.name

    '������� ��������� ������ � ����� ��� ��������
    LastRow = Workbooks(a).ActiveSheet.Cells.SpecialCells(xlLastCell).Row


    Call Create_File_Ostatki(abook, asheet)

    '������ �������� �������� �� �������� � 5 �������
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
        Call ������������������(i, msvyaz, mostatki, r, ii, ili, mprice)
    Next
    For i = 2 To UBound(msvyaz, 1)
        Call ��������������������(i, msvyaz, abook, asheet, mostatki, r, ii, ili, mprice)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����� ������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ������������������(ByRef i, ByRef msvyaz, ByRef mostatki, ByRef r, ByRef ii, ByRef ili, ByRef mprice)

    ii = 0
    r = 1

    'msvyaz 2 = ��� "UUID"
    'msvyaz 8 = ��� "��� 1�"
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

Private Sub ��������������������(ByRef i, ByRef msvyaz, ByRef abook, ByRef asheet, ByRef mostatki, _
ByRef r, ByRef ii, ByRef ili, ByRef mprice)

    ili = 0

    r = 1

    'msvyaz 3 = ��� "UUID"
    'msvyaz 9 = ��� "��� 1�"
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
    
        'msvyaz 4 = ��� "UUID"
        'msvyaz 10 = ��� "��� 1�"
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

    'msvyaz 4 = ��� "UUID"
    'msvyaz 10 = ��� "��� 1�"
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
    
        'msvyaz 5 = ��� "UUID"
        'msvyaz 11 = ��� "��� 1�"
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

    'msvyaz 5 = ��� "UUID"
    'msvyaz 11 = ��� "��� 1�"
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
    
        'msvyaz 6 = ��� "UUID"
        'msvyaz 12 = ��� "��� 1�"
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

    'msvyaz 6 = ��� "UUID"
    'msvyaz 12 = ��� "��� 1�"
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

Private Sub �WoodVille()
    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Woodville.xlsx"
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, _
    a, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� WoodVille"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.OpenText Filename:=a, DataType:=xlDelimited, Local:=True
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:AG" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    ReDim mostatki(UBound(msvyaz, 1))

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call �����WoodVille(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� WoodVille " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� WoodVille ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Woodville.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����WoodVille(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub �����������()
    Application.ScreenUpdating = False

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\���� ������\������� ����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Application.Workbooks.Open "https://cvet-mebel.ru/files/uploads/stock/stock.xls"

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ���������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ���� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ���� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ����.xlsx").Close False
    Workbooks("stock.xls").Close False
    Application.ScreenUpdating = True

End Sub

Private Sub ���������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub ��������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\Red-black(�)\�������  ���.xlsx"
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
        Call ������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Red Black " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Red Black ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������  ���.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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


Private Sub �����()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, _
    mostatki As Variant, mprice1 As Variant, mprice2 As Variant, mprice4 As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\����\������� ���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    MsgBox "�������� ������� ����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    
    mprice = Workbooks(a).ActiveSheet.Range("A1:J" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call ���������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="��� ����� ", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="��� ����� ", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ���� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ���� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ���.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ���������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
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
        Or Application.Trim(StrConv(mostatki(i), vbLowerCase)) = "��� �����" Then
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

Private Sub ����()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ��������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False
ended:
    Workbooks("�������� ���.xlsx").Close False

End Sub

Private Sub ��������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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


Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\������\������ ������.xlsx"
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="����� 10", Replacement:="21", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
      
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ ������.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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
        ElseIf mostatki(i) = "����� 10" Then
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

Private Sub �Romana()


    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Romana\������ Romana.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "�������� ������� Romana"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Romana(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Romana " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������ Romana " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ Romana.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����Romana(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub �Kampfer()

    Application.ScreenUpdating = False
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, _
    abook1 As String, asheet1 As String

    Set dict1 = CreateObject("Scripting.Dictionary")

    dict1.Add "�101977", 1
    dict1.Add "�101978", 1
    dict1.Add "�101979", 1
    dict1.Add "�101980", 1
    dict1.Add "�101981", 1
    dict1.Add "�101982", 1
    dict1.Add "�101983", 1
    dict1.Add "�101984", 1
    dict1.Add "�101985", 1
    dict1.Add "�101986", 1
    dict1.Add "�042181", 1
    dict1.Add "�042187", 1
    dict1.Add "�042253", 1
    dict1.Add "�042254", 1
    dict1.Add "�049962", 1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Kampfer\������ ILGC.xlsx"
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
        Call �����Kampfer(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next


    Workbooks(abook).Worksheets(asheet).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=">20", Replacement:="21", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Kampfer " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������ Kampfer " & Date & ".csv", FileFormat:=xlCSV

    End If

    Workbooks("������ ILGC.xlsx").Close False
    Workbooks("price_ilgc-group-dilerskaya-dropshipping-RRC.xls").Close False

    Application.ScreenUpdating = True
End Sub

Private Sub �����Kampfer(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub ������������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a _
    , abook1 As String, asheet1 As String, ncolumn


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������ �����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1), 2)

    MsgBox "�������� ������� ������ �����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B1:L" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "�������� �����")
    If ncolumn = "none" Then GoTo ended


    For i = 2 To UBound(msvyaz, 1)
        Call ����������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn)
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ����� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ����� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������ �����.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub


Private Sub ����������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki _
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

Private Sub �100�����()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\100 �����\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� 100 �����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("F1:R" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����100�����(i, msvyaz, mprice, abook, asheet, mostatki)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� 100����� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub


Private Sub �����100�����(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �Everproof1()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant, _
    abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\Everproof\������� ��.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Everproof MSK"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("C1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended
    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����Everproof1(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Everproof MSK " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Everproof MSK ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ��.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Everproof1(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub �GardenWay()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, mostatki As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\Garden way\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Garden Way"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����GardenWay(i, msvyaz, mprice, abook, asheet, mostatki)
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Garden Way " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����GardenWay(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �ormatek()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Ormatek\������ Ormatek.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Ormatek"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Ormatek(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Ormatek " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ Ormatek.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Ormatek(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub ��������������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\������ ��� ����\������ ������ ��� ����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������ ��� ����"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ������������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ��� ���� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ��� ���� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ ������ ��� ����.xlsx").Close False
    Workbooks(a).Close False

ended:


End Sub

Private Sub ������������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, _
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

Private Sub ������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\�����\������� ���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� �����"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ����� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ���.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

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

Private Sub �������()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, a, _
    abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������\���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("���.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, _
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

Private Sub �Dekomo()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, dict

    Application.Workbooks.Open "\\AKABINET\Doc\����������������\��� �������\SOT ������\���������� ������\������� ��� SOT(������ 938).xlsb"
    msvyaz = Workbooks("������� ��� SOT(������ 938).xlsb").Worksheets("������� ���").Range("Q10:S" & _
    Workbooks("������� ��� SOT(������ 938).xlsb").Worksheets("������� ���").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

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
        Call �����Dekomo(i, msvyaz, mprice, abook, asheet, mostatki, dict)
    Next

    Workbooks(abook).Worksheets(asheet).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="��� �����", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Dekomo " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ��� SOT(������ 938).xlsb").Close False
    Workbooks(nstock).Close False

End Sub

Private Sub �����Dekomo(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef dict)
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

Private Sub ������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\�������� ���-��.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���-��"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, abook2, asheet2)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=">", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
      

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ����� ���������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook2).SaveAs Filename:=Environ("OSTATKI") & "������ ����� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ���-��.xlsx").Close False
    Workbooks(nstock).Close False

ended:

End Sub

Private Sub ����������(ByRef i, ByRef msvyaz, ByRef mprice, _
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
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = IIf(Trim(msvyaz(i, 1)) <> "�103354", mostatki(i), 21)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 3) = IIf(Trim(msvyaz(i, 1)) <> "�103354", mostatki(i), 21)
    Workbooks(abook1).Worksheets(asheet1).Cells(i, 2) = msvyaz(i, 1)
    Workbooks(abook2).Worksheets(asheet2).Cells(i, 1) = msvyaz(i, 3)
    Workbooks(abook2).Worksheets(asheet2).Cells(i, 2) = msvyaz(i, 1)

End Sub

Private Sub �OLSS()

    Dim r, r1, i, mostatki As Variant, abook, asheet, razm1 As Integer, razm2 As Integer

    If Dir(path & "\�������OLSS.xls") = "" Then
        Workbooks.Add
        ActiveWorkbook.ActiveSheet.Cells(1, 1).Value = "kod 1s"
        ActiveWorkbook.SaveAs Filename:=path & "\�������OLSS.xls", FileFormat:=56
    Else
        Application.Workbooks.Open path & "\�������OLSS.xls"
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
        Call �����OLSS(i, mostatki, razm2, abook, asheet)
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
            ActiveWorkbook.SaveAs Filename:=Environ("OSTATKI") & "������� OLSS " & Date & ".csv", FileFormat:=xlCSV
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
            ActiveWorkbook.SaveAs Filename:=Environ("OSTATKI") & "������� OLSS ������ " & Date & ".csv", FileFormat:=xlCSV
        End If
    End If


    Workbooks(abook).Close False

End Sub

Private Sub �����OLSS(ByRef i, ByRef mostatki, ByRef razm2, ByRef abook, ByRef asheet)
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

Private Sub �����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, _
    mostatki As Variant, a, abook1 As String, asheet1 As String, b, mprice1 As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\����������\������� ���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���������� ���"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("F1:P" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���������� ������"

    b = Application.GetOpenFilename
    If b = False Then
        MsgBox "���� �� ������"
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
        Call ���������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    For i = 2 To UBound(msvyaz, 1)
        Call ���������������1(i, msvyaz, mprice1, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ���������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ���������� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ���.xlsx").Close False
    Workbooks(a).Close False
    Workbooks(b).Close False

ended:

End Sub

Private Sub ���������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

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

Private Sub ���������������1(ByRef i, ByRef msvyaz, ByRef mprice1, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

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

Private Sub ����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, _
    mostatki As Variant, mostatkiili As Variant, a

    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\��������� ��������.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ��������� - ������"

    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ��������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("��������� ��������.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ��������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    Dim PortalRow As Long
    Dim OchagRow As Long
    r = 1
    PortalRow = 1
    OchagRow = 1

    '���������������
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

        '�������
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
        '�����
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
        '������ ��������
    Else
        mostatki(i) = 30000
    End If

onerror:
    mostatki(i) = 30000

ended:
    Workbooks(abook).Worksheets(asheet).Cells(i, 1) = msvyaz(i, 1)
    Workbooks(abook).Worksheets(asheet).Cells(i, 2) = IIf(mostatki(i) <= 0, 30000, mostatki(i))

End Sub

Private Sub �Avanti()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\Avanti\������� ���.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Avanti"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Avanti(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="� �������", Replacement:="1", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
       
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="� �������", Replacement:="1", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
       

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Avanti " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Avanti ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ���.xlsx").Close False
    Workbooks(nstock).Close False

ended:

End Sub

Private Sub �����Avanti(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Sititek()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\�������\������ �������.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="TDSheet", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Excel.Workbook(Web.Contents(""https://www.sititek.ru/dealer/Price.xls""), null, true)," & Chr(13) & "" & Chr(10) & "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    TDSheet1"
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
        Call �����Sititek(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Sititek " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Sititek ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ �������.xlsx").Close False
    Workbooks(nstock).Close False
End Sub

Private Sub �����Sititek(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Aquaplast()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\���������\���1.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Aquaplast"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Aquaplast(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Aquaplast " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("���1.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Aquaplast(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Unixline()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\UNIX\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="catalog", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://unixfit.ru/bitrix/catalog_export/catalog.php""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""���������� ���""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""���������� ���1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""url""," & _
    " type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""���������� ���1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""���������� ���2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""vendorCode"", type text}, {""name"", type text}, {""Attribute:id"", Int64.Typ" & _
    "e}, {""Attribute:available"", type logical}, {""picture"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""���������� ���2"""
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
        Call �����Unixline(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Unixline " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(nstock).Close False
    Workbooks("�����.xlsx").Close False

End Sub

Private Sub �����Unixline(ByRef i, ByRef msvyaz, ByRef mprice, _
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


Private Sub �Amfitness()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\American fitness\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Amfitness"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    b = ActiveSheet.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:G" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ncolumn = fNcolumn(mprice, "������������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "���")
    If ncolumn1 = "none" Then GoTo ended

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Amfitness(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Amfitness " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Amfitness(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef b, ByRef ncolumn, _
ByRef ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop


        If Trim(StrConv(mprice(r, ncolumn1), vbUpperCase)) = "����" Then
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

Private Sub �ZSO()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\���\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="full", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Csv.Document(Web.Contents(""https://www.zavodsporta.ru/stockexport/index/full""),[Delimiter="";"",Encoding=1251])," & Chr(13) & "" & Chr(10) & "    #""���������� ���������"" = Table.PromoteHeaders(��������)," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(#""���������� ���������"",{{""�������"", type text}, {""���������"", type text}, {""����"", Int64.Type}, {""����-�" & _
    "��"", type text}, {""������������"", type text}, {""�������"", Int64.Type}, {""�����������"", type text}, {""��������"", type text}, {""��������"", type text}, {""�������������"", type text}, {""URL"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""���������� ���"""
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
        Call �����ZSO(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ZSO " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub �����ZSO(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �rtoys()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\R-toys\�������\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="api_export_3", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://r-toys.ru/bitrix/catalog_export/api_export_3.xml""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""���������� ���""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""���������� ���1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""ur" & _
    "l"", type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""���������� ���1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""���������� ���2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""name"", type text}, {""description"", type text}, {""adult"", type text}," & _
    " {""weight"", type text}, {""Attribute:id"", Int64.Type}, {""Attribute:available"", type logical}, {""age"", type text}})," & Chr(13) & "" & Chr(10) & "    #""����������� ������� barcode"" = Table.ExpandTableColumn(#""���������� ���2"", ""barcode"", {""Element:Text""}, {""barcode.Element:Text""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""����������� ������� barcode"""
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
        Call �����Rtoys(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Rtoys " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub �����Rtoys(ByRef i, ByRef msvyaz, ByRef mprice, _
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



Private Sub �Family()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\������\������ ������.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Family"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    For i = 2 To UBound(msvyaz, 1)
        Call �����Family(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Family " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ ������.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Family(ByRef i, ByRef msvyaz, ByRef mprice, _
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


Private Sub �Atemi()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\�������� �����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Atemi"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Atemi(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Atemi " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� �����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Atemi(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Sparta()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\������\�������\����� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Sparta(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("����� ������.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Sparta(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �sporthouse()

    Application.ScreenUpdating = False

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\�����\sporthouse\�����.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="yml (2)", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://sportpro.ru/yml""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    #""����������� ������� shop"" = Table.ExpandTableColumn(#""���������� ���"", ""shop"", {""offers""}, {""shop.offers""})," & Chr(13) & "" & Chr(10) & "    #""shop offers"" = #""����������� ������� shop""{0}[shop.of" & _
    "fers]," & Chr(13) & "" & Chr(10) & "    #""����������� ������� offer"" = Table.ExpandTableColumn(#""shop offers"", ""offer"", {""Attribute:available"", ""vendorCode""}, {""offer.Attribute:available"", ""offer.vendorCode""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""����������� ������� offer"""
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
        Call �����sporthouse(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Sporthouse " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�����.xlsx").Close False
    Workbooks(nstock).Close False

    Application.ScreenUpdating = True

End Sub

Private Sub �����sporthouse(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Afina()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\�������� �����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.OpenText Filename:="http://afinalux.ru/clients/import.csv", DataType:=xlDelimited, Local:=True

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("D1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Afina(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Afina " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� �����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Afina(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Hudora()
    Application.ScreenUpdating = False


    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Hudora\������ Hudora.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add

    nstock = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="������1", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://hudoraworld.ru/wholesale.xml""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    #""����������� ������� shop"" = Table.ExpandTableColumn(#""���������� ���"", ""shop"", {""offers""}, {""shop.offers""})," & Chr(13) & "" & Chr(10) & "    #""����������� ������� shop.offers"" = Table.E" & _
    "xpandTableColumn(#""����������� ������� shop"", ""shop.offers"", {""offer""}, {""shop.offers.offer""})," & Chr(13) & "" & Chr(10) & "    #""����������� ������� shop.offers.offer"" = Table.ExpandTableColumn(#""����������� ������� shop.offers"", ""shop.offers.offer"", {""vendorCode"", ""Attribute:available""}, {""shop.offers.offer.vendorCode"", ""shop.offers.offer.Attribute:available""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "  " & _
    "  #""����������� ������� shop.offers.offer"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=������1" _
    , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [������1]")
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
        .ListObject.DisplayName = "������1"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
 

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Hudora(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Hudora " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ Hudora.xlsx").Close False
    Workbooks(nstock).Close False
    Application.ScreenUpdating = True

End Sub

Private Sub �����Hudora(ByRef i, ByRef msvyaz, ByRef mprice, _
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


Private Sub �ESF()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\�������� ���.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ESF"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "��������� �������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����ESF(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ESF " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ESF ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ���.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����ESF(ByRef i, ByRef msvyaz, ByRef mprice, _
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
        ElseIf mostatki(i) = "� �������" Then
            mostatki(i) = 50
        ElseIf IsDate(mostatki(i)) = True Then
            mostatki(i) = 30000
        End If

        If InStr(StrConv(mostatki(i), vbLowerCase), "�����") <> 0 Then
            mostatki(i) = Application.Trim(Replace(StrConv(mostatki(i), vbLowerCase), "�����", vbNullString))
            mostatki(i) = CLng(mostatki(i))
        End If

        If InStr(mostatki(i), "��") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), "��", vbNullString))
        End If

        If InStr(mostatki(i), "���") <> 0 Then
            mostatki(i) = Application.Trim(Replace(mostatki(i), "���", vbNullString))
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
Private Sub �kettler()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\kettler\�������\����� kettler.xlsx"
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
        Call �����kettler(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, b, ncolumn, ncolumn1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� kettler " & Date & ".csv", FileFormat:=xlCSV
    End If



ended:
Workbooks("����� kettler.xlsx").Close False
If nstock <> vbNullString Then
    Workbooks(nstock).Close False
End If
End Sub

Private Sub �����kettler(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �fundesk()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\FunDesk\�����\������� Fundesk.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Fundesk"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����Fundesk(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Fundesk " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Fundesk ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� Fundesk.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Fundesk(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Ruskresla()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ������\���1.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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


    ncolumn = fNcolumn(mprice, "������������.���")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "��������� �������", 2)
    If ncolumn1 = "none" Then
        ncolumn1 = fNcolumn(mprice, "����������")
        If ncolumn1 = "none" Then GoTo ended
    End If

    For i = 2 To UBound(msvyaz, 1)
        Call �����Ruskresla(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, abook2, asheet2)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook2).SaveAs Filename:=Environ("OSTATKI") & "����� ������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("���1.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Ruskresla(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �mealux()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Mealux\������ Mealux.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    abook2 = ActiveWorkbook.name

    MsgBox "�������� ������� Mealux"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����mealux(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� mealux " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� mealux ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("������ Mealux.xlsx").Close False
    Workbooks(a).Close False
    Workbooks(abook2).Close False

ended:

End Sub

Private Sub �����mealux(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" And msvyaz(i, 1) <> "-" Then
        Do While Trim(msvyaz(i, 1)) <> Trim(mprice(r, ncolumn))
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)



        If mostatki(i) = "����" Then
            mostatki(i) = 21
        ElseIf mostatki(i) = "���������" Then
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

Private Sub �GoodChair()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ������(o)\�����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = Workbooks(a).ActiveSheet.Range("B5:E" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "����� / ������������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����Goodchair(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="���", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="���", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Goodchair(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub ��������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\�������\SOT.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� �������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value


    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "��������")
    If ncolumn = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call ������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("SOT.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ������������(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn)

    Dim r As Long, a() As String

    r = 1

    If msvyaz(i, 1) <> 0 And msvyaz(i, 1) <> "" Then
        Do While msvyaz(i, 1) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = StrConv(CStr(mprice(r, ncolumn)), vbLowerCase)
        
        If mostatki(i) = "����������" Then
            mostatki(i) = 20
        ElseIf mostatki(i) = "����� �����" Then
            mostatki(i) = 20
        ElseIf mostatki(i) = "����" Then
            mostatki(i) = 5
        ElseIf mostatki(i) = "���" Then
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

Private Sub ���������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1 As String

    Dim naimen() As Variant
    Dim naimenconcat() As Variant

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\�.�.������\�����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� �.�.������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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

    If fNcolumn(mprice, "������������ ������") <> 1 And fNcolumn(mprice, "�����������") <> 2 Then
        GoTo ended
    End If

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn1 = fNcolumn(mprice, "��������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, naimenconcat, ncolumn1)
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� �.�.������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� �.�.������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("�����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �������������(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, asheet2 As String

    Application.Workbooks.Open "\\Akabinet\doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:H" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn1 = fNcolumn(mprice, "��������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ������ ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Isit()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, a, _
    mostatki As Variant, mprice1 As Variant, mprice2 As Variant, _
    abook1 As String, asheet1 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\i-sit\������� ���+�� �������.xlsx"
    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� I-sit"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Isit(i, msvyaz, mprice, abook, asheet, mostatki, mprice1, mprice2, abook1, asheet1)
    Next

    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="��� � �������", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="��� � �������", Replacement:="30000", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook).Activate
    Workbooks(abook).Worksheets(asheet).Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="� �������", Replacement:="30001", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Workbooks(abook1).Activate
    Workbooks(abook1).Worksheets(asheet1).Range("C2:C" & UBound(msvyaz, 1)).Select
    Selection.Replace What:="� �������", Replacement:="30001", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� I-sit " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� I-sit ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ���+�� �������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����Isit(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
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

Private Sub �Officio()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\Officio\�������� Officio.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("C1:D" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Officio"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("D1:G" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)


    ncolumn = fNcolumn(mprice, "������������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "��������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����Officio(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Officio " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Officio ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("�������� Officio.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Officio(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Rhome()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� R-HOME.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� R-HOME"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:C" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "���")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "� �������")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����RHome(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� R-Home " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� R-Home ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("�������� R-HOME.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����RHome(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub O�����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� �����������.xlsx"

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
        Call ����������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks(a).Close False
    Workbooks("�������� �����������.xlsx").Close False
End Sub

Private Sub ����������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �Hasttings()
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Hasting.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Hasttings"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name

    mprice = Workbooks(a).ActiveSheet.Range("A1:L" & _
    Workbooks(a).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Hasttings(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Hasttings " & Date & ".csv", FileFormat:=xlCSV
    End If

    Application.CutCopyMode = False
    Workbooks("�������� Hasting.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����Hasttings(ByRef i, ByRef msvyaz, ByRef mprice, _
ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long
    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" Then

        Do While msvyaz(i, 2) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop
    
        mostatki(i) = mprice(r, 12)
    
        If mostatki(i) = "���" Or mostatki(i) = "���" Then
            mostatki(i) = 30000
        ElseIf mostatki(i) = "����" Or mostatki(i) = "����" Then
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

Private Sub �Riva()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������ Riva\�������� Riva.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Riva"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A3:R" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "������������.���")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "������ �����")
    If ncolumn1 = "none" Then
    End If

    For i = 2 To UBound(msvyaz, 1)
        Call �����Riva(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, abook2, asheet2)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Riva " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Riva ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Riva.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Riva(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �����()
    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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

    ncolumn = fNcolumn(mprice, "��������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "���c���")
    If ncolumn1 = "none" Then GoTo ended

    ncolumn2 = fNcolumn(mprice, "�������")
    If ncolumn2 = "none" Then GoTo ended

    ncolumn3 = fNcolumn(mprice, "�������")
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
        Call ���������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, ncolumn2, ncolumn3, mconcat)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ���� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ����.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub ���������(ByRef i, ByRef msvyaz, ByRef mprice, _
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
    
        If Trim(StrConv(mostatki(i), vbLowerCase)) = "��� � �������" Then
            mostatki(i) = 30000
        ElseIf Trim(StrConv(mostatki(i), vbLowerCase)) = "� �������" Then
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

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1, abook2 As String, asheet2 As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������� �������\�������\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("B1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="TDSheet (2)", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Excel.Workbook(Web.Contents(""ftp://178.248.71.47/OSTATKI_SPORT.xls""), null, true)," & Chr(13) & "" & Chr(10) & "    TDSheet1 = ��������{[Name=""TDSheet""]}[Data]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    TDSheet1"
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
        
        If mprice(1, i) = "�������" Or mprice(2, i) = "�������" Or mprice(3, i) = "�������" _
        Or mprice(1, i) = "������� ���" Or mprice(2, i) = "������� ���" Or mprice(3, i) = "������� ���" Then
            ncolumn = i
        End If
    
        If mprice(2, i) & " " & mprice(3, i) = "��������� ��������� �������" Then
            ncolumn1 = i
        End If
        
        If mprice(2, i) & " " & mprice(3, i) = "���� ��������� �������" Then
            ncolumn2 = i
        End If
    
    Next i
    
    If ncolumn = 0 Then
        MsgBox "������� ������� �� ������"
        GoTo ended
    End If
    
    If ncolumn1 = 0 Then
        MsgBox "������� ��������� ��������� ������� �� ������"
        GoTo ended
    End If
    
    If ncolumn2 = 0 Then
        MsgBox "������� ���� ��������� ������� �� ������"
        GoTo ended
    End If
    
    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1, ncolumn2)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, _
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
Private Sub �������()

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    '������� ����� ������ � �����, ���� �� ������������
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

    ncolumn = fNcolumn(mprice, "��������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "����")
    If ncolumn1 = "none" Then GoTo ended

    ncolumn2 = fNcolumn(mprice, "������� ���")
    If ncolumn2 = "none" Then GoTo ended

    ncolumn3 = fNcolumn(mprice, "������� ���")
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, _
        ncolumn1, ncolumn2, ncolumn3, mconcat)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub
Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub ����������()

    Dim r As Long

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ���������.xlsx"

    stock_name = ActiveWorkbook.name

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A2:H" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
            If InStr(msvyaz(i, 1), "��������") <> 0 Then
            
                msvyaz(i, 1) = Replace(msvyaz(i, 1), "��������", "")
            
                '������������ ���������� ������ ��� ���������
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
                
                    '������� ���������� ����������
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
            
                'mprice(r, 6) - ��������� ������� ����� ����������
                'mprice(r, 10) - ��������� ������� ����� ���
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
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ���������.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub ������()

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\�����\������� �����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    MsgBox "�������� ������� �����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ����������(i, msvyaz, mprice, abook, asheet, abook1, asheet1, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ����� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� �����.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub
Private Sub ����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, _
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

Private Sub �Intex()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\Intex\������ Intex.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz)
        If IsNumeric(msvyaz(i, 1)) = True Then
            msvyaz(i, 1) = CStr(msvyaz(i, 1))
        End If
    Next i

    MsgBox "�������� ������� Intex"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("D1:N" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call �����Intex(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next


    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Intex " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������ Intex.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Intex(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ncolumn, ncolumn1)

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

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("B2:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)


    For i = 2 To UBound(msvyaz, 1)
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �AllBestGames()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� AllBestGames.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� AllBestGames"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    For Each Sh In Workbooks(a).Sheets
        If Application.Trim(StrConv(Sh.name, vbLowerCase)) = "������ ��� �����" Then
            SheetName = Sh.name
        End If
    Next Sh

    If SheetName = sEmpty Then
        MsgBox "���� '������ ��� �����' � ������ �� ������"
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
        Call �����AllBestGames(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� AllBestGames " & Date & ".csv", FileFormat:=xlCSV
    End If


    Workbooks("�������� AllBestGames.xlsx").Close False
    Workbooks(a).Close False
ended:
End Sub

Private Sub �����AllBestGames(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, 1)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, 3)
    
        If mostatki(i) = "� �������" Then
            mostatki(i) = 30001
        ElseIf mostatki(i) = "��� � �������" Then
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

Private Sub �CharBroil()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� CharBroil.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name

    ActiveWorkbook.Queries.Add name:="yandex_896364", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://dealer.clmbs.ru/bitrix/catalog_export/yandex_896364.php?1623166096""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""���������� ���""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""���������� ���1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company""," & _
    " type text}, {""url"", type text}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""���������� ���1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    offer"
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
        Call �����CharBroil(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� CharBroil " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� Charbroil.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub �����CharBroil(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)

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

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub ����������()

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ���������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
                If Application.Trim(ConcatArray(i, 1)) = "������" Then
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
            If Application.Trim(ConcatArray(i, 1)) = "������" Then
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
        Call ��������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� ���������.xlsx").Close False
    Workbooks(fileostatki_name).Close False


End Sub

Private Sub ��������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(0, r)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(1, r)
       
        If Trim(StrConv(mostatki(i), vbLowerCase)) = "��� � �������" Then
            mostatki(i) = 30000
        ElseIf Trim(StrConv(mostatki(i), vbLowerCase)) = "� �������" Then
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

Private Sub �Swollen()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Swollen.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Swollen"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Swollen(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� SWOLLEN " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� Swollen.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub �����Swollen(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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
        ElseIf mostatki(i) = "����� 10" Then
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

Private Sub �����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ���������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "�������� ������� ���������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ���������������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ��������� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� ���������.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub ���������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

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

Private Sub �����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ����� �����.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "�������� ������� ����� �����"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("B4:G" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    '������������ - C(2)
    '������� - E(4)

    For i = 1 To UBound(mprice, 2)
        mprice(i, 2) = Application.Trim(mprice(i, 2))
    Next

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)

    For i = 2 To UBound(msvyaz, 1)
        Call ���������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ����� ����� " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� ����� �����.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub ���������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �FoxGamer()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� FoxGamer.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For i = 1 To UBound(msvyaz, 2)
        msvyaz(i, 2) = Application.Trim(msvyaz(i, 2))
    Next

    MsgBox "�������� ������� FoxGamer"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:F" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ncolumn = fNcolumn(mprice, "�������")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������")
    If ncolumn1 = "none" Then GoTo ended

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki(abook1, asheet1)

    For i = 2 To UBound(msvyaz, 1)
        Call �����FoxGamer(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� FoxGamer " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� FoxGamer ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

ended:
    Workbooks("�������� FoxGamer.xlsx").Close False
    Workbooks(a).Close False


End Sub

Private Sub �����FoxGamer(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1, ByRef ncolumn, ByRef ncolumn1)

    Dim r As Long

    r = 1

    If msvyaz(i, 2) <> 0 And msvyaz(i, 2) <> "" And msvyaz(i, 2) <> "-" Then
        Do While Trim(msvyaz(i, 2)) <> mprice(r, ncolumn)
            On Error GoTo onerror2
            r = r + 1
        Loop

        mostatki(i) = mprice(r, ncolumn1)
    
        If StrConv(CStr(mostatki(i)), vbLowerCase) = "��" Then
            mostatki(i) = 5
        ElseIf StrConv(CStr(mostatki(i)), vbLowerCase) = "���" Then
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

Private Sub �Sheffilton()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1


    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Sheffilton.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Sheffilton"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
        GoTo ended:
    End If
    Workbooks.Open Filename:=a

    a = ActiveWorkbook.name
 
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:I" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    ReDim mostatki(UBound(msvyaz, 1))

    Call Create_File_Ostatki(abook, asheet)
    Call Create_File_Ostatki_KK(abook1, asheet1)

    ncolumn = fNcolumn(mprice, "UID (�������� ������, ��� �����)")
    If ncolumn = "none" Then GoTo ended

    ncolumn1 = fNcolumn(mprice, "�������, ��")
    If ncolumn1 = "none" Then GoTo ended

    For i = 2 To UBound(msvyaz, 1)
        Call �����Sheffilton(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Sheffilton " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� Sheffilton ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Sheffilton.xlsx").Close False
    Workbooks(a).Close False

ended:

End Sub

Private Sub �����Sheffilton(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub ��������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� �������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� �������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� �������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ������������(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �VictoryFit()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Victory Fit.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name
    ActiveWorkbook.Queries.Add name:="feed-yml-0", Formula:= _
    "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""http://victoryfit.ru/wp-content/uploads/feed-yml-0.xml""))," & Chr(13) & "" & Chr(10) & "    shop = ��������{0}[shop]," & Chr(13) & "" & Chr(10) & "    offers = ��������{0}[shop]{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    offer"
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
        Call �����VictoryFit(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Victory Fit " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Victory Fit.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����VictoryFit(ByRef i, ByRef msvyaz, ByRef mprice, _
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

Private Sub �Luxfire()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Luxfire.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Workbooks.Add
    a = ActiveWorkbook.name
    ActiveWorkbook.Queries.Add name:="export_luxfire", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    �������� = Xml.Tables(Web.Contents(""https://luxfire.ru/bitrix/catalog_export/export_luxfire.xml""))," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(��������,{{""Attribute:date"", type datetimezone}})," & Chr(13) & "" & Chr(10) & "    shop = #""���������� ���""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""���������� ���1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}" & _
        ", {""url"", type text}, {""platform"", type text}, {""version"", type date}})," & Chr(13) & "" & Chr(10) & "    offers = #""���������� ���1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""���������� ���2"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""vendor"", type text}, {""model"", type t" & _
        "ext}, {""description"", type text}, {""Attribute:id"", Int64.Type}, {""Attribute:type"", type text}, {""Attribute:available"", type logical}, {""dimensions"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""���������� ���2"""
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
        Call �����Luxfire(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Luxfire " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Luxfire.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����Luxfire(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, _
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

Private Sub ����()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ���.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ���"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ��������(i, msvyaz, mprice, abook, asheet, mostatki, abook1, asheet1)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ��� " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "������� ��� ����������" & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ���.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ��������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef abook1, ByRef asheet1)

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

Private Sub �UltraGym()
    
    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� UltraGym.xlsx"
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Workbooks.Add
    a = ActiveWorkbook.name
    
    ActiveWorkbook.Queries.Add name:="market", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Xml.Tables(Web.Contents(""https://sportholding.ru/index.php?route=extension/payment/yandex_money/market""))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Changed Type""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""���������� ���"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type " & _
        "text}, {""url"", type text}, {""platform"", type text}, {""version"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""���������� ���""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""���������� ���1"" = Table.TransformColumnTypes(offer,{{""categoryId"", Int64.Type}, {""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""description"", type text}, {""vendor" & _
        """, type text}, {""model"", type text}, {""weight"", Int64.Type}, {""dimensions"", type text}, {""market-sku"", Int64.Type}, {""available"", Int64.Type}, {""Attribute:id"", Int64.Type}, {""Attribute:type"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""���������� ���1"""
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
    
    '����� ������� ������� � ��������������
    ncolumn = fNcolumn(mprice, "Attribute:id", 1)
    If ncolumn = 0 Then GoTo ended
    
    ncolumn1 = fNcolumn(mprice, "available", 1)
    If ncolumn1 = 0 Then GoTo ended
    
    ReDim mostatki(UBound(msvyaz, 1))
    
    Call Create_File_Ostatki(abook, asheet)
    
    For i = 2 To UBound(msvyaz, 1)
        Call �����UltraGym(i, msvyaz, mprice, abook, asheet, mostatki, ncolumn, ncolumn1)
    Next
    
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� UltraGym " & Date & ".csv", FileFormat:=xlCSV
    End If
    
ended:
    Workbooks("�������� UltraGym.xlsx").Close False
    Workbooks(a).Close False
    
End Sub

Private Sub �����UltraGym(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef ncolumn, ByRef ncolumn1)
    
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

Private Sub �Ramart()

    Dim msvyaz As Variant, mprice() As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1
    Dim mstock() As Long
    Dim StockBookName As String
    Dim StockSheetName As String

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� RAMART DESIGN.xlsx"
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Ramart Design"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Ramart(i, msvyaz, mprice, abook, asheet, mostatki, mstock, StockBookName, StockSheetName)
    Next
    
    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Ramart " & Date & ".csv", FileFormat:=xlCSV
        Workbooks(StockBookName).SaveAs Filename:=Environ("OSTATKI") & "������ Ramart " & Date & ".csv", FileFormat:=xlCSV
    End If
    
ended:
    Workbooks("�������� RAMART DESIGN.xlsx").Close False
    Workbooks(a).Close False

End Sub

Private Sub �����Ramart(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki, ByRef mstock, _
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

Private Sub �����������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\���������� �����\0������\����������\������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ����������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ���������������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ���������� " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ���������������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �Gardeck()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� Gardeck.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� Gardeck"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����Gardeck(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� Gardeck " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� Gardeck.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����Gardeck(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub �������()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� ������.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� ������"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call �����������(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� ������ " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� ������.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub �����������(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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

Private Sub ��������77()

    Dim msvyaz As Variant, mprice As Variant, abook As String, asheet As String, i As Long, mostatki As Variant, nstock, _
    mprice1 As Variant, abook1 As String, asheet1 As String, ncolumn, ncolumn1

    Application.Workbooks.Open "\\AKABINET\Doc\���������� ������\������\������� ���\��������\�������� �������77.xlsx"

    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    MsgBox "�������� ������� �������77"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "���� �� ������"
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
        Call ������������77(i, msvyaz, mprice, abook, asheet, mostatki)
    Next

    If UserForm1.Autosave = True Then
        Workbooks(abook).SaveAs Filename:=Environ("OSTATKI") & "������� �������77 " & Date & ".csv", FileFormat:=xlCSV
    End If

    Workbooks("�������� �������77.xlsx").Close False
    Workbooks(a).Close False

ended:
End Sub

Private Sub ������������77(ByRef i, ByRef msvyaz, ByRef mprice, ByRef abook, ByRef asheet, ByRef mostatki)

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
