Attribute VB_Name = "Module1"
Public mproschet As Variant, dict, dict1, dict2, msvyaz As Variant, mprice As Variant, abook As String, _
asheet As String, i As Long, book1 As String, _
a, b As Long, mostatki As Variant, abook1 As String, asheet1 As String, otv, xx, uf1 As UserForm2, _
uf2 As UserForm2, ncolumn, ncolumn1, ncolumn2, a1, mprice1 As Variant, xx1, ncolumn3, ncolumn4

Private Function fNcolumn(source, namecol, Optional par As Integer = 1)

If par = 1 Then

    For i1 = 1 To UBound(source, 1)
        For i3 = 1 To UBound(source, 2)
            If Trim(source(i1, i3)) = namecol Then
                fNcolumn = i3
                Exit Function
            End If
        Next i3
    Next i1

Else

    For i3 = 1 To UBound(source, 2)
    
        For i1 = 1 To UBound(source, 1)
    
            If Trim(source(i1, i3)) = namecol Then
    
                fNcolumn = i3
            Exit Function
            End If
        Next i1
    
    Next i3
End If

fNcolumn = "none"

If par = 1 Then MsgBox "Столбец " & namecol & " не найден"

End Function


Private Sub CreateSheet(name As String)
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.name = name
End Sub

Public Sub Rnw()


Dim i2 As Long
Dim forced As Boolean
Dim m_postavshik
Dim m_update_date
Dim m_kody_1c

With Range("A10", Cells(Rows.Count, 1).End(xlUp))
 m_postavshik = WorksheetFunction.Transpose(.Value)
End With

For i = 1 To UBound(m_postavshik)
    m_postavshik(i) = Trim(m_postavshik(i))
Next i

otv = 0

abook = ActiveWorkbook.name

For Each wks In Application.Worksheets
    If wks.name = "ДатыО" Then
        GoTo start
    End If
Next wks

Call CreateSheet("ДатыО")

start:

'Создает массив с названиями поставщиков и датами обновления из листа "ДатыО"
LastRow_Date = Workbooks(abook).Worksheets("ДатыО").UsedRange.Rows.Count
m_update_date = Workbooks(abook).Worksheets("ДатыО").Range("A1:B" & _
LastRow_Date).Value

For i = 1 To UBound(m_update_date)
    m_update_date(i, 1) = Trim(m_update_date(i, 1))
    m_update_date(i, 2) = Trim(m_update_date(i, 2))
    If m_update_date(i, 2) <> "" And m_update_date(i, 2) <> 0 Then
       m_update_date(i, 2) = CDate(m_update_date(i, 2))
    End If
Next i

'Создает массив кодов 1С из просчете без лишних пробелов
With Range("Q10", Cells(Rows.Count, 17).End(xlUp))
 m_kody_1c = WorksheetFunction.Transpose(.Value)
End With

For i = 1 To UBound(m_kody_1c)
    m_kody_1c(i) = Trim(m_kody_1c(i))
Next i

If MsgBox("Принудительно обновить?", vbYesNo) = vbYes Then
    forced = True
    UserForm1.Show
    
    If UserForm1.CheckBox1 = True Then
        Call оТетчер(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox2 = True Then
        Call оДриада(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    'If UserForm1.CheckBox3 = True Then
    '    Call о100печей(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    'End If
    
    If UserForm1.CheckBox4 = True Then
        Call оUNIX(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox5 = True Then
        Call оSwollen(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox6 = True Then
        Call oАвтовентури(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox7 = True Then
        Call oDekomo(i2, m_update_date, forced)
    End If
    
    If UserForm1.CheckBox8 = True Then
        Call oАфина(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox9 = True Then
        Call oWeekend(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox10 = True Then
        Call oVictoryFit(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox11 = True Then
        Call oWoodville(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox12 = True Then
        Call oПромет(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
'    If UserForm1.CheckBox13 = True Then
'        Call oUltragym(i2, m_postavshik, m_update_date, m_kody_1c, forced)
'    End If
    
    If UserForm1.CheckBox14 = True Then
        Call oLuxfire(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox15 = True Then
        Call oAvalon(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
    
    If UserForm1.CheckBox16 = True Then
        Call oSporthouse(i2, m_postavshik, m_update_date, m_kody_1c, forced)
    End If
        
Else
    forced = False
    For i2 = 1 To UBound(m_update_date)
        If m_update_date(i2, 1) = "Тетчер" Then
            Call оТетчер(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Дриада" Or m_postavshik(i2) = "Дриада-спорт" Then
            Call оДриада(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        
        ''''100 печей'''''
'        ElseIf m_postavshik(i2) = "100 печей" Then
'        Call о100печей(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ''''100 печей'''''
        
        ElseIf m_update_date(i2, 1) = "UNIX" Then
            Call оUNIX(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Swollen" Then
            Call оSwollen(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Автовентури" Then
            Call oАвтовентури(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Афина" Then
            Call oАфина(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Weekend" Then
           Call oWeekend(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Dekomo" Then
            Call oDekomo(i2, m_update_date, forced)
        ElseIf m_update_date(i2, 1) = "Victory Fit" Then
            Call oVictoryFit(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Woodville" Then
            Call oWoodville(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Промет" Then
            Call oПромет(i2, m_postavshik, m_update_date, m_kody_1c, forced)
'        ElseIf m_update_date(i2, 1) = "Ultragym" Then
'            Call oUltragym(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Luxfire" Then
            Call oLuxfire(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "AVALON" Then
            Call oAvalon(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        ElseIf m_update_date(i2, 1) = "Sporthouse" Then
            Call oSporthouse(i2, m_postavshik, m_update_date, m_kody_1c, forced)
        End If
    
    Next i2
End If

If DatePart("d", Now) = 1 And ActiveWorkbook.name <> "Просчет цен B2B.xlsb" Then
    Call JU_Check
End If

Unload UserForm1

End Sub

Private Sub JU_Check()

If MsgBox("Проверить коды с ЖУ?", vbYesNo) = vbYes Then
    
    Application.ScreenUpdating = False
    
    For Each wks In Application.Worksheets
        If wks.name = "ЖУ" Then
            WorksheetFinded = True
            Exit For
        End If
    Next wks
    
    If WorksheetFinded = False Then
        Call CreateSheet("ЖУ")
    End If
    
    For Each pq In ThisWorkbook.Queries
        pq.Delete
    Next
    ActiveWorkbook.Worksheets("ЖУ").Activate
    ActiveWorkbook.Worksheets("ЖУ").Cells.Clear
    
    Application.ScreenUpdating = True
    MsgBox "Выберите выгрузку кодов 1С с ЖУ"
    ExportFileName = Application.GetOpenFilename
    SourceFormula = "Источник = Csv.Document(File.Contents(" & Chr(34) & ExportFileName & Chr(34)
    Application.ScreenUpdating = False
    
    On Error GoTo ended
    ActiveWorkbook.Queries.Add name:="data_export (74)", Formula:= _
        "let" & Chr(10) & SourceFormula & "),[Delimiter="";"",Encoding=65001])," & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Column1"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Удалены пустые строки"" = Table.SelectRows(#""Измененный тип"", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {" & _
        """"", null})))," & Chr(13) & "" & Chr(10) & "    #""Удаленные дубликаты"" = Table.Distinct(#""Удалены пустые строки"")" & Chr(10) & "in" & Chr(10) & "    #""Удаленные дубликаты"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""data_export (74)""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [data_export (74)]")
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
        .ListObject.DisplayName = "data_export__74"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    ActiveWorkbook.Queries("data_export (74)").Delete
    
    Dim Codes()
    Codes = ActiveWorkbook.Worksheets("Просчет цен").Range("Q1:Q" & _
    ActiveWorkbook.Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 10 To UBound(Codes)
        If Not dict.exists(Codes(i, 1)) Then
            dict.Add Codes(i, 1), i
        End If
    Next
    
    Dim CodesArray()
    CodesArray = ActiveWorkbook.Worksheets("ЖУ").Range("A3:A" & _
    ActiveWorkbook.Worksheets("ЖУ").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Dim coll As New Collection
    
    For i = 1 To UBound(CodesArray)
        If CodesArray(i, 1) = "" Or Not dict.exists(CodesArray(i, 1)) Then GoTo continue
        
        If dict.exists(CodesArray(i, 1)) Then
            Index = dict.Item(CodesArray(i, 1))
            If ActiveWorkbook.Worksheets("Просчет цен").Cells(Index, 14) = 0 Then
                ActiveWorkbook.Worksheets("Просчет цен").Cells(Index, 14) = 0.01
                coll.Add (CodesArray(i, 1))
            End If
        End If
continue:
    Next
    
    If coll.Count > 0 Then
        UserForm2.TextBox1.Text = coll(1)
    End If
    For i = 2 To coll.Count
        UserForm2.TextBox1.Text = UserForm2.TextBox1.Text & Chr(13) & coll(i)
    Next
    
    ActiveWorkbook.Worksheets("Просчет цен").Activate
    
    UserForm2.Show

End If

ended:
For Each wks In Application.Worksheets
    If wks.name = "ЖУ" Then
        ActiveWorkbook.Worksheets("ЖУ").Delete
    End If
Next wks


Application.ScreenUpdating = True

End Sub


Private Sub оТетчер(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)

For i = 1 To UBound(m_postavshik)
    
    
    'Если принудительное обновление - то сразу старт
    'Если непринудительное обновление,
    'и текущая дата > даты в просчете
    'и текущий месяц > месяца в просчете
    'ИЛИ если текущий день месяца >= 16, а день месяца в просчете < 16, то старт
    'Если не подходит под оба условия - выход
    If m_postavshik(i) = "Тетчер" Then
        If forced = True Then
            Exit For
        ElseIf _
        Date > m_update_date(nomer, 2) And _
        DatePart("m", Now) > DatePart("m", m_update_date(nomer, 2)) Or _
        Date > m_update_date(nomer, 2) And _
        DatePart("d", Now) >= 16 And DatePart("d", m_update_date(nomer, 2)) < 16 Or _
        m_update_date(nomer, 2) = sEmpty Then
            If MsgBox("Обновить Тетчер?", vbYesNo) = vbNo Then Exit Sub
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i


    i1 = i
    row_i = i + 9

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Тетчер.xlsx", ReadOnly

    msvyaz = Workbooks("Привязка Тетчер.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Тетчер.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For a = 1 To UBound(msvyaz, 1)
        If IsNumeric(msvyaz(a, 2)) = True Then
            msvyaz(a, 2) = CStr(msvyaz(a, 2))
        End If
    Next a

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add msvyaz(a, 1), a
    Next a

    MsgBox "Выберите прайс Тетчер"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
    GoTo ended:
    End If

    Workbooks.Open Filename:=a
    a = ActiveWorkbook.name
    xx = a

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Тетчер.xlsx" Then
            Workbooks("Привязка Тетчер.xlsx").Close
        ElseIf wbk.name = xx Then
            Workbooks(xx).Close
        End If
    Next wbk
    Application.ScreenUpdating = False
    
    
    For row_count = 1 To UBound(mprice, 1)
        mprice(row_count, 1) = Trim(mprice(row_count, 1))
        mprice(row_count, 9) = Trim(mprice(row_count, 9))
    Next row_count

    Set dict1 = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(mprice, 1)

    If Trim(mprice(a, 1)) <> sEmpty Then
        dict1.Add mprice(a, 1), a
    End If

    Next a

    Workbooks(abook).Activate

    Do While m_postavshik(i) = "Тетчер"

        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            
            If dict1.exists(privyazka) Then
                row_privyazka_in_pricelist = dict1.Item(privyazka)
            Else
                row_privyazka_in_pricelist = sEmpty
            End If
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, 9) = vbNullString Or mprice(row_privyazka_in_pricelist, 9) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, 9) = 0
                End If
                    
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, 9))
                
                If Total_1 <> 0 Then
                    Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_1
                End If
            Else
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
            End If
        Else
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
        End If
        i = i + 1
        row_i = row_i + 1
    Loop
    
    Erase msvyaz, mprice
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Тетчер"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Тетчер"
    For i1 = i1 + 9 To row_i - 1 Step 1
        
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If

    Next i1

    uf1.Show vbModeless

ended:

End Sub

Private Sub оДриада(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)


For i = 1 To UBound(m_postavshik)
    
    'Если принудительное обновление - старт
    'Если непринудительное обновление и дата в просчете <> текущей дате - старт
    'В ином случае - выход.
    If m_postavshik(i) = "Дриада-спорт" Then
        If forced = True Then
            Exit For
        ElseIf m_update_date(nomer, 2) <> Date Then
            If MsgBox("Обновить Дриада?", vbYesNo) = vbNo Then Exit Sub
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i

    i1 = i
    row_i = i + 9

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Гиззатулин Антон\0СТАТКИ\Дриада\Остатк Дриада.xlsx", _
    ReadOnly

    msvyaz = Workbooks("Остатк Дриада.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Остатк Дриада.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    For a = 1 To UBound(msvyaz, 1)
        If IsNumeric(msvyaz(a, 1)) = True Then
            msvyaz(a, 1) = CStr(msvyaz(a, 1))
        End If
    Next a

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add CStr(msvyaz(a, 2)), a
    Next a

    Workbooks.Open "https://driada-sport.ru/images/blog/2/driada-price.xlsx", ReadOnly

    Dim Sh As Worksheet
    On Error Resume Next
    For Each Sh In Worksheets
        Sh.UsedRange.ClearComments
        Sh.UsedRange.Hyperlinks.Delete
    Next
    
    xx = ActiveWorkbook.name

    On Error Resume Next
    ActiveWorkbook.ActiveSheet.ShowAllData
    On Error GoTo 0

    mprice = ActiveWorkbook.ActiveSheet.Range("A1:AC" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Остатк Дриада.xlsx" Then
            Workbooks("Остатк Дриада.xlsx").Close
        ElseIf wbk.name = xx Then
            Workbooks(xx).Close False
        End If
    Next wbk
    Application.ScreenUpdating = False

    Set dict1 = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(mprice, 1)
        If dict1.exists(Trim(mprice(a, 1))) = False Then
            If Trim(mprice(a, 1)) <> sEmpty And Trim(mprice(a, 1)) <> 0 Then
                dict1.Add CStr(Trim(mprice(a, 1))), a
            End If
        End If
    Next a

    Workbooks(abook).Activate
          
    Do While m_postavshik(i) = "Дриада-спорт"
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 1)
            
            If dict1.exists(privyazka) Then
                row_privyazka_in_pricelist = dict1.Item(privyazka)
            Else
                row_privyazka_in_pricelist = sEmpty
            End If
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, 29))
                Total_2 = CLng(mprice(row_privyazka_in_pricelist, 28))
                
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = Total_1
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_2
            
            Else
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
            End If
        Else
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
        End If
        i = i + 1
        row_i = row_i + 1
        If i > UBound(m_postavshik) Then Exit Do
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Дриада"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Set uf2 = New UserForm2
    uf2.Caption = "Коды 1С Дриада"

    For i1 = i1 + 9 To row_i - 1 Step 1

        If uf2.TextBox1 = sEmpty Then
            uf2.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf2.TextBox1 = uf2.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1

    uf2.Show vbModeless

End Sub

'''''100 печей''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub о100печей(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c)
'
'
'For i = 1 To Workbooks(abook).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row
'
'If m_postavshik(i) = "100 печей" And _
'Date <> m_update_date(nomer, 2) Or m_postavshik(i) = "100 Печей" And _
'Date <> m_update_date(nomer, 2) Then
'
'If MsgBox("Обновить 100 печей?", vbYesNo) = vbNo Then Exit Sub
'
'i1 = i
'
'Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\100 печей\Остатки\связь.xlsx", ReadOnly
'
'msvyaz = Workbooks("связь.xlsx").ActiveSheet.Range("A1:B" & _
'Workbooks("связь.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'Set dict = CreateObject("Scripting.Dictionary")
'
'For a = 1 To UBound(msvyaz, 1)
'
'dict.Add msvyaz(a, 1), a
'
'Next a
'
'MsgBox "Выберите прайс 100 печей"
'a = Application.GetOpenFilename
'If a = False Then
'MsgBox "Файл не выбран"
'GoTo ended:
'End If
'
'Workbooks.Open Filename:=a
'a = ActiveWorkbook.name
'xx = a
'
'mprice = ActiveWorkbook.ActiveSheet.Range("A1:R" & _
'ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'ncolumn = fNcolumn(mprice, "Номенклатура, Характеристика, Упаковка")
'If ncolumn = "none" Then GoTo ended
'ncolumn1 = fNcolumn(mprice, "ррц")
'If ncolumn1 = "none" Then GoTo ended
'ncolumn2 = fNcolumn(mprice, "опт")
'If ncolumn2 = "none" Then GoTo ended
'ncolumn3 = fNcolumn(mprice, "Склад 100 печей, Екатеринбург")
'If ncolumn3 = "none" Then GoTo ended
'ncolumn4 = fNcolumn(mprice, "Центральный склад")
'If ncolumn4 = "none" Then GoTo ended
'
'Set dict1 = CreateObject("Scripting.Dictionary")
'
'For a = 1 To UBound(mprice, 1)
'
'If Trim(mprice(a, ncolumn)) <> sEmpty And dict1.exists(mprice(a, ncolumn)) = False Then
'dict1.Add mprice(a, ncolumn), a
'End If
'
'Next a
'
'Workbooks.Add
'abook1 = ActiveWorkbook.name
'Workbooks(abook1).ActiveSheet.Cells(1, 1).Value = "Kod 1s"
'Workbooks(abook1).ActiveSheet.Cells(1, 2).Value = "Ostatok"
'
'Workbooks(abook).Activate
'
'Do While StrConv(m_postavshik(i), vbLowerCase) = "100 печей"
'
'If dict.exists(m_kody_1c(i)) = True Then
'If dict1.exists(msvyaz(dict.Item(m_kody_1c(i)), 2)) = True Then
'
''If Trim(mprice(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), 11)) <> sEmpty Or _
''mprice(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), 11) <> 0 Then
''Workbooks(abook).Worksheets("Просчет цен").Cells(i, 19).Value = mprice _
''(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), 11)
''Else
'
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 18).Value = mprice _
'(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), ncolumn1)
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 19).Value = mprice _
'(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), ncolumn2)
'Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
'Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).Value = msvyaz(dict.Item(m_kody_1c(i)), 1)
'Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
'Cells.SpecialCells(xlCellTypeLastCell).Row, 2).Value = CLng(mprice _
'(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), ncolumn3)) + CLng(mprice _
'(dict1.Item(msvyaz(dict.Item(m_kody_1c(i)), 2)), ncolumn4))
'
'Else
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 18).Value = 0
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 19).Value = 0
'Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
'Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).Value = msvyaz(dict.Item(Workbooks(abook). _
'Worksheets("Просчет цен").Cells(i, 17).Value), 1)
'Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
'Cells.SpecialCells(xlCellTypeLastCell).Row, 2).Value = 30000
'
'End If
'
'Else
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 18).Value = 0
'Workbooks(abook).Worksheets("Просчет цен").Cells(i, 19).Value = 0
''Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
''Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).Value = msvyaz(dict.Item(Workbooks(abook). _
''Worksheets("Просчет цен").Cells(i, 17).Value), 1)
''Workbooks(abook1).ActiveSheet.Cells(Workbooks(abook1).ActiveSheet. _
''Cells.SpecialCells(xlCellTypeLastCell).Row, 2).Value = 30000
'End If
'
'i = i + 1
'Loop
'
'Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "100 печей"
'Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
'
'Set uf1 = New UserForm2
'uf1.Caption = "Коды 1С 100 печей"
'
'For i1 = i1 To i - 1 Step 1
''ActiveWorkbook.ActiveSheet.Cells(b, 1).Value = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
'If uf1.TextBox1 = sEmpty Then
'uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
'Else
'uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
'End If
'
''b = b + 1
'
'Next i1
'
'uf1.Show vbModeless
'
''Workbooks.Add
''b = 1
''
''For i1 = i1 To i - 1 Step 1
'''ActiveWorkbook.ActiveSheet.Cells(b, 1).Value = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
''b = b + 1
''Next i1
'
'Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки 100Печей " & Date & ".csv", FileFormat:=xlCSV
'
'End If
'
'Next i
'
'
'For Each wbk In Application.Workbooks
'If wbk.name = "связь.xlsx" Then
'Workbooks("связь.xlsx").Close
'ElseIf wbk.name = xx Then
'Workbooks(xx).Close
'End If
'Next wbk
'
'ended:
'
'End Sub
'''''100 печей''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub оUNIX(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)

For i = 1 To UBound(m_postavshik)
    
    'Если принудительное обновление - старт
    'Если непринудительное, и текущий день месяца нечетный, и не совпадает с днем в просчете - старт
    'В иных случаях - выход
    If m_postavshik(i) = "Unix-line" Then
        If forced = True Then
            Exit For
        ElseIf DatePart("d", Now) Mod 2 <> 0 And m_update_date(nomer, 2) <> Date Then
            If MsgBox("Обновить UNIX?", vbYesNo) = vbNo Then Exit Sub
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i
        
    i1 = i
    row_i = i + 9
    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Unix.xlsx", _
    ReadOnly

    msvyaz = Workbooks("Привязка Unix.xlsx").ActiveSheet.Range("A1:C" & _
    Workbooks("Привязка Unix.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")
    
    'До последнего кода привязки
    'В dict заносится: Ключ - Код 1С, Элемент - номер строки в файле привязки
    For a = 1 To UBound(msvyaz, 1)
        dict.Add msvyaz(a, 1), a
    Next a

    MsgBox "Выберите прайс UNIX-LINE"
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
            Range("A8:E" & ActiveWorkbook.Sheets(ActiveSheetName). _
            Cells.SpecialCells(xlCellTypeLastCell).Row).Value
        
'            For concat_row = 1 To UBound(ConcatArray, 1)
'                If Application.Trim(ConcatArray(concat_row, 1)) = "Москва" Then
'                    ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName).Range("A5:C" & (concat_row + 3)).Value
'                    Exit For
'                End If
'            Next
          
            ReDim mprice(2, 1 To UBound(ConcatArray, 1))
            For j = 1 To UBound(mprice, 2)
                mprice(0, j) = Application.Trim(ConcatArray(j, 1))
                mprice(1, j) = Application.Trim(ConcatArray(j, 5))
                mprice(2, j) = Application.Trim(ConcatArray(j, 4))
            Next
        
            GoTo continue
        End If
    
        ActiveWorkbook.Sheets(ActiveSheetName). _
        Range("A8:E" & ActiveWorkbook.Sheets(ActiveSheetName). _
        Cells.SpecialCells(xlCellTypeLastCell).Row).MergeCells = False
    
        ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName). _
        Range("A8:E" & ActiveWorkbook.Sheets(ActiveSheetName). _
        Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
'        For concat_row = 1 To UBound(ConcatArray, 1)
'            If Application.Trim(ConcatArray(concat_row, 1)) = "Москва" Then
'                ConcatArray = ActiveWorkbook.Sheets(ActiveSheetName).Range("A5:C" & (concat_row + 3)).Value
'                Exit For
'            End If
'        Next
    
        StartPos = UBound(mprice, 2)
        ReDim Preserve mprice(2, 1 To (UBound(ConcatArray, 1) + UBound(mprice, 2)))
        concat_row = 1
        For j = StartPos To UBound(mprice, 2)
            If concat_row > UBound(ConcatArray, 1) Then Exit For
            If Application.Trim(ConcatArray(concat_row, 1)) <> sEmpty Or Application.Trim(ConcatArray(concat_row, 1)) <> vbNullString Then
                mprice(0, j) = Application.Trim(ConcatArray(concat_row, 1))
                mprice(1, j) = Application.Trim(ConcatArray(concat_row, 5))
                mprice(2, j) = Application.Trim(ConcatArray(concat_row, 4))
            End If
            concat_row = concat_row + 1
        Next j
    
continue:
    Next Sh

'    Workbooks.Add
'    xx = ActiveWorkbook.name
'
'    ActiveWorkbook.Queries.Add name:="catalog", Formula:= _
'        "let" & Chr(13) & "" & Chr(10) & "    Source = Xml.Tables(Web.Contents(""https://unixfit.ru/bitrix/catalog_export/catalog.php""))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Changed Type""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Eciaiaiiue oei"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""url"", type tex" & _
'        "t}, {""platform"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Eciaiaiiue oei""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Eciaiaiiue oei1"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""categoryId"", Int64.Type}, {""vendorCode"", type text}, {""name"", type text}, {""Attribute:id"", Int64.Type}, {""Att" & _
'        "ribute:available"", type logical}, {""picture"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Eciaiaiiue oei1"""
'    Sheets.Add After:=ActiveSheet
'    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=catalog" _
'        , Destination:=Range("$A$1")).QueryTable
'        .CommandType = xlCmdSql
'        .CommandText = Array("SELECT * FROM [catalog]")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = False
'        .ListObject.DisplayName = "catalog"
'        .Refresh BackgroundQuery:=False
'    End With
'   Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

'ActiveWorkbook.Queries.Add name:="catalog", Formula:= _
'    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://unixfit.ru/bitrix/catalog_export/catalog.php""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop"" = Table.ExpandTableColumn(#""Измененный тип"", ""shop"", {""name"", ""company"", ""url"", ""platform"", ""currencies"", ""cate" & _
'    "gories"", ""offers""}, {""shop.name"", ""shop.company"", ""shop.url"", ""shop.platform"", ""shop.currencies"", ""shop.categories"", ""shop.offers""})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers"" = Table.ExpandTableColumn(#""Развернутый элемент shop"", ""shop.offers"", {""offer""}, {""shop.offers.offer""})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers.offer"" = Table.Expand" & _
'    "TableColumn(#""Развернутый элемент shop.offers"", ""shop.offers.offer"", {""url"", ""price"", ""currencyId"", ""categoryId"", ""vendorCode"", ""name"", ""description"", ""param"", ""Attribute:id"", ""Attribute:available"", ""picture""}, {""shop.offers.offer.url"", ""shop.offers.offer.price"", ""shop.offers.offer.currencyId"", ""shop.offers.offer.categoryId"", ""shop" & _
'    ".offers.offer.vendorCode"", ""shop.offers.offer.name"", ""shop.offers.offer.description"", ""shop.offers.offer.param"", ""shop.offers.offer.Attribute:id"", ""shop.offers.offer.Attribute:available"", ""shop.offers.offer.picture""})," & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers.offer.vendorCode"" = Table.ExpandTableColumn(#""Развернутый элемент shop.offers.offer"", ""shop." & _
'    "offers.offer.vendorCode"", {""Element:Text""}, {""shop.offers.offer.vendorCode.Element:Text""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Развернутый элемент shop.offers.offer.vendorCode"""
'
''    Sheets.Add After:=ActiveSheet
'
'With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
'    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=catalog" _
'    , Destination:=Range("$A$1")).QueryTable
'    .CommandType = xlCmdSql
'    .CommandText = Array("SELECT * FROM [catalog]")
'    .RowNumbers = False
'    .FillAdjacentFormulas = False
'    .PreserveFormatting = True
'    .RefreshOnFileOpen = False
'    .BackgroundQuery = True
'    .RefreshStyle = xlInsertDeleteCells
'    .SavePassword = False
'    .SaveData = True
'    .AdjustColumnWidth = True
'    .RefreshPeriod = 0
'    .PreserveColumnInfo = False
'    .ListObject.DisplayName = "catalog"
'    .Refresh BackgroundQuery:=False
'End With
'
'    mprice = ActiveWorkbook.ActiveSheet.Range("A1:P" & _
'    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Unix.xlsx" Then
            Workbooks("Привязка Unix.xlsx").Close False
        ElseIf wbk.name = xx Then
            Workbooks(xx).Close False
        End If
    Next wbk
    Application.ScreenUpdating = False
    
    ncolumn = 0
    ncolumn1 = 1
    ncolumn2 = 2
    
'    ncolumn = fNcolumn(mprice, "vendorCode")
'    If ncolumn = "none" Then GoTo ended
'    ncolumn1 = fNcolumn(mprice, "price")
'    If ncolumn1 = "none" Then GoTo ended
'    ncolumn2 = fNcolumn(mprice, "Attribute:available")
'    If ncolumn2 = "none" Then GoTo ended

    Set dict1 = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(mprice, 2)
        If dict1.exists(Trim(mprice(ncolumn, a))) = False Then
            If Trim(mprice(ncolumn, a)) <> "" And Trim(mprice(ncolumn, a)) <> 0 Then
                dict1.Add Trim(mprice(ncolumn, a)), a
            End If
        End If
    Next a

'    Workbooks.Add
'    abook1 = ActiveWorkbook.name
'    Workbooks(abook1).ActiveSheet.Cells(1, 1).Value = "Kod 1s"
'    Workbooks(abook1).ActiveSheet.Cells(1, 2).Value = "Ostatok"

    Workbooks(abook).Activate

    Do While m_postavshik(i) = "Unix-line"
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            discount = msvyaz(row_svyaz, 3)
            
            If dict1.exists(privyazka) Then
                row_privyazka_in_pricelist = dict1.Item(privyazka)
            Else
                row_privyazka_in_pricelist = sEmpty
            End If
            
            If row_privyazka_in_pricelist <> sEmpty Then
                Total_1 = CLng(mprice(ncolumn1, row_privyazka_in_pricelist))
                Total_2 = CLng(mprice(ncolumn2, row_privyazka_in_pricelist))
                                
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = Total_1
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_2
            Else
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
            End If
        Else
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
        End If
        i = i + 1
        row_i = row_i + 1
    Loop
    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "UNIX"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If

    Set uf2 = New UserForm2
    uf2.Caption = "Коды 1С Unix-line"

    For i1 = i1 + 9 To row_i - 1 Step 1

        If uf2.TextBox1 = sEmpty Then
            uf2.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf2.TextBox1 = uf2.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1

'    For i2 = 2 To UBound(msvyaz, 1)
'        If dict1.exists(msvyaz(i2, 2)) = True Then
'            Workbooks(abook1).ActiveSheet.Cells(i2, 1).Value = msvyaz(i2, 1)
'            If mprice(dict1.Item(msvyaz(i2, 2)), ncolumn2) = "true" Or mprice(dict1.Item(msvyaz(i2, 2)), ncolumn2) = True Then
'                Workbooks(abook1).ActiveSheet.Cells(i2, 2).Value = 30001
'            Else
'                Workbooks(abook1).ActiveSheet.Cells(i2, 2).Value = 30000
'            End If
'        Else
'            Workbooks(abook1).ActiveSheet.Cells(i2, 1).Value = msvyaz(i2, 1)
'            Workbooks(abook1).ActiveSheet.Cells(i2, 2).Value = 30000
'        End If
'    Next i2
'
'    Workbooks(abook1).SaveAs Filename:=Environ("OSTATKI") & "Остатки UNIX " & Date & ".csv", FileFormat:=xlCSV

    uf2.Show vbModeless

ended:

End Sub

Private Sub оSwollen(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)

For i = 1 To UBound(m_postavshik)

    'Если принудительное обновление - старт
    'Если непринудительное обновление, и текущая дата нечетная, и текущая дата не совпадает с датой просчета - старт
    'В иных случаях - выход
    If m_postavshik(i) = "Swollen" Then
        If forced = True Then
            Exit For
        ElseIf Date <> m_update_date(nomer, 2) And DatePart("d", Now) Mod 2 <> 0 Then
            If MsgBox("Обновить Swollen?", vbYesNo) = vbNo Then Exit Sub
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i
    
    i1 = i
    row_i = i + 9

    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Swollen.xlsx", ReadOnly

    msvyaz = Workbooks("Привязка Swollen.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Swollen.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    'Ключ - Код 1С, элемент - номер строки в привязке
    Set dict = CreateObject("Scripting.Dictionary")
    For a = 1 To UBound(msvyaz, 1)
        dict.Add msvyaz(a, 1), a
    Next a

    MsgBox "Выберите прайс SWOLLEN"
    SwollenFilename = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файлы не выбраны"
        GoTo ended
    End If
    Workbooks.Open Filename:=SwollenFilename
    SwollenFilename = ActiveWorkbook.name
    'xx1 = a

    'РРЦ - G(7)
    'ОПТ - H?(8)
    'Наличие - F(6)
    MpriceSwollen = ActiveWorkbook.ActiveSheet.Range("A1:K" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Swollen.xlsx" Then
            Workbooks("Привязка Swollen.xlsx").Close
        ElseIf wbk.name = SwollenFilename Then
            Workbooks(SwollenFilename).Close
        End If
    Next wbk
    Application.ScreenUpdating = True

    'Ключ - Артикул в прайсе, Элемент - номер строки в прайсе
    Set dict1 = CreateObject("Scripting.Dictionary")
    For a = 1 To UBound(MpriceSwollen, 1)
        If Trim(MpriceSwollen(a, 1)) <> sEmpty And dict1.exists(MpriceSwollen(a, 1)) = False Then
            dict1.Add MpriceSwollen(a, 1), a
        End If
    Next a

another1:
    Workbooks.Add
    OstatkiFilename = ActiveWorkbook.name
    Workbooks(OstatkiFilename).ActiveSheet.Cells(1, 1).Value = "Kod 1s"
    Workbooks(OstatkiFilename).ActiveSheet.Cells(1, 2).Value = "Ostatok"
    Workbooks(abook).Activate
        
        Do While StrConv(m_postavshik(i), vbLowerCase) = "swollen"
            If dict.exists(m_kody_1c(i)) = True Then
                row_svyaz = dict.Item(m_kody_1c(i))
                privyazka = msvyaz(row_svyaz, 2)
                row_privyazka_in_pricelist = dict1.Item(privyazka)

                If row_privyazka_in_pricelist <> sEmpty Then
                    
                    Total_1 = CLng(MpriceSwollen(row_privyazka_in_pricelist, 7))
                    Total_2 = CLng(MpriceSwollen(row_privyazka_in_pricelist, 8))
                     
                    If Total_1 <> 0 And Total_2 = 0 Then
                        Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                        Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_1
                    Else
                        Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = Total_1
                        Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_2
                    End If
                        
                Else
                    Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                    Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
                End If

            Else
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
            End If
            i = i + 1
            row_i = row_i + 1
            If i > UBound(m_postavshik) Then Exit Do
        Loop
    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Swollen"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Swollen"

    For i1 = i1 + 9 To row_i - 1 Step 1

        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If

    Next i1

    For i2 = 2 To UBound(msvyaz, 1)
        kod_1c = msvyaz(i2, 1)
        If dict1.exists(msvyaz(i2, 2)) = True Then
            privyazka = msvyaz(i2, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            If row_privyazka_in_pricelist <> sEmpty Then
                If MpriceSwollen(row_privyazka_in_pricelist, 6) = 0 _
                Or MpriceSwollen(row_privyazka_in_pricelist, 6) = sEmpty Then
                    Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 1).Value = kod_1c
                    Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 2).Value = 30000
                Else
                    Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 1).Value = kod_1c
                    Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 2).Value = MpriceSwollen(row_privyazka_in_pricelist, 6)
                End If
            Else
                Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 1).Value = kod_1c
                Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 2).Value = 30000
            End If
        Else
            Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 1).Value = kod_1c
            Workbooks(OstatkiFilename).ActiveSheet.Cells(i2, 2).Value = 30000
        End If
    Next i2
    
    Workbooks(OstatkiFilename).Activate
    Workbooks(OstatkiFilename).ActiveSheet.Range("B2:B" & UBound(msvyaz, 1)).Select
    Selection.Replace What:=" шт.", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    uf1.Show vbModeless

    Workbooks(OstatkiFilename).SaveAs Filename:=Environ("OSTATKI") & "Остатки Swollen " & Date & ".csv", FileFormat:=xlCSV

ended:
End Sub

Private Sub oАвтовентури(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)

For i = 1 To UBound(m_postavshik)
    
    'Если принудительное обновление - старт
    'Если непринудительное обновление и текущая дата отличается от даты в листе "ОДаты"
    'И день недели - понедельник или четверг - старт
    'В иных случаях - выход
    
    If m_postavshik(i) = "Автовентури" Then
        If forced = True Then
            Exit For
        ElseIf Date <> m_update_date(nomer, 2) And _
        DatePart("w", Now, vbMonday) = 1 Or _
        Date <> m_update_date(nomer, 2) And _
        DatePart("w", Now, vbMonday) = 4 Then
            If MsgBox("Обновить Автовентури?", vbYesNo) = vbNo Then Exit Sub
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i

    Set dict1 = CreateObject("Scripting.Dictionary")
    
    i1 = i
    row_i = i + 9
    
    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Автовентури.xlsx", _
    ReadOnly
    
    msvyaz = Workbooks("Привязка Автовентури.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Автовентури.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add CStr(msvyaz(a, 1)), a
    Next a

    
    Application.Workbooks.Open "http://www.autoventuri.ru/upload/prices/price.xlsx"
    
    price_name = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("Y1:AK" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Автовентури.xlsx" Then
            Workbooks("Привязка Автовентури.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, 1)) = True And mprice(j, 1) <> sEmpty Then
            mprice(j, 1) = CStr(mprice(j, 1))
        End If
    Next j

    For a = 1 To UBound(mprice, 1)
        If Trim(mprice(a, 1)) <> sEmpty And dict1.exists(mprice(a, 1)) = False Then
            dict1.Add mprice(a, 1), a
        End If
    Next a
            
    Do While StrConv(m_postavshik(i), vbLowerCase) = "автовентури"
    
        If dict.exists(m_kody_1c(i)) = True Then
    
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
    
            If row_privyazka_in_pricelist <> sEmpty Then
    
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, 12))
                Total_2 = CLng(mprice(row_privyazka_in_pricelist, 13))
    
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = Total_1
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = Total_2
            Else
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
                Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
            End If
        Else
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 18).Value = 0
            Workbooks(abook).Worksheets("Просчет цен").Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        If i > UBound(m_postavshik) Then Exit Do
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Автовентури"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Автовентури"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless

ended:
End Sub

Private Sub oDekomo(ByRef nomer, ByRef m_update_date, forced)

Dim m_artikul() As String
Dim m_rrc() As Long
Dim m_opt() As Long
Dim fprice() As Long
Dim userform_kody_1C() As String
Dim kod_userform As String
Dim ncolumn_artikul As Long
Dim ncolumn_rrc As Long
Dim ncolumn_opt As Long
Dim i As Long
Dim i1 As Long
Dim row_privyazka_in_price As Long
Dim kody_1c_combine As String

'Название файла основного просчета
main_proschet_name = ActiveWorkbook.name

'Если принудительное обновление - старт
'Если непринудительное обновление дата обновления не совпадает с текущей датой, и
'сегодня нечетный день - старт
'В иных случаях - выход
If forced = True Then
ElseIf Date = m_update_date(nomer, 2) _
Or Date <> m_update_date(nomer, 2) And DatePart("d", Now) Mod 2 = 0 Then
    Exit Sub
ElseIf MsgBox("Обновить Dekomo?", vbYesNo) = vbNo Then
    Exit Sub
Else
    Exit Sub
End If

Application.Workbooks.Open "\\AKABINET\Doc\СпортОтдыхТуризм\Для Николая\SOT резерв\актуальный резерв\Просчет цен SOT(Люстры 938).xlsb"

dekomo_proschet_name = ActiveWorkbook.name

'Добавляется наименование поставщика (Dekomo) в массив
m_postavshik = Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Range("A11:A" & _
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row - 1).Value

For i = 1 To UBound(m_postavshik)
    m_postavshik(i, 1) = Trim(m_postavshik(i, 1))
Next i

PrivyazkaArray = Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Range("C11:C" & _
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row - 1).Value

m_kody_1c = Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Range("Q11:Q" & _
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row - 1).Value

For i = 1 To UBound(m_kody_1c)
    m_kody_1c(i, 1) = Trim(m_kody_1c(i, 1))
    PrivyazkaArray(i, 1) = Trim(PrivyazkaArray(i, 1))
Next i


'Скачивается файл прогрузки, и добавляется в массив mprice
Application.Workbooks.Open "http://www.dekomo.ru/get_content/v4/excel.php?ost=1&brand_type[]=30&token=94fbe3099b8d207301b891b9ea7dc667"
dekomo_progruzka_file = ActiveWorkbook.name

mprice = ActiveWorkbook.ActiveSheet.Range("A1:EH" & _
Cells(Rows.Count, 2).End(xlUp).Row).Value

'Поиск номеров столбца с наименованиями
ncolumn_artikul = fNcolumn(mprice, "Артикул")
If ncolumn = "none" Then GoTo ended

ncolumn_rrc = fNcolumn(mprice, "МРЦ/РРЦ")
If ncolumn1 = "none" Then GoTo ended

ncolumn_opt = fNcolumn(mprice, "Закупочная цена")
If ncolumn1 = "none" Then GoTo ended

'В словарь добавляются: Ключ - артикул, Элемент - номер строки в файле выгрузки
Set dict = CreateObject("Scripting.Dictionary")
For i1 = 1 To Cells(Rows.Count, 2).End(xlUp).Row
dict.Add mprice(i1, ncolumn_artikul), i1
Next i1

Workbooks(dekomo_progruzka_file).Close False

ReDim m_rrc(UBound(mprice))
ReDim m_opt(UBound(mprice))

'mprice разделяется на столбец с РРЦ и оптом, и очищается
i1 = 1
Do While i1 < UBound(mprice)
    m_rrc(i1) = mprice(i1 + 1, ncolumn_rrc)
    m_opt(i1) = mprice(i1 + 1, ncolumn_opt)
    i1 = i1 + 1
Loop

Erase mprice

ReDim fprice(UBound(PrivyazkaArray) - 1, 1)
ReDim userform_kody_1C(UBound(PrivyazkaArray) - 1, 1)

i = 1

'Вносятся старые цены из просчета люстр
old_fprice = Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Range("R11:S" & _
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row - 1).Value

'Сравниваются коды 1С из просчета и артикулы из файла выгрузки,
'При соответствии в fprice вносится РРЦ и ОПТ из файла выгрузки,
Do While m_postavshik(i, 1) = "Dekomo"
    
    If dict.exists(PrivyazkaArray(i, 1)) = True Then
                
        row_privyazka_in_pricelist = dict.Item(PrivyazkaArray(i, 1))
        
        If row_privyazka_in_pricelist <> sEmpty Then
            
            fprice(i - 1, 0) = m_rrc(row_privyazka_in_pricelist - 1)
            fprice(i - 1, 1) = m_opt(row_privyazka_in_pricelist - 1)
        Else
            fprice(i - 1, 0) = 0
            fprice(i - 1, 1) = 0
        End If
    Else
        fprice(i - 1, 0) = 0
        fprice(i - 1, 1) = 0
    End If
    
    i = i + 1
    If i > UBound(m_postavshik, 1) Then Exit Do
Loop

'Сравнивается старая цена из просчета и новая цена из массива fprice
'для формирования списка кодов 1С с изменениями в цене
'userform_kody_1c - массив кодов 1С по которым есть изменения в цене
i1 = 0
For i = 1 To UBound(fprice)
If old_fprice(i, 1) <> fprice(i - 1, 0) Or old_fprice(i, 2) <> fprice(i - 1, 1) Then
    userform_kody_1C(i1, 0) = m_kody_1c(i, 1)
    i1 = i1 + 1
End If
Next i
  
'Вносится новая цена в просчет люстр
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Range("R11:S" & _
Workbooks(dekomo_proschet_name).Worksheets("Просчет цен").Cells.SpecialCells(xlCellTypeLastCell).Row - 1).Value = fprice

'Массив кодов 1С с изменениями объединяется в одну переменную
i1 = 0
Do While i1 < UBound(userform_kody_1C)
    If userform_kody_1C(i1, 0) <> sEmpty Then
    kody_1c_combine = kody_1c_combine & vbCrLf & userform_kody_1C(i1, 0)
    End If
    i1 = i1 + 1
Loop
       
'Внесение в юзерформу кодов 1С с изменениями и её показ
Set uf1 = New UserForm2
uf1.Caption = "Коды 1С c изменениями Dekomo"
uf1.TextBox1 = kody_1c_combine
uf1.Show vbModeless


If forced = False Then
    Workbooks(main_proschet_name).Worksheets("ДатыО").Cells(nomer, 1).Value = "Dekomo"
    Workbooks(main_proschet_name).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
End If
    
ended:
End Sub

Private Sub oАфина(nomer, ByRef m_postavshik, ByRef m_update_date, _
ByRef m_kody_1c, forced)

Dim i As Long
Dim i1 As Long
Dim row_i As Long
Dim price_name As String
Dim wbk As Object
Dim row_svyaz As Long
Dim row_privyazka_in_pricelist As Long
Dim Total_1 As Long
Dim Total_2 As Long
Dim privyazka As String
        
    For i = 1 To UBound(m_postavshik)
        
        'Если принудительное обновление - старт
        'Если непринудительное обновление и текущая дата отличается от даты в листе "ОДаты"
        'И день недели - четверг - старт
        'В иных случаях - выход
        
        If m_postavshik(i) = "Афина" Then
            If forced = True Then
                Exit For
            ElseIf DatePart("w", Now, vbMonday) = 4 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Афина?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    i1 = i
    row_i = i + 9
    
    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Афина.xlsx", ReadOnly
    
    msvyaz = Workbooks("Привязка Афина.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Афина.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    Workbooks.OpenText Filename:="http://afinalux.ru/clients/import.csv", DataType:=xlDelimited, local:=True
    
    mprice = ActiveWorkbook.ActiveSheet.Range("D1:N" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    price_name = ActiveWorkbook.name
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Афина.xlsx" Then
            Workbooks("Привязка Афина.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "_SKU_", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "_PRICE_", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim ncolumn_opt As Integer
    ncolumn_opt = fNcolumn(mprice, "_WHOLESALE_1M_", 1)
    If ncolumn_opt = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = CStr(mprice(j, ncolumn_artikul))
        End If
        
        mprice(j, ncolumn_opt) = Replace(mprice(j, ncolumn_opt), ".", ",")
        mprice(j, ncolumn_rrc) = Replace(mprice(j, ncolumn_rrc), ".", ",")
        
        
        If Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
       
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "Афина"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                Total_1 = mprice(row_privyazka_in_pricelist, ncolumn_rrc)
                Total_2 = mprice(row_privyazka_in_pricelist, ncolumn_opt)

                proschet.Cells(row_i, 18).Value = Total_1
                proschet.Cells(row_i, 19).Value = Total_2
            Else

                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Афина"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Афина"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
    
ended:
End Sub

Private Sub oWeekend(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)


Dim i As Long
Dim i1 As Long
Dim row_i As Long
Dim price_name As String
Dim wbk As Object
Dim row_svyaz As Long
Dim row_privyazka_in_pricelist As Long
Dim Total_1 As Long
Dim Total_2 As Long
Dim privyazka As String
        
    For i = 1 To UBound(m_postavshik)
        
        'Если принудительное обновление - старт
        'Если непринудительное обновление и текущая дата отличается от даты в листе "ОДаты"
        'И день недели - четверг - старт
        'В иных случаях - выход
        
        If m_postavshik(i) = "Weekend" Then
            If forced = True Then
                Exit For
            ElseIf DatePart("w", Now, vbMonday) = 4 And m_update_date(nomer, 2) <> Date Or _
            DatePart("w", Now, vbMonday) = 1 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Weekend?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    i1 = i
    row_i = i + 9
    
    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Weekend.xlsx", ReadOnly
    
    msvyaz = Workbooks("Привязка Weekend.xlsx").ActiveSheet.Range("A1:B" & _
    Workbooks("Привязка Weekend.xlsx").ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    Application.Workbooks.Open "https://opt.weekend-billiard.ru/upload/catalog_WB.xlsx"
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:FB" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    price_name = ActiveWorkbook.name
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Weekend.xlsx" Then
            Workbooks("Привязка Weekend.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "Артикул [XLS_ARTIKUL] {IP_PROP312}", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "Цена " & Chr(34) & "Розничная цена" & Chr(34) & " {ICAT_PRICE1_PRICE}", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim ncolumn_opt As Integer
    ncolumn_opt = fNcolumn(mprice, "Цена " & Chr(34) & "Оптовая цена" & Chr(34) & " {ICAT_PRICE2_PRICE}", 1)
    If ncolumn_opt = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = CStr(mprice(j, ncolumn_artikul))
        End If
        
        mprice(j, ncolumn_opt) = Replace(mprice(j, ncolumn_opt), ".", ",")
        mprice(j, ncolumn_rrc) = Replace(mprice(j, ncolumn_rrc), ".", ",")
        
        
        If Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    'i = 5259
    Do While m_postavshik(i) = "Weekend"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                
                If mprice(row_privyazka_in_pricelist, ncolumn_opt) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_opt) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_opt) = 0
                End If

                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_opt))
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Weekend"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Weekend"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub


Private Sub oVictoryFit(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)
    For i = 1 To UBound(m_postavshik)
        
        'Если принудительное обновление - старт
        'Если непринудительное обновление и текущая дата отличается от даты в листе "ОДаты"
        'И день недели - пятница - старт
        'В иных случаях - выход
        
        If m_postavshik(i) = "Victory Fit" Then
            If forced = True Then Exit For
            '''''''''''''''''''''''''''''Для теста - вторник
            If DatePart("w", Now, vbMonday) = 2 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить VictoryFit?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    i1 = i
    row_i = i + 9
    
    Application.Workbooks.Open "\\Akabinet\doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Victory Fit.xlsx", ReadOnly
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    Workbooks.Add
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
    price_name = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Victory Fit привязка.xlsx" Then
            Workbooks("Victory Fit привязка.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "vendorCode", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "price", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "Victory Fit"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If

                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = WorksheetFunction.RoundUp(Total_1 * 0.7, 0)
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Victory Fit"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Victory Fit"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub

Private Sub oWoodville(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)
    For i = 1 To UBound(m_postavshik)
        
        'Если принудительное обновление - старт
        'Если непринудительное обновление и текущая дата отличается от даты в листе "ОДаты"
        'И день недели - пятница - старт
        'В иных случаях - выход
        
        If m_postavshik(i) = "Woodville" Or m_postavshik(i) = "WoodVILLE" Then
            If forced = True Then Exit For
            '''''''''''''''''''''''''''''Для теста - вторник
            If DatePart("w", Now, vbMonday) = 2 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Woodville?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    i1 = i
    row_i = i + 9
    
    Application.Workbooks.Open ("\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Woodville.xlsx")
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        msvyaz(a, 2) = CStr(msvyaz(a, 2))
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    MsgBox "Выберите прайс WoodVille"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.OpenText Filename:=a, DataType:=xlDelimited, local:=True
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:M" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Woodville.xlsx" Then
            Workbooks("Привязка Woodville.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "Артикул", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "МРЦ", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim ncolumn_opt As Integer
    ncolumn_opt = fNcolumn(mprice, "до 150 т.р.", 1)
    If ncolumn_opt = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "Woodville"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                
                If mprice(row_privyazka_in_pricelist, ncolumn_opt) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_opt) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_opt) = 0
                End If

                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_opt))
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
                    
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Woodville"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Woodville"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub

Private Sub oПромет(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)
    For i = 1 To UBound(m_postavshik)
        If m_postavshik(i) = "Промет" Then
            If forced Then
                Exit For
            ElseIf DatePart("w", Now, vbMonday) = 1 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Промет?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
    
    i1 = i
    row_i = i + 9
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    Application.Workbooks.Open ("\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Промет.xlsx")
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        msvyaz(a, 2) = CStr(msvyaz(a, 2))
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    MsgBox "Выберите прайс Промет"
    a = Application.GetOpenFilename
    If a = False Then
        MsgBox "Файл не выбран"
        GoTo ended:
    End If
    Workbooks.OpenText Filename:=a, DataType:=xlDelimited, local:=True
    
    ActiveWorkbook.Sheets("Главная").Activate
    
    If ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.ExclusiveAccess
        Application.DisplayAlerts = True
    End If
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:E" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    price_name = ActiveWorkbook.name
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Промет.xlsx" Then
            Workbooks("Привязка Промет.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "артикул", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "Московская розничная цена", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "Промет"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = Round(Total_1 * 0.73)
                
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Промет"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Промет"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub

'Private Sub oUltragym(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)
'
'    For i = 1 To UBound(m_postavshik)
'        If m_postavshik(i) = "Ultragym" Then
'            If forced Then
'                Exit For
'            ElseIf DatePart("w", Now, vbMonday) = 1 And m_update_date(nomer, 2) <> Date Then
'                If MsgBox("Обновить Ultragym?", vbYesNo) = vbNo Then Exit Sub
'                Exit For
'            Else
'                Exit Sub
'            End If
'        End If
'    Next i
'
'    i1 = i
'    row_i = i + 9
'
'    Set dict1 = CreateObject("Scripting.Dictionary")
'
'    Application.Workbooks.Open ("\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Ultragym.xlsx")
'
'    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
'    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'    Set dict = CreateObject("Scripting.Dictionary")
'
'    For a = 1 To UBound(msvyaz, 1)
'        msvyaz(a, 2) = CStr(msvyaz(a, 2))
'        dict.Add CStr(msvyaz(a, 1)), a
'    Next a
'
'    Workbooks.Add
'
'    ActiveWorkbook.Queries.Add name:="market", Formula:= _
'    "let" & Chr(13) & "" & Chr(10) & "    Источник = Xml.Tables(Web.Contents(""https://sportholding.ru/index.php?route=extension/payment/yandex_money/market""))," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(Источник,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Измененный тип""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company" & _
'    """, type text}, {""url"", type text}, {""platform"", type text}, {""version"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип1""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"" = Table.TransformColumnTypes(offer,{{""categoryId"", Int64.Type}, {""url"", type text}, {""price"", Int64.Type}, {""oldprice"", Int64.Type}, {""currencyId"", type text}, {" & _
'    """description"", type text}, {""vendor"", type text}, {""model"", type text}, {""weight"", Int64.Type}, {""dimensions"", type text}, {""market-sku"", type text}, {""available"", Int64.Type}, {""Attribute:id"", Int64.Type}, {""Attribute:type"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип2"""
'    Sheets.Add After:=ActiveSheet
'    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=market" _
'        , Destination:=Range("$A$1")).QueryTable
'        .CommandType = xlCmdSql
'        .CommandText = Array("SELECT * FROM [market]")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = False
'        .ListObject.DisplayName = "market"
'        .Refresh BackgroundQuery:=False
'    End With
'    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
'
'    price_name = ActiveWorkbook.name
'
'    mprice = ActiveWorkbook.ActiveSheet.Range("A1:O" & _
'    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
'
'    Application.ScreenUpdating = False
'    For Each wbk In Application.Workbooks
'        If wbk.name = "Привязка Ultragym.xlsx" Then
'            Workbooks("Привязка Ultragym.xlsx").Close False
'        ElseIf wbk.name = price_name Then
'            Workbooks(price_name).Close False
'        End If
'    Next wbk
'    Application.ScreenUpdating = True
'
'    'Поиск номеров столбца с наименованиями
'    Dim ncolumn_artikul As Integer
'    ncolumn_artikul = fNcolumn(mprice, "Attribute:id", 1)
'    If ncolumn_artikul = 0 Then GoTo ended
'
'    Dim ncolumn_rrc As Integer
'    ncolumn_rrc = fNcolumn(mprice, "price", 1)
'    If ncolumn_rrc = 0 Then GoTo ended
'
'    Dim j As Long
'    For j = 1 To UBound(mprice)
'        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
'            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
'        End If
'
'        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
'            dict1.Add mprice(j, ncolumn_artikul), j
'        End If
'    Next j
'
'    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
'
'    Do While m_postavshik(i) = "Ultragym"
'
'        If dict.exists(m_kody_1c(i)) = True Then
'
'            row_svyaz = dict.Item(m_kody_1c(i))
'            privyazka = msvyaz(row_svyaz, 2)
'            discount = msvyaz(row_svyaz, 3)
'            row_privyazka_in_pricelist = dict1.Item(privyazka)
'
'            If row_privyazka_in_pricelist <> sEmpty Then
'
'                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
'                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
'                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
'                End If
'
'                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
'                If discount = "домашка" Then
'                    discount = 0.7
'                ElseIf discount = "кардио" Then
'                    If Total_1 < 500000 Then
'                        discount = 0.7
'                    ElseIf Total_1 >= 500000 Then
'                        discount = 0.65
'                    End If
'                Else
'                    discount = 1
'                End If
'                Total_2 = Round(Total_1 * discount)
'
'                If Total_2 = 0 Or Total_2 = sEmpty Then
'                    proschet.Cells(row_i, 18).Value = 0
'                    proschet.Cells(row_i, 19).Value = 0
'                Else
'                    proschet.Cells(row_i, 18).Value = Total_1
'                    proschet.Cells(row_i, 19).Value = Total_2
'                End If
'
'            Else
'                proschet.Cells(row_i, 18).Value = 0
'                proschet.Cells(row_i, 19).Value = 0
'            End If
'
'        Else
'            proschet.Cells(row_i, 18).Value = 0
'            proschet.Cells(row_i, 19).Value = 0
'        End If
'
'        i = i + 1
'        row_i = row_i + 1
'
'        If i > UBound(m_postavshik) Then Exit Do
'
'    Loop
'
'    Erase mprice, msvyaz
'    Set dict = Nothing
'    Set dict1 = Nothing
'    If forced = False Then
'        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Ultragym"
'        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
'    End If
'
'    Workbooks(abook).Activate
'
'    Set uf1 = New UserForm2
'    uf1.Caption = "Коды 1С Ultragym"
'
'    For i1 = i1 + 9 To row_i - 1 Step 1
'        If uf1.TextBox1 = sEmpty Then
'            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
'        Else
'            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
'        End If
'    Next i1
'
'    uf1.Show vbModeless
'
'ended:
'End Sub

Private Sub oLuxfire(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)


For i = 1 To UBound(m_postavshik)
        If m_postavshik(i) = "Luxfire" Then
            If forced Then
                Exit For
            ElseIf DatePart("w", Now, vbMonday) = 1 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Luxfire?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
        
    i1 = i
    row_i = i + 9
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    Application.Workbooks.Open ("\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка Luxfire.xlsx")
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:B" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        msvyaz(a, 2) = CStr(msvyaz(a, 2))
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    Workbooks.Add
    
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
        
    price_name = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Luxfire.xlsx" Then
            Workbooks("Привязка Luxfire.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "model", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "price", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "Luxfire"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = Round(Total_1 * 0.8)
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Luxfire"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Luxfire"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub

Private Sub oAvalon(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)


For i = 1 To UBound(m_postavshik)
        If m_postavshik(i) = "AVALON" Then
            If forced Then
                Exit For
            ElseIf DatePart("d", Now) = 1 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить AVALON?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i
        
    i1 = i
    row_i = i + 9
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    Application.Workbooks.Open ("\\AKABINET\Doc\Наполнение сайтов\Наташа\ОСТАТКИ Сот\Привязка\Привязка AVALON.xlsx")
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        msvyaz(a, 2) = CStr(msvyaz(a, 2))
        dict.Add CStr(msvyaz(a, 1)), a
    Next a
    
    Workbooks.Add
    
    IndentStr = Chr(13) & "" & Chr(10)
    
    ActiveWorkbook.Queries.Add name:="record", Formula:= _
        "let" _
            & IndentStr & "Источник = Xml.Tables(Web.Contents(" & Chr(34) & "https://www.avalonmebel.ru/upload/XML_inv_bal/EXPORT.XML" & Chr(34) & "), null, 1251)," _
            & IndentStr & "Table0 = Источник{0}[Table]," _
            & IndentStr & "#" & Chr(34) & "Измененный тип" & Chr(34) & " = Table.TransformColumnTypes(Table0,{{" & Chr(34) & "leader" & Chr(34) & ", type text}, {" & Chr(34) & "Attribute:format" & Chr(34) & ", type text}})," _
            & IndentStr & "#" & Chr(34) & "Добавлен индекс" & Chr(34) & " = Table.AddIndexColumn(#" & Chr(34) & "Измененный тип" & Chr(34) & ", " & Chr(34) & "Индекс" & Chr(34) & ", 0, 1)," _
            & IndentStr & "#" & Chr(34) & "Развернутый элемент field" & Chr(34) & " = Table.ExpandTableColumn(#" & Chr(34) & "Добавлен индекс" & Chr(34) & ", " & Chr(34) & "field" & Chr(34) & ", {" & Chr(34) & "Element:Text" & Chr(34) & ", " & Chr(34) & "Attribute:tag" & Chr(34) & "}, {" & Chr(34) & "field.Element:Text" & Chr(34) & ", " & Chr(34) & "field.Attribute:tag" & Chr(34) & "})," _
            & IndentStr & "#" & Chr(34) & "Удаленные столбцы" & Chr(34) & " = Table.RemoveColumns(#" & Chr(34) & "Развернутый элемент field" & Chr(34) & ",{" & Chr(34) & "leader" & Chr(34) & ", " & Chr(34) & "Attribute:format" & Chr(34) & "})," _
            & IndentStr & "#" & Chr(34) & "Измененный тип1" & Chr(34) & " = Table.TransformColumnTypes(#" & Chr(34) & "Удаленные столбцы" & Chr(34) & ",{{" & Chr(34) & "field.Element:Text" & Chr(34) & ", type text}})," _
            & IndentStr & "#" & Chr(34) & "Сведенный столбец" & Chr(34) & " = Table.Pivot(#" & Chr(34) & "Измененный тип1" & Chr(34) & ", List.Distinct(#" & Chr(34) & "Измененный тип1" & Chr(34) & "[#" & Chr(34) & "field.Attribute:tag" & Chr(34) & "]), " & Chr(34) & "field.Attribute:tag" & Chr(34) & ", " & Chr(34) & "field.Element:Text" & Chr(34) & ")" _
            & IndentStr & _
        "in" _
          & IndentStr & "#" & Chr(34) & "Сведенный столбец" & Chr(34)
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=record" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [record]")
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
        .ListObject.DisplayName = "record"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    price_name = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:U" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "Привязка Luxfire.xlsx" Then
            Workbooks("Привязка Luxfire.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "1", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "33", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim ncolumn_opt As Integer
    ncolumn_opt = fNcolumn(mprice, "28", 1)
    If ncolumn_opt = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
        
    Do While m_postavshik(i) = "AVALON"
               
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 2)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_opt))
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "AVALON"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С AVALON"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub

Private Sub oSporthouse(nomer, ByRef m_postavshik, ByRef m_update_date, ByRef m_kody_1c, forced)

For i = 1 To UBound(m_postavshik)
        If m_postavshik(i) = "Sporthouse" Then
            If forced Then
                Exit For
            ElseIf DatePart("w", Now, vbMonday) = 1 And m_update_date(nomer, 2) <> Date Then
                If MsgBox("Обновить Sporthouse?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next i

    i1 = i
    row_i = i + 9
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    
    Application.Workbooks.Open "\\AKABINET\Doc\Наполнение сайтов\Павел\sporthouse\связь.xlsx"
    
    msvyaz = ActiveWorkbook.ActiveSheet.Range("A1:C" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    Set dict = CreateObject("Scripting.Dictionary")

    For a = 1 To UBound(msvyaz, 1)
        msvyaz(a, 1) = CStr(msvyaz(a, 1))
        dict.Add CStr(msvyaz(a, 2)), a
    Next a
    
    Workbooks.Add

    ActiveWorkbook.Queries.Add name:="yml", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Xml.Tables(Web.Contents(""https://sportpro.ru/yml""))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Attribute:date"", type datetime}})," & Chr(13) & "" & Chr(10) & "    shop = #""Changed Type""{0}[shop]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип"" = Table.TransformColumnTypes(shop,{{""name"", type text}, {""company"", type text}, {""platform"", type text}, {""version"", type t" & _
        "ext}, {""url"", type text}})," & Chr(13) & "" & Chr(10) & "    offers = #""Измененный тип""{0}[offers]," & Chr(13) & "" & Chr(10) & "    offer = offers{0}[offer]," & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"" = Table.TransformColumnTypes(offer,{{""url"", type text}, {""price"", Int64.Type}, {""currencyId"", type text}, {""count"", Int64.Type}, {""categoryId"", Int64.Type}, {""store"", type logical}, {""pickup"", type logical}, {""delivery""" & _
        ", type logical}, {""vendor"", type text}, {""name"", type text}, {""description"", type text}, {""manufacturer_warranty"", type logical}, {""country_of_origin"", type text}, {""barcode"", Int64.Type}, {""vendorCode"", type text}, {""adult"", type logical}, {""weight"", Int64.Type}, {""dimensions"", type text}, {""downloadable"", type logical}, {""Attribute:id"", Int" & _
        "64.Type}, {""Attribute:available"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Измененный тип1"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=yml", _
        Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [yml]")
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
        .ListObject.DisplayName = "yml"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    price_name = ActiveWorkbook.name
    
    mprice = ActiveWorkbook.ActiveSheet.Range("A1:Q" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    
    Application.ScreenUpdating = False
    For Each wbk In Application.Workbooks
        If wbk.name = "связь.xlsx" Then
            Workbooks("связь.xlsx").Close False
        ElseIf wbk.name = price_name Then
            Workbooks(price_name).Close False
        End If
    Next wbk
    Application.ScreenUpdating = True
    
    'Поиск номеров столбца с наименованиями
    Dim ncolumn_artikul As Integer
    ncolumn_artikul = fNcolumn(mprice, "vendorCode", 1)
    If ncolumn_artikul = 0 Then GoTo ended
    
    Dim ncolumn_rrc As Integer
    ncolumn_rrc = fNcolumn(mprice, "price", 1)
    If ncolumn_rrc = 0 Then GoTo ended
    
    Dim j As Long
    For j = 1 To UBound(mprice)
        If IsNumeric(mprice(j, ncolumn_artikul)) = True And mprice(j, ncolumn_artikul) <> sEmpty Then
            mprice(j, ncolumn_artikul) = Application.Trim(CStr(mprice(j, ncolumn_artikul)))
        End If
        
        If Application.Trim(mprice(j, ncolumn_artikul)) <> sEmpty And dict1.exists(mprice(j, ncolumn_artikul)) = False Then
            dict1.Add mprice(j, ncolumn_artikul), j
        End If
    Next j
    
    Set proschet = Workbooks(abook).Worksheets("Просчет цен")
    Do While m_postavshik(i) = "Sporthouse"
                
        If dict.exists(m_kody_1c(i)) = True Then
            
            row_svyaz = dict.Item(m_kody_1c(i))
            privyazka = msvyaz(row_svyaz, 1)
            discount = msvyaz(row_svyaz, 3)
            row_privyazka_in_pricelist = dict1.Item(privyazka)
            
            If row_privyazka_in_pricelist <> sEmpty Then
                
                If mprice(row_privyazka_in_pricelist, ncolumn_rrc) = vbNullString Or _
                mprice(row_privyazka_in_pricelist, ncolumn_rrc) = sEmpty Then
                    mprice(row_privyazka_in_pricelist, ncolumn_rrc) = 0
                End If
                
                Total_1 = CLng(mprice(row_privyazka_in_pricelist, ncolumn_rrc))
                Total_2 = Round(Total_1 * discount)
                
                If Total_2 = 0 Or Total_2 = sEmpty Then
                    proschet.Cells(row_i, 18).Value = 0
                    proschet.Cells(row_i, 19).Value = 0
                Else
                    proschet.Cells(row_i, 18).Value = Total_1
                    proschet.Cells(row_i, 19).Value = Total_2
                End If
                
            Else
                proschet.Cells(row_i, 18).Value = 0
                proschet.Cells(row_i, 19).Value = 0
            End If
            
        Else
            proschet.Cells(row_i, 18).Value = 0
            proschet.Cells(row_i, 19).Value = 0
        End If
        
        i = i + 1
        row_i = row_i + 1
        
        If i > UBound(m_postavshik) Then Exit Do
        
    Loop
    
    Erase mprice, msvyaz
    Set dict = Nothing
    Set dict1 = Nothing
    If forced = False Then
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 1).Value = "Sporthouse"
        Workbooks(abook).Worksheets("ДатыО").Cells(nomer, 2).Value = Date
    End If
    
    Workbooks(abook).Activate

    Set uf1 = New UserForm2
    uf1.Caption = "Коды 1С Sporthouse"
    
    For i1 = i1 + 9 To row_i - 1 Step 1
        If uf1.TextBox1 = sEmpty Then
            uf1.TextBox1 = Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        Else
            uf1.TextBox1 = uf1.TextBox1 & vbCrLf & Workbooks(abook).Worksheets("Просчет цен").Cells(i1, 17).Value
        End If
    Next i1
    
    uf1.Show vbModeless
    
ended:
End Sub
