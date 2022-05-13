Attribute VB_Name = "Module1"
Public IsCancelled
Public Sub Requestmaker()

    Dim a, b, c() As String, EmailsArray() As Variant, d

    ReDim EmailsArray(0)

    TotalArray = ActiveWorkbook.ActiveSheet.Range("A1:J" & _
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    UserForm1.Show

    If UserForm1.CommandButton2.Cancel = True Or IsCancelled = True Then
        Exit Sub
    End If

    'Цикл по листу
    For i = 2 To ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row

        b = 0
    
        If ActiveWorkbook.ActiveSheet.Cells(i, 1).Font.Color = 255 _
        Or ActiveWorkbook.ActiveSheet.Cells(i, 10).Font.Color = 255 Then
            GoTo continue
            '3 месяца
        ElseIf ActiveWorkbook.ActiveSheet.Cells(i, 1).Interior.Color = 65535 Then
            d = 93
            '4 месяца
        ElseIf ActiveWorkbook.ActiveSheet.Cells(i, 1).Interior.Color = 15773696 Then
            d = 124
            '6 месяцев
        ElseIf ActiveWorkbook.ActiveSheet.Cells(i, 1).Interior.Color = 255 Then
            d = 186
            '2 месяца
        Else
            d = 62
        End If

        'Если Текущая дата - Дата последнего прайса >= Периоду D(2/3/4/6 месяцев) и в столбце MailSend значение даты, тогда
        If Date - TotalArray(i, 8) >= d And IsDate(TotalArray(i, 9)) = True Then
            If Date - TotalArray(i, 9) >= 14 Then
                If InStr(TotalArray(i, 10), ";") <> 0 Then
                    c = Split(TotalArray(i, 10), ";")
                    For i1 = 0 To UBound(c, 1)
                        If InStr(c(i1), "@") <> 0 Then
                            
                            If UserForm1.CheckBox1 = True Then
                                If Trim(StrConv(TotalArray(i, 2), vbLowerCase)) <> "зелененький" Then
                                    EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i1))
                                    ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                                End If
                            ElseIf UserForm1.CheckBox2 = True Then
                                If Trim(StrConv(TotalArray(i, 3), vbLowerCase)) <> "серенький" Then
                                    EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i1))
                                    ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                                End If
                            ElseIf UserForm1.CheckBox3 = True Then
                                If Trim(StrConv(TotalArray(i, 4), vbLowerCase)) <> "оранжевый" Then
                                    EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i1))
                                    ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                                End If
                            ElseIf UserForm1.CheckBox4 = True Then
                                If Trim(StrConv(TotalArray(i, 5), vbLowerCase)) <> "бордовый" Then
                                    EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i1))
                                    ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                                End If
                            Else
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i1))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If

                            If UserForm1.CheckBox5 = False Then
                                ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = Date
                            End If
                        Else
                            ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = "Некорректный адрес электронной почты"
                        End If
                    Next i1
                Else
                    If InStr(TotalArray(i, 10), "@") <> 0 Then
                    
                        If UserForm1.CheckBox1 = True Then
                            If Trim(StrConv(TotalArray(i, 2), vbLowerCase)) <> "зелененький" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox2 = True Then
                            If Trim(StrConv(TotalArray(i, 3), vbLowerCase)) <> "серенький" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox3 = True Then
                            If Trim(StrConv(TotalArray(i, 4), vbLowerCase)) <> "оранжевый" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox4 = True Then
                            If Trim(StrConv(TotalArray(i, 5), vbLowerCase)) <> "бордовый" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        Else
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    
                        If UserForm1.CheckBox5 = False Then
                            ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = Date
                        End If
                
                    Else
                        ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = "Некорректный адрес электронной почты"
                    End If
                End If
            End If
        ElseIf Date - TotalArray(i, 8) >= d Then
            If InStr(TotalArray(i, 10), ";") <> 0 Then
                c = Split(TotalArray(i, 10), ";")
                For i2 = 0 To UBound(c, 1)
                    If InStr(c(i2), "@") <> 0 Then
                        
                        If UserForm1.CheckBox1 = True Then
                            If Trim(StrConv(TotalArray(i, 2), vbLowerCase)) <> "зелененький" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i2))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox2 = True Then
                            If Trim(StrConv(TotalArray(i, 3), vbLowerCase)) <> "серенький" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i2))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox3 = True Then
                            If Trim(StrConv(TotalArray(i, 4), vbLowerCase)) <> "оранжевый" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i2))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        ElseIf UserForm1.CheckBox4 = True Then
                            If Trim(StrConv(TotalArray(i, 5), vbLowerCase)) <> "бордовый" Then
                                EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i2))
                                ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                            End If
                        Else
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(c(i2))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    
                        If UserForm1.CheckBox5 = False Then
                            ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = Date
                        End If
                    Else
                        ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = "Некорректный адрес электронной почты"
                    End If
                Next i2
            Else
                If InStr(TotalArray(i, 10), "@") <> 0 Then
                
                    If UserForm1.CheckBox1 = True Then
                        If Trim(StrConv(TotalArray(i, 2), vbLowerCase)) <> "зелененький" Then
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    ElseIf UserForm1.CheckBox2 = True Then
                        If Trim(StrConv(TotalArray(i, 3), vbLowerCase)) <> "серенький" Then
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    ElseIf UserForm1.CheckBox3 = True Then
                        If Trim(StrConv(TotalArray(i, 4), vbLowerCase)) <> "оранжевый" Then
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    ElseIf UserForm1.CheckBox4 = True Then
                        If Trim(StrConv(TotalArray(i, 5), vbLowerCase)) <> "бордовый" Then
                            EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                            ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                        End If
                    Else
                        EmailsArray(UBound(EmailsArray)) = "To: " & Trim(TotalArray(i, 10))
                        ReDim Preserve EmailsArray(UBound(EmailsArray, 1) + 1)
                    End If
                
                    If UserForm1.CheckBox5 = False Then
                        ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = Date
                    End If
            
                Else
                    ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = "Некорректный адрес электронной почты"
                End If
            End If
        End If
continue:
    Next i

    Workbooks.Add
    RequestsFilename = ActiveWorkbook.Name

    For i = 0 To UBound(EmailsArray, 1)
        ActiveWorkbook.ActiveSheet.Cells(i + 1, 1).Value = EmailsArray(i)
    Next
  
    'Запуск Outlook и процесса рассылки

    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
    Dim lr As Long, lLastR As Long

    If UserForm1.CheckBox5 = False Then
        If MsgBox("Запустить рассылку?", vbYesNo, vbDefaultButton2) = vbYes Then
            On Error Resume Next
            'Попытка подключения к OutLook, если он уже открыт
            Set objOutlookApp = GetObject(, "Outlook.Application")
            Err.Clear                            'Outlook закрыт, очищаем ошибку
            If objOutlookApp Is Nothing Then
                Set objOutlookApp = CreateObject("Outlook.Application")
            End If
            'Произошла ошибка создания объекта - выход
            If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
        
            lLastR = Cells(Rows.Count, 1).End(xlUp).Row
            '            Цикл от второй строки(начало данных с адресами) до последней ячейки таблицы
            '
            '        Записывается тело письма из Word файла
            Set Word = CreateObject("word.application")
            Set doc = Word.Documents.Open _
            (filename:="C:\Users\argakov\Desktop\Email\MailSend\body_price.docx", ReadOnly:=True)
            MsgTxt = doc.Range(Start:=doc.Paragraphs(1).Range.Start, _
            End:=doc.Paragraphs(doc.Paragraphs.Count).Range.End)
            x = 0
            For lr = 1 To lLastR
                Set objMail = objOutlookApp.CreateItem(0)   'Создаем новое сообщение
                    'Создаем сообщение
                With objMail
                    .BCC = Cells(lr, 1).Value 'Адрес получателя
                    .Subject = "Запрос прайсов для актуализации" ' Тема сообщения
                    .BodyFormat = olformatPlain
                    .Body = MsgTxt 'Текст сообщения
                    .Send
                End With
                x = x + 1
                ActiveWorkbook.ActiveSheet.Cells(lr, 2).Value = " + "
            Next lr
        
            Word.Quit
            Set objOutlookApp = Nothing: Set objMail = Nothing: Set Word = Nothing: Set doc = Nothing
            Application.ScreenUpdating = True
            MsgBox "Успешно отправлено " & x & " писем"
        Else
            Exit Sub
        End If
    End If

    Workbooks(RequestsFilename).SaveAs filename:=Environ("PRICEREQUESTS") & "Запрос прайсов " & Date & ".xlsx", FileFormat:=xlOpenXMLWorkbook

End Sub
