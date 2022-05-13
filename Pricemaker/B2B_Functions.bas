Attribute VB_Name = "B2B_Functions"
'@Folder("VBAProject")
Option Private Module
Option Explicit

Public Function SetRRC(OldRrc) As Long
    
    If Trim(OldRrc) = "-" Or Trim(OldRrc) = vbNullString _
    Or Trim(OldRrc) = vbNullString Then
        SetRRC = 0
    Else
        SetRRC = CLng(OldRrc)
    End If

End Function

Public Function SetOldPrice(OldPrice) As Long
    
    If Trim(OldPrice) = "-" Or Trim(OldPrice) = vbNullString _
    Or Trim(OldPrice) = vbNullString Then
        SetOldPrice = 0
    Else
        SetOldPrice = CLng(OldPrice)
    End If

End Function

Public Function FindPriceLimit(LimitsArray, ByRef Komponent As String, ByRef Nrow_PriceLimit As Long) As Long
    Dim i As Long
    For i = 1 To UBound(LimitsArray, 1)
        LimitsArray(i, 3) = CStr(LimitsArray(i, 3))
        LimitsArray(i, 4) = CStr(LimitsArray(i, 4))
        If (Komponent & Nrow_PriceLimit) = (LimitsArray(i, 3) & LimitsArray(i, 4)) Then
            FindPriceLimit = LimitsArray(i, 5)
            Exit For
        End If
    Next
End Function

Public Function GetUnique(ByRef DuplicatePrice, ByRef UniquePriceDict, ByRef PriceStep) As Double
    Dim UniqDiff As Long
    If PriceStep > 0 Then
        UniqDiff = PriceStep

        Do While (UniquePriceDict.exists(DuplicatePrice + UniqDiff))
            UniqDiff = UniqDiff + PriceStep
        Loop
        
        GetUnique = DuplicatePrice + UniqDiff

    Else
        GetUnique = DuplicatePrice
    End If
End Function

Public Sub Unify(Progruzka, Nrow_Progruzka, Sitename, ByRef Purpose)
    'Progruzka - весь массив прогрузки
    'Nrow_Progruzka - текущий номер строки прогрузки
    'ProgruzkaLen - Количество столбцов с ценой
    
    Dim NCol_Progruzka As Long
    Dim PriceStep As Long
    Dim NComp As Long
    Dim StartPos As Long
    Dim r1 As Long
    Dim i As Long
    Dim NotUnifedPrice As Long
    Dim UniqPrice As Long
    Dim PriceValue As Long
    Dim UniquePriceDict As Object
    Set UniquePriceDict = CreateObject("Scripting.Dictionary")
    If Purpose = "Row" Then

        For i = 5 To UBound(Progruzka, 2)
            If Progruzka(Nrow_Progruzka, i) <> 0 Then
                If UniquePriceDict.exists(Progruzka(Nrow_Progruzka, i)) = False Then
                    UniquePriceDict.Add Progruzka(Nrow_Progruzka, i), " "
                ElseIf Progruzka(Nrow_Progruzka, i) <> 0 Then
                    PriceValue = CLng(Progruzka(Nrow_Progruzka, i))
                    PriceStep = 0
                    PriceStep = GetPriceStep(Sitename, i, PriceValue)
                    UniqPrice = GetUnique(PriceValue, UniquePriceDict, PriceStep)
                    Progruzka(Nrow_Progruzka, i) = UniqPrice
                    UniquePriceDict.Add UniqPrice, " "
                End If
            End If
        Next

    ElseIf Purpose = "Column" Then

        NComp = 0

        'Находит номер строки с самым первым компонентом кода 1C
        Do While Progruzka(Nrow_Progruzka, 3) = Progruzka(Nrow_Progruzka - NComp, 3)
            NComp = NComp + 1
            StartPos = Nrow_Progruzka - NComp + 1
        Loop

        'Прогоняет по колонкам прогрузки до последней колонки
        For NCol_Progruzka = 5 To UBound(Progruzka, 2)
            'От номера строки первого компонента до номера строки последнего компонента
            For r1 = StartPos To Nrow_Progruzka
                
                NotUnifedPrice = Progruzka(r1, NCol_Progruzka)
                PriceStep = GetPriceStep(Sitename, NCol_Progruzka, NotUnifedPrice)
                'От номера строки первого компонента до номера строки последнего компонента
                For i = StartPos To Nrow_Progruzka
                    If i <> r1 Then
                        If NotUnifedPrice = Progruzka(i, NCol_Progruzka) Then
                            If Progruzka(r1, NCol_Progruzka) <> 0 Then
                                Progruzka(r1, NCol_Progruzka) = Progruzka(r1, NCol_Progruzka) + PriceStep
                                NotUnifedPrice = NotUnifedPrice + PriceStep
                                i = StartPos - 1
                            End If
                        End If
                    End If
                Next i
            Next r1
        Next NCol_Progruzka

    End If
    
End Sub

