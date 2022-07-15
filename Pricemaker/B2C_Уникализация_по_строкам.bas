Attribute VB_Name = "B2C_Уникализация_по_строкам"
Option Private Module
Option Explicit

Public Sub fMixer(ByRef ProgruzkaMainArray, ByRef TotalCitiesCount As Long, NRowProgruzka As Long)

    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim schetdict As Object
    Dim mceny() As String
    Dim tempmceny As String
    Dim NColumnProgruzkaArray() As String
    Dim m() As Long
    Dim m1 As Long

    Set schetdict = CreateObject("Scripting.Dictionary")
      
    For i = 5 To TotalCitiesCount - 15
    
        If schetdict.exists(CLng(ProgruzkaMainArray(NRowProgruzka, i))) = True Then
            GoTo continue
        End If

        ReDim mceny(0)
        Dim step As Long
        step = -1
        For j = 5 To TotalCitiesCount - 15
            If ProgruzkaMainArray(NRowProgruzka, j) = ProgruzkaMainArray(NRowProgruzka, i) Then
                step = step + 1
                ReDim Preserve mceny(step)
                mceny(step) = j
            End If
        Next j
        tempmceny = Join(mceny, ",")
        schetdict.Add CLng(ProgruzkaMainArray(NRowProgruzka, i)), tempmceny
continue:
    Next i

    ReDim m(UBound(schetdict.Keys))
    For i = LBound(schetdict.Keys) To UBound(schetdict.Keys)
        m(i) = CLng(schetdict.Keys()(i))
    Next i
    
    N = UBound(m)
    
    Do While N > 0
        For i = 0 To N
            If m(i) > m(N) Then
                m1 = m(N)
                m(N) = m(i)
                m(i) = m1
            End If
        Next i
        N = N - 1
    Loop
    
    Dim PriceStep As Long
    For i = 0 To UBound(m)
        
        NColumnProgruzkaArray = Split(schetdict.Item(m(i)), ",")
        If i = 0 Then
            For j = 0 To UBound(NColumnProgruzkaArray)
                PriceStep = GetPriceStep("Б2С СОТ/Торуда", CLng(NColumnProgruzkaArray(j)), CLng(ProgruzkaMainArray(NRowProgruzka, 4)))
                ProgruzkaMainArray(NRowProgruzka, NColumnProgruzkaArray(j)) = ProgruzkaMainArray(NRowProgruzka, NColumnProgruzkaArray(j)) + PriceStep * j
            Next j
            N = CLng(NColumnProgruzkaArray(UBound(NColumnProgruzkaArray)))
        ElseIf m(i) > ProgruzkaMainArray(NRowProgruzka, N) Then
            For j = 0 To UBound(NColumnProgruzkaArray)
                PriceStep = GetPriceStep("Б2С СОТ/Торуда", CLng(NColumnProgruzkaArray(j)), CLng(ProgruzkaMainArray(NRowProgruzka, 4)))
                ProgruzkaMainArray(NRowProgruzka, NColumnProgruzkaArray(j)) = ProgruzkaMainArray(NRowProgruzka, NColumnProgruzkaArray(j)) + PriceStep * j
            Next j
            N = CLng(NColumnProgruzkaArray(UBound(NColumnProgruzkaArray)))
        Else
            For j = 0 To UBound(NColumnProgruzkaArray)
                PriceStep = GetPriceStep("Б2С СОТ/Торуда", CLng(NColumnProgruzkaArray(j)), CLng(ProgruzkaMainArray(NRowProgruzka, 4)))
                ProgruzkaMainArray(NRowProgruzka, NColumnProgruzkaArray(j)) = ProgruzkaMainArray(NRowProgruzka, N) + PriceStep * (j + 1)
            Next j
            N = CLng(NColumnProgruzkaArray(UBound(NColumnProgruzkaArray)))
        End If
            
    Next i
End Sub

Public Sub fMixerEK(ByRef ProgruzkaKompkreslaTerminalArray, ByVal TotalCitiesCount As Long, NRowProgruzka As Long)
    
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim schetdict As Object
    Dim mceny
    Dim mceny1() As String
    Dim m
    Dim m1
    
    Set schetdict = CreateObject("Scripting.Dictionary")
    
    For i = 5 To TotalCitiesCount - 15

        If schetdict.exists(ProgruzkaKompkreslaTerminalArray(NRowProgruzka, i)) = True Then GoTo continue
        mceny = 0

        For j = 5 To TotalCitiesCount - 15

            If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, j) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, i) Then

                If mceny = 0 Then
                    mceny = j
                Else
                    mceny = mceny & "," & j
                End If

            End If

        Next j

        schetdict.Add ProgruzkaKompkreslaTerminalArray(NRowProgruzka, i), mceny

continue:
    Next i

    m = schetdict.Keys
    N = UBound(m)

    Do While N > 0

        For i = 0 To N

            If m(i) > m(N) Then
                m1 = m(N)
                m(N) = m(i)
                m(i) = m1
            End If

        Next i

        N = N - 1

    Loop

    For i = 0 To UBound(m)
        mceny1 = Split(schetdict.Item(m(i)), ",")
        If i = 0 Then

            If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 7500 Then
                For j = 0 To UBound(mceny1)
                    If mceny1(j) >= 120 And mceny1(j) < 126 Then
                        If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 1 * j
                        End If
                    Else
                        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                    End If
                Next j
            Else
                For j = 0 To UBound(mceny1)
                    If mceny1(j) >= 108 And mceny1(j) < 120 _
                    Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                    Or mceny1(j) = 127 Then
                        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                    Else
                        ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 1 * j
                    End If
                Next j

            End If

            N = mceny1(UBound(mceny1))

        Else

            If m(i) > ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) Then

                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 120 And mceny1(j) < 126 Then
                            If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                            Else
                                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 1 * j
                            End If
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                        End If
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 108 And mceny1(j) < 120 Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                        Or mceny1(j) = 127 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 10 * j
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) + 1 * j
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            Else

                If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 120 And mceny1(j) < 126 Then
                            If ProgruzkaKompkreslaTerminalArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) + 10 * (j + 1)
                            Else
                                ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) + 1 * (j + 1)
                            End If
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) + 10 * (j + 1)
                        End If
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 108 And mceny1(j) < 120 Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                        Or mceny1(j) = 127 Then
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) + 10 * (j + 1)
                        Else
                            ProgruzkaKompkreslaTerminalArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaTerminalArray(NRowProgruzka, N) + 1 * (j + 1)
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            End If

        End If

    Next i

    schetdict.RemoveAll

End Sub

Public Sub fMixerK(ByRef ProgruzkaKompkreslaDverArray, ByVal TotalCitiesCount As Long, NRowProgruzka As Long)

    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim schetdict As Object
    Dim mceny
    Dim mceny1() As String
    Dim m
    Dim m1
    
    Set schetdict = CreateObject("Scripting.Dictionary")
    
    For i = 5 To TotalCitiesCount - 15

        If schetdict.exists(ProgruzkaKompkreslaDverArray(NRowProgruzka, i)) = True Then GoTo continue
        mceny = 0

        For j = 5 To TotalCitiesCount - 15

            If ProgruzkaKompkreslaDverArray(NRowProgruzka, j) = ProgruzkaKompkreslaDverArray(NRowProgruzka, i) Then

                If mceny = 0 Then
                    mceny = j
                Else
                    mceny = mceny & "," & j
                End If

            End If

        Next j

        schetdict.Add ProgruzkaKompkreslaDverArray(NRowProgruzka, i), mceny

continue:
    Next i

    m = schetdict.Keys
    N = UBound(m)

    Do While N > 0

        For i = 0 To N

            If m(i) > m(N) Then
                m1 = m(N)
                m(N) = m(i)
                m(i) = m1
            End If

        Next i

        N = N - 1

    Loop

    For i = 0 To UBound(m)
        mceny1 = Split(schetdict.Item(m(i)), ",")
        If i = 0 Then

            If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 7500 Then
                For j = 0 To UBound(mceny1)
                    If mceny1(j) >= 120 And mceny1(j) < 126 Then
                        If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 150000 Then
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                        Else
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 1 * j
                        End If
                    Else
                        ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                    End If
                Next j
            Else
                For j = 0 To UBound(mceny1)
                    If mceny1(j) >= 108 And mceny1(j) < 120 _
                    Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                    Or mceny1(j) = 127 Then
                        ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                    Else
                        ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 1 * j
                    End If
                Next j

            End If

            N = mceny1(UBound(mceny1))

        Else

            If m(i) > ProgruzkaKompkreslaDverArray(NRowProgruzka, N) Then

                If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 120 And mceny1(j) < 126 Then
                            If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                            Else
                                ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 1 * j
                            End If
                        Else
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                        End If
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 108 And mceny1(j) < 120 _
                        Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                        Or mceny1(j) = 127 Then
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 10 * j
                        Else
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) + 1 * j
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            Else

                If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 120 And mceny1(j) < 126 Then
                            If ProgruzkaKompkreslaDverArray(NRowProgruzka, 4) > 150000 Then
                                ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, N) + 10 * (j + 1)
                            Else
                                ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, N) + 1 * (j + 1)
                            End If
                        Else
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, N) + 10 * (j + 1)
                        End If
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 108 And mceny1(j) < 120 Or mceny1(j) >= 148 And mceny1(j) <= 155 _
                        Or mceny1(j) = 127 Then
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, N) + 10 * (j + 1)
                        Else
                            ProgruzkaKompkreslaDverArray(NRowProgruzka, mceny1(j)) = ProgruzkaKompkreslaDverArray(NRowProgruzka, N) + 1 * (j + 1)
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            End If

        End If

    Next i

    schetdict.RemoveAll

End Sub

Public Sub fMixer_ПМК(ByRef ProgruzkaPMKArray, ByRef TotalCitiesCount As Long, ByRef NRowProgruzka As Long)
    
    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim schetdict As Object
    Dim mceny
    Dim mceny1() As String
    Dim m
    Dim m1
    
    Set schetdict = CreateObject("Scripting.Dictionary")
    
    For i = 5 To TotalCitiesCount - 15
        If ProgruzkaPMKArray(NRowProgruzka, i) > 0 Then
            If schetdict.exists(ProgruzkaPMKArray(NRowProgruzka, i)) = True Then GoTo continue
            mceny = 0
    
            For j = 5 To TotalCitiesCount - 15
                If ProgruzkaPMKArray(NRowProgruzka, j) = ProgruzkaPMKArray(NRowProgruzka, i) Then
    
                    If mceny = 0 Then
                        mceny = j
                    Else
                        mceny = mceny & "," & j
                    End If
    
                End If
            Next j
            schetdict.Add ProgruzkaPMKArray(NRowProgruzka, i), mceny
        End If

continue:
    Next i

    m = schetdict.Keys
    N = UBound(m)

    Do While N > 0

        For i = 0 To N

            If m(i) > m(N) Then
                m1 = m(N)
                m(N) = m(i)
                m(i) = m1
            End If

        Next i

        N = N - 1

    Loop

    For i = 0 To UBound(m)
        mceny1 = Split(schetdict.Item(m(i)), ",")
        If i = 0 Then

            If ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 Then
                For j = 0 To UBound(mceny1)
                    ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                Next j
            Else
                For j = 0 To UBound(mceny1)
                    
                    If mceny1(j) >= 92 And mceny1(j) <= 98 Then
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                    ElseIf mceny1(j) = 100 Then
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                    ElseIf mceny1(j) = 101 Then
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                    ElseIf mceny1(j) = 102 Then
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                    Else
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                    End If
                Next j
            End If

            N = mceny1(UBound(mceny1))

        Else
            If m(i) > ProgruzkaPMKArray(NRowProgruzka, N) Then

                If ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 92 And mceny1(j) <= 98 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                        ElseIf mceny1(j) = 100 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                        ElseIf mceny1(j) = 101 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                        ElseIf mceny1(j) = 102 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 10 * j
                        Else
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) + 1 * j
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            Else

                If ProgruzkaPMKArray(NRowProgruzka, 4) > 7500 Then
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 120 And mceny1(j) Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 10 * (j + 1)
                        End If
                    Next j
                Else
                    For j = 0 To UBound(mceny1)
                        If mceny1(j) >= 92 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 10 * (j + 1)
                        Else
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 1 * (j + 1)
                        End If
                        
                        If mceny1(j) >= 92 And mceny1(j) <= 98 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 10 * j
                        ElseIf mceny1(j) = 100 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 1 * j
                        ElseIf mceny1(j) = 101 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 1 * j
                        ElseIf mceny1(j) = 102 Then
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 10 * j
                        Else
                            ProgruzkaPMKArray(NRowProgruzka, mceny1(j)) = ProgruzkaPMKArray(NRowProgruzka, N) + 1 * j
                        End If
                    Next j

                End If

                N = mceny1(UBound(mceny1))

            End If

        End If

    Next i

    schetdict.RemoveAll

End Sub


