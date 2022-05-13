Attribute VB_Name = "GeneralFunctions"
Option Private Module
Option Explicit

'� �������� ����� ��������� ������ ���� � ����� ���������
Public Sub CreateSheet(ByRef name As String)

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
    ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    ws.name = name
    
End Sub

Public Sub Autosave(ByRef Postavshik As String, ProgruzkaFilename As String, ByRef Site As String)

On Error GoTo ended
    
    Application.ScreenUpdating = False
    
    'B2�
    Dim NewFilename As String
    If Site = "�2� ���" Then
        NewFilename = "���������"
    ElseIf Site = "�2� ������" Then
        NewFilename = "��������� ������"
    ElseIf Site = "�2� ��� �����" Then
        NewFilename = "��������� �����"
    ElseIf Site = "�2� ������ �����" Then
        NewFilename = "��������� ������ �����"
    ElseIf Site = "�2� ��������" Then
        NewFilename = "��������� 620"
    ElseIf Site = "�2� �������� �����" Then
        NewFilename = "��������� 620 �����"
    ElseIf Site = "�2� ��� ����� � �����" Then
        NewFilename = "��������� ����."
    ElseIf Site = "�2� ������ ����� � �����" Then
        NewFilename = "��������� ������ ����."
    ElseIf Site = "�2� ����������" Then
        NewFilename = "��������� �"
    ElseIf Site = "�2� ���������� �����" Then
        NewFilename = "��������� � �����"
    ElseIf Site = "�2� �����" Then
        NewFilename = "��������� �����"
    ElseIf Site = "�2� ����� �����" Then
        NewFilename = "��������� ����� �����"
    ElseIf Site = "�2� ���" Then
        NewFilename = "��������� DDI"
    ElseIf Site = "�2� ���" Then
        NewFilename = "��������� ���"
    ElseIf Site = "�2� ��� 1034" Then
        NewFilename = "��������� ��� 1034"
    ElseIf Site = "�2� ��� �����" Then
        NewFilename = "��������� ��� �����"
    ElseIf Site = "�2� ��� 1034 �����" Then
        NewFilename = "��������� ��� 1034 �����"
    'B2B
    ElseIf Site = "�2� ���" Then
        NewFilename = "��������� B2B ���"
    ElseIf Site = "�2� ��� 830/917" Then
        NewFilename = "��������� B2B 830-917 ���"
    ElseIf Site = "�2� ������" Then
        NewFilename = "��������� B2B ������"
    ElseIf Site = "�2� ���������" Then
        NewFilename = "��������� B2B ���������"
    ElseIf Site = "�2� ��� ��������" Then
        NewFilename = "��������� B2B ��� ��������"
    Else
        MsgBox "����������� ������ �������� Site � ������� Autosave"
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
            MsgBox "����������� � ��������! " & UFKods_Array(i)
            DublicateFind = True
            Exit Function
        End If

        If ProverkaDict.exists(UFKods_Array(i)) = True Then
            MsgBox "�����! " & UFKods_Array(i)
            DublicateFind = True
            Exit Function
        Else
            ProverkaDict.Add UFKods_Array(i), i
        End If

        If UserForm1.OptionButton3 = True Then
            If InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "�") = 0 Then
                MsgBox UFKods_Array(i) & " ��������� ��������� �� �������!"
                DublicateFind = True
                Exit Function
            End If
        ElseIf UserForm1.OptionButton2 = True Then
            If InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat)) = 0 Then
                MsgBox UFKods_Array(i) & " ��������� ��������� �� �������!"
                DublicateFind = True
                Exit Function
            ElseIf InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "�") <> 0 Or _
            InStr(ProschetArray(Proschet_Dict(UFKods_Array(i)), 7), Trim(kat) & "��") <> 0 Then
                MsgBox UFKods_Array(i) & " ���������� ��� �����!"
                DublicateFind = True
                Exit Function
            End If
        End If

    Next i

End Function

Public Function FillInfocolumnRetailPrice(ByRef RetailPrice As Long, ByRef RetailPriceInfo As Long) As String

    If RetailPrice <> 0 And RetailPriceInfo <> 0 Then
        FillInfocolumnRetailPrice = RetailPrice & " (���), " & RetailPriceInfo & " (���)"
    ElseIf RetailPrice <> 0 Then
        FillInfocolumnRetailPrice = RetailPrice & " (���)"
    ElseIf RetailPriceInfo <> 0 Then
        FillInfocolumnRetailPrice = RetailPriceInfo & " (���)"
    End If

End Function

Public Function FillInfocolumnJU(ByRef Progruzka_Name, JU, ByRef Komplektuyeshee, _
ByRef Quantity, GorodOtpravki, ByRef ObreshetkaPrice) As String

If Progruzka_Name = "���/������" Then
    If JU = 0 Or Len(JU) = 0 Then
        FillInfocolumnJU = "���"
    ElseIf Komplektuyeshee = "�" Or Komplektuyeshee = "�1" Or Komplektuyeshee = "�2" Or Komplektuyeshee = "�" Then
        FillInfocolumnJU = "���"
    Else
        If Quantity > 1 Then
            If DictLgotniiGorod.exists(GorodOtpravki) = False Then
                FillInfocolumnJU = ObreshetkaPrice / 2 / Quantity & " +30% ��������� �������� DPD /" & Quantity
            Else
                FillInfocolumnJU = ObreshetkaPrice / 2 / Quantity & " +30% ��������� �������� DPD /" & Quantity
            End If
        ElseIf DictLgotniiGorod.exists(GorodOtpravki) = False Then
            FillInfocolumnJU = ObreshetkaPrice / 2 & " +30% ��������� �������� DPD"
        Else
            FillInfocolumnJU = ObreshetkaPrice / 2 & " +30% ��������� �������� DPD"
        End If
    End If
ElseIf Progruzka_Name = "�����" Then
    If JU = 0 Or JU = vbNullString Then
        FillInfocolumnJU = "���"
    ElseIf Komplektuyeshee = "�" Or Komplektuyeshee = "�1" Or Komplektuyeshee = "�2" Then
        FillInfocolumnJU = "���"
    Else
        FillInfocolumnJU = "��� ������� ������� �������� ������ ��������� �� ������������!"
    End If
Else
    MsgBox "����������� ������ �������� Progruzka_Name ������� FillInfocolumnJU"
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
    
        If Komplektuyeshee = "�" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: " & Margin & ", ����������� - " & Margin & " (����� ����!)"
        ElseIf Komplektuyeshee = "�2" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: 2, ����������� - 1.9" & " (����� ����!)"
        ElseIf NeiskluchenieToruda = True Then
            FillInfocolumnMargin = "1.35 ������," & 1.35 + SNG_MARGIN & " � ���, ����������� - 1.25" & " (����� ����!)"
        ElseIf Komplektuyeshee = "�" Then
            FillInfocolumnMargin = "2 - ������, ����������� - 1.9" & " (����� ����!)"
        ElseIf Komplektuyeshee = "�1" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: 1.7, ����������� - 1.6" & " (����� ����!)"
        Else
            FillInfocolumnMargin = Margin & " - ������," & Margin + SNG_MARGIN & " � ���, ����������� - " & Margin - 0.1 & " (����� ����!)"
        End If
        
    ElseIf Promo = False Then
    
        If Komplektuyeshee = "�" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: " & Margin & ", ����������� - " & Margin & " (�� ����� ����)"
        ElseIf Komplektuyeshee = "�2" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: 2, ����������� - 1.9" & " (�� ����� ����)"
        ElseIf NeiskluchenieToruda = True Then
            FillInfocolumnMargin = "1.5 ������," & 1.5 + SNG_MARGIN & " � ���, ����������� - 1.4" & " (�� ����� ����)"
        ElseIf Komplektuyeshee = "�" Then
            FillInfocolumnMargin = "2, ����������� - 1.9" & " (�� ����� ����)"
        ElseIf Komplektuyeshee = "�1" Then
            FillInfocolumnMargin = "������� ����������� � ������� ������: 1.7, ����������� - 1.6" & " (�� ����� ����)"
        Else
            FillInfocolumnMargin = Margin & " - ������," & Margin + SNG_MARGIN & " � ���, ����������� - " & Margin - 0.1 & " (�� ����� ����)"
        End If
        
    End If


End Function

'������ �������� � ������� TypeZabor ���������
'TypeZabor = FillInfocolumnTypeZabor(str, ZaborPrice, ProschetArray(NRowsArray(d), 12), NRowsArray(d), LgotniiGorod)
Public Function FillInfocolumnTypeZabor(Komponent, ZaborPrice, RazgruzkaPrice, NRowProschet, LgotniiGorod, tps, Optional spez As Integer = 0) As String
    
    
    
    '���� ���-�� ���� > 1, �����
    If ProschetArray(NRowProschet, 5) > 1 Then
        If LgotniiGorod = True Then
            If spez = 1 Then
                '���� ����� ��������, � ?spez = 1?
                'TypeZabor:
                '��� ������ � ��: DPD
                '��� ������ � ���: �� �������� �� ������� "��� ������"
                '��������� ������:
                '� ������ (������ ������) = 200 / ���-�� ����
                '� ������ (������) = 0
                '� ��� = ����� �� �������� / ���-�� ����
                '���. �������� = ��������� �� �������� / ���-�� ����
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� DPD(��� �������� �� ������), ��� ������ � ���-  " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - � ������(������ ������), 0 - � ������(������)) " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - � ���, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            ElseIf spez = 2 Then
                '���� ����� ��������, � ?spez = 2?
                'TypeZabor:
                '��� ������ � ��: Pickpoint
                '��� ������ � ���: DPD
                '��������� ������:
                '� ��� = ����� �� �������� / ���-�� ����
                '���. �������� = ��������� �� �������� / ���-�� ����
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� Pickpoint(��� �������� �� ������), ��� ������ � ���-DPD, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - � ���, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            Else
                '���� ����� ��������, �����
                'TypeZabor:
                '��� ������ � ��: DPD
                '��� ������ � ���: �� �������� �� ������� "��� ������"
                '��������� ������:
                '� ������ = 200 / ���-�� ����
                '� ��� = ����� �� �������� / ���-�� ����
                '���. �������� = ��������� �� �������� / ���-�� ����
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� DPD(��� �������� �� ������), ��� ������ � ���- " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - � ������, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
                " - � ���, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            End If
        ElseIf LgotniiGorod = False Then
            '���� ����� ����������, �����
            'TypeZabor:
            '��� ������ � ��: �� �������� �� ������� "��� ������"
            '��������� ������:
            '� ������ = 200 / ���-�� ����
            '� ��� = ����� �� �������� / ���-�� ����
            '���. �������� = ��������� �� �������� / ���-�� ����
            '�������� �� ������� ������� = tps
            FillInfocolumnTypeZabor = ProschetArray(NRowProschet, 8) & " " & CLng(ZaborPrice) / ProschetArray(NRowProschet, 5) & " - � ������, " & CLng(ProschetArray(NRowProschet, 13)) / ProschetArray(NRowProschet, 5) & _
            " - � ���, " & CLng(RazgruzkaPrice) / ProschetArray(NRowProschet, 5) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
            tps & " - �������� �� ������� �������"
        End If
    Else
        If LgotniiGorod = True Then
            If spez = 1 Then
                '���� ����� �������� � ?spez = 1?, �����
                'TypeZabor:
                '��� ������ � ��: DPD
                '��� ������ � ���: �� �������� �� ������� "��� ������"
                '��������� ������:
                '� ������ (������ ������) = 200
                '� ������ (������) = 0
                '� ��� = ����� �� ��������
                '���. �������� = ��������� �� ��������
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� DPD(��� �������� �� ������), ��� ������ � ���-  " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) & " - � ������(������ ������), 0 - � ������(������)) " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - � ���, " & CLng(RazgruzkaPrice) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            ElseIf spez = 2 Then
                '���� ����� ��������, � ?spez = 2?
                'TypeZabor:
                '��� ������ � ��: Pickpoint
                '��� ������ � ���: DPD
                '��������� ������:
                '� ��� = ����� �� ��������
                '���. �������� = ��������� �� ��������
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� Pickpoint(��� �������� �� ������), ��� ������ � ���-DPD, " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - � ���, " & CLng(RazgruzkaPrice) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            Else
                '���� ����� ��������, �����
                'TypeZabor:
                '��� ������ � ��: DPD
                '��� ������ � ���: �� �������� �� ������� "��� ������"
                '��������� ������:
                '� ������ = 200
                '� ��� = ����� �� ��������
                '���. �������� = ��������� �� ��������
                '�������� �� ������� ������� = tps
                FillInfocolumnTypeZabor = "�������� DPD(��� �������� �� ������), ��� ������ � ���- " & ProschetArray(NRowProschet, 8) & ", " & _
                CLng(ZaborPrice) & " - � ������, " & CLng(ProschetArray(NRowProschet, 13)) & _
                " - � ���, " & CLng(RazgruzkaPrice) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
                tps & " - �������� �� ������� �������"
            End If
        Else
            '���� ����� ����������, �����
            'TypeZabor:
            '��� ������ � ��: �� �������� �� ������� "��� ������"
            '��������� ������:
            '� ������ = 200
            '� ��� = ����� �� ��������
            '���. �������� = ��������� �� ��������
            '�������� �� ������� ������� = tps
            FillInfocolumnTypeZabor = ProschetArray(NRowProschet, 8) & " " & CLng(ZaborPrice) & " - � ������, " & CLng(ProschetArray(NRowProschet, 13)) & _
            " - � ���, " & CLng(RazgruzkaPrice) & " - ���. �������� (���������, ����������� ������, ���. �������� � ��.), " & _
            tps & " - �������� �� ������� �������"
        End If
    End If

End Function

Public Function DeclareDatatypeProschet(ByVal NColumn, Value)

'���������
If NColumn = 1 Then
    DeclareDatatypeProschet = CStr(Value)
'�������
ElseIf NColumn = 2 Then
    DeclareDatatypeProschet = CStr(Value)
'��������� - �� ������������ � �������� ����
ElseIf NColumn = 3 Then
    DeclareDatatypeProschet = vbNullString
'���������� ���� - �� ������������ � �������� ����
ElseIf NColumn = 4 Then
    DeclareDatatypeProschet = vbNullString
'���������� ����
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
'����� ��������
ElseIf NColumn = 6 Then
    DeclareDatatypeProschet = CStr(Value)
'�������
ElseIf NColumn = 7 Then
    DeclareDatatypeProschet = CStr(Value)
'��� ������
ElseIf NColumn = 8 Then
    DeclareDatatypeProschet = CStr(Value)
'�����
ElseIf NColumn = 9 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'���
ElseIf NColumn = 10 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'�������� ���
ElseIf NColumn = 11 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'���������
ElseIf NColumn = 12 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'�����
ElseIf NColumn = 13 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Value)
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'������� ��������
ElseIf NColumn = 14 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CDbl(Value)
    Else
        DeclareDatatypeProschet = CDbl(0)
    End If
'����� � �� - �� ������������ � �������� ����
ElseIf NColumn = 15 Then
    DeclareDatatypeProschet = vbNullString
'������� (�������������)
ElseIf NColumn = 16 Then
    DeclareDatatypeProschet = CStr(Value)
'��� 1�
ElseIf NColumn = 17 Then
    DeclareDatatypeProschet = CStr(Value)
'���
ElseIf NColumn = 18 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'������� ����
ElseIf NColumn = 19 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
'��� (����)
ElseIf NColumn = 20 Then
    If IsNumeric(Value) = True Then
        DeclareDatatypeProschet = CLng(Round_Up(Value, 0))
    Else
        DeclareDatatypeProschet = CLng(0)
    End If
Else
    MsgBox "����������� ������� �������� NColumn ������� DeclareDatatypeProschet)"
    Stop
End If

End Function

Public Function GetPriceStep(ByRef Sitename, ByRef NCol_Progruzka As Long, ByRef PriceValue As Long) As Long

    If Sitename = "�2� ���/������" Then
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
        
    ElseIf Sitename = "�2� ���" Then
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

    ElseIf Sitename = "�2� ������" Then
        If NCol_Progruzka < 128 Then
            If PriceValue < 7500 Then
                GetPriceStep = 1
            Else
                GetPriceStep = 10
            End If
        Else
            GetPriceStep = 10
        End If

    ElseIf Sitename = "�2� ���������" Then
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
    
    ElseIf Sitename = "�2� ��� ��������" Then
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
        MsgBox "����������� ������ �������� Sitename ������� GetPriceStep"
        Stop
    End If

End Function
