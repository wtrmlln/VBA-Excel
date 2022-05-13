Attribute VB_Name = "Main"
Option Explicit

Public AkciaArray As Variant                            '������ ����� "�����" �� A1:F �� ��������� ����������� ������
Public DictAkcia As Object                                   '���� - ��� 1� �� ����� "�����", ������� - ����� ������
Public ProschetArray As Variant                      '�������� ���� �� �1:T �� ��������� ����������� ������ + 1
Public Proschet_Dict As Object                             '���� - ��� 1� �� ����� "������� ���", ������� - ����� ������
Public ProgruzkaMainArray As Variant                     '�������� �����. �������� ����. �� A1:T �� ��������� ����������� ������ + 1
Public ProschetWorkbookName As String
Public kat As String                             '����� �� ��������� "�������"

Public TarifOMAXTArray As Variant                     '����(omaxT) �� �1:K �� ��������� ����������� ������ + 1
Public DictTarifOMAXT As Object                            '���� - ���� (omaxT), ������� NColumnProgruzka, ������� - ����� ������
Public TarifOMAXDArray As Variant                     '����(omax�) �� �1:K �� ��������� ����������� ������ + 1
Public DictTarifOMAXD As Object                           '���� - ����(omax�) ,������� NColumnProgruzka, ������� - ����� ������ (�������� NColumnProgruzka:K)
Public LgotniiGorod As Integer                          '0 - ���� ����� �� ��������, 1 - ���� ��������, ������������ � FillInfocolumnTypeZabor
Public DictLgotniiGorod As Object                                '���� - ������������ ������ �� ����� "��������", ������� - ����� ������
Public SpecPostavshikArray() As Variant                                   '����.����������(A1:A) � ��������� ����������� ������ + 1
Public DictSpecPostavshik As Object                                     '���� - ������ ������� ����� "��������������", ������� - ����� ������
Public DictIsklucheniyaToruda As Object                                     '���� - ��� 1� �� ����� "����������(toruda<=1000), ������� - ����� ������


Public wks As Object
Public KURS_TENGE As Double
Public KURS_BELRUB As Double
Public KURS_SOM As Double
Public KURS_DRAM As Double
Public SNG_MARGIN As Double                                     '������ ����� �������� �� ��������(P7)

Sub Main()
    
    Dim i As Long
    
    ProschetWorkbookName = ActiveWorkbook.name
    
    If ProschetWorkbookName = "������� ��� B2B.xlsb" Then
        Call ���������_�2�
        Exit Sub
    End If

    Set DictSpecPostavshik = CreateObject("Scripting.Dictionary")

    '��������� � ������� DictSpecPostavshik: ���� - ������������ ����.����������, ������� - ����� ������
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "����.-����������(1)" Then
            'SpecPostavshikArray = ����.����������(A1:A) � ��������� ����������� ������ + 1
            SpecPostavshikArray = Workbooks(ProschetWorkbookName).Worksheets("����.-����������(1)").Range("A1:A" & _
            Workbooks(ProschetWorkbookName).Worksheets("����.-����������(1)").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value
            For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("����.-����������(1)").Cells.SpecialCells(xlCellTypeLastCell).Row
                DictSpecPostavshik.Add CStr(SpecPostavshikArray(i, 1)), i
            Next i
        End If
    Next wks

    'ProschetArray = �������� �����. �������� ����. �� A1:T �� ��������� ����������� ������ + 1
    ProschetArray = Workbooks(ProschetWorkbookName).ActiveSheet.Range("A1:T" & _
    Workbooks(ProschetWorkbookName).ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

    'Set schetdict = CreateObject("Scripting.Dictionary")
    Set Proschet_Dict = CreateObject("Scripting.Dictionary")

    '�������� ����� 1� �� ����� � ��������
    '� ���������� � ������� Proschet_Dict: ���� - ��� 1� �� ����� "������� ���", ������� - � ������
    For i = 1 To UBound(ProschetArray, 1)
        If Trim(ProschetArray(i, 17)) <> vbNullString Then
            If Proschet_Dict.exists(Trim(ProschetArray(i, 17))) = True Then
                MsgBox "����� � ��������! " & ProschetArray(i, 17)
                Exit Sub
            End If
            Proschet_Dict.Add Trim(ProschetArray(i, 17)), i
        End If
    Next i

    '�������� �� ����� � ����� omaxT �� ������� NColumnProgruzka
    '� ���������� � ������� DictTarifOMAXT: ���� - �����&����� �� ����� "�����(�maxT)", ������� - ����� ������
    'TarifOMAXTArray = ����(omaxT) �� �1:K �� ��������� ����������� ������ + 1
    TarifOMAXTArray = Workbooks(ProschetWorkbookName).Worksheets("�����(omax�)").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("�����(omax�)").Cells.SpecialCells(xlCellTypeLastCell).Row + 1).Value

    Set DictTarifOMAXT = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(TarifOMAXTArray, 1)
        If Trim(TarifOMAXTArray(i, 1)) <> vbNullString Then
            If DictTarifOMAXT.exists(Trim(TarifOMAXTArray(i, 1))) = True Then
                MsgBox "����� � �������! " & TarifOMAXTArray(i, 1)
                Exit Sub
            End If
            DictTarifOMAXT.Add Trim(TarifOMAXTArray(i, 1)), i
        End If
    Next i

    '�������� �� ����� � ����� omax� �� ������� �
    '� ���������� � ������� DictTarifOMAXD: ���� - �����&����� �� ����� "�����(�max�)", ������� - ����� ������
    'TarifOMAXDArray = ����(omax�) �� �1:K �� ��������� ����������� ������ + 1
    TarifOMAXDArray = Workbooks(ProschetWorkbookName).Worksheets("�����(omax�)").Range("C1:K" & _
    Workbooks(ProschetWorkbookName).Worksheets("�����(omax�)").Cells.SpecialCells(xlCellTypeLastCell).Row).Value
    Set DictTarifOMAXD = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(TarifOMAXDArray, 1)
        If Trim(TarifOMAXDArray(i, 1)) <> vbNullString Then
            If DictTarifOMAXD.exists(Trim(TarifOMAXDArray(i, 1))) = True Then
                MsgBox "����� � �������(�� �����)! " & TarifOMAXDArray(i, 1)
                Exit Sub
            End If
            DictTarifOMAXD.Add Trim(TarifOMAXDArray(i, 1)), i
        End If
    Next i

    '��������� ������ ���� �� ��������
    '���� ���� ���� "����������(toruda<=1000)" - �� ��������� � start3
    '���� ��� - ������� ���� "����������(toruda<=1000)" � ��������� � start3
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "����������(toruda<=1000)" Then
            GoTo start3
        End If
    Next wks

    Call CreateSheet("����������(toruda<=1000)")

start3:
    '��������� � ������� DictIsklucheniyaToruda: ���� - ��� 1� �� ����� "����������(toruda<=1000), ������� - ����� ������
    Set DictIsklucheniyaToruda = CreateObject("Scripting.Dictionary")
    For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("����������(toruda<=1000)").Cells.SpecialCells(xlCellTypeLastCell).Row
        If Trim(Workbooks(ProschetWorkbookName).Worksheets("����������(toruda<=1000)").Cells(i, 1).Value) <> vbNullString Then
            DictIsklucheniyaToruda.Add Trim(Workbooks(ProschetWorkbookName).Worksheets("����������(toruda<=1000)").Cells(i, 1).Value), i
        End If
    Next i

    '��������� ������ ���� �� ��������
    '���� ���� ���� "��������" - �� ��������� � start2
    '���� ��� - ������� ���� "��������" � ��������� � start2
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "��������" Then
            GoTo start2
        End If
    Next wks

    Call CreateSheet("��������")

    Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(1, 1).Value = "������"
    Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(2, 1).Value = "������"
    Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(3, 1).Value = "���������"
    Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(4, 1).Value = "�����-���������"
    Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(5, 1).Value = "������������"

    '��������� � ������� DictLgotniiGorod: ���� - ������������ ������ �� ����� "��������", ������� - ����� ������
start2:
    Set DictLgotniiGorod = CreateObject("Scripting.Dictionary")
    For i = 1 To Workbooks(ProschetWorkbookName).Worksheets("��������").Cells.SpecialCells(xlCellTypeLastCell).Row
        If Trim(Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(i, 1).Value) <> vbNullString Then
            DictLgotniiGorod.Add Trim(Workbooks(ProschetWorkbookName).Worksheets("��������").Cells(i, 1).Value), i
        End If
    Next i

    '��������� ������ ���� �� ��������
    '���� ���� ���� "�����" - �� ��������� � start1
    '���� ��� - ������� ���� "�����" � ��������� � start1
    For Each wks In Workbooks(ProschetWorkbookName).Worksheets
        If wks.name = "�����" Then
            GoTo start1
        End If
    Next wks

    Call CreateSheet("�����")
start1:
    'AkciaArray = ������ ����� "�����" �� A1:F �� ��������� ����������� ������
    AkciaArray = Workbooks(ProschetWorkbookName).Worksheets("�����").Range("A1:F" & _
    Workbooks(ProschetWorkbookName).Worksheets("�����").Cells.SpecialCells(xlCellTypeLastCell).Row).Value

    '��������� � ������� DictAkcia: ���� - ��� 1� �� ����� "�����", ������� - ����� ������
    Set DictAkcia = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(AkciaArray, 1)
        If Trim(AkciaArray(i, 1)) <> vbNullString Then
            DictAkcia.Add Trim(AkciaArray(i, 1)), i
        End If
    Next i

    '��������� �� ��������:
    If Workbooks(ProschetWorkbookName).name = "������� ��� DDI.xlsb" Then
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
        Call ���������_���_������

'        '���� "��������� ���������"
'    ElseIf UserForm1.���������_��������� = True Then
'        Call ���������_���������

        '���� "������ �������"
    ElseIf UserForm1.������_������� = True Then
        Call ������_�������

        '���� "��������� �����"
    ElseIf UserForm1.���������_���������� = True Then
        Call ���������_����������
    
    ElseIf UserForm1.��������_��������� = True Or UserForm1.�������_��������� = True Then
        Call ��������_�������_���������
    
        '���� ������� ���������� ������� ��� DDI.xlsb
    ElseIf Workbooks(ProschetWorkbookName).name = "������� ��� DDI.xlsb" Then
        Call ���������_DDI

        '���� ������ Default
    ElseIf UserForm1.OptionButton1 = True Or UserForm1.OptionButton3 = True Then
        Call ���������_���_������

        '���� ������� "�����" � � ���� "���-�" ����� - ��������� ���� "������������ ���������"
        '� ����������� Main
    ElseIf UserForm1.OptionButton2 = True And Trim(UserForm1.TextBox3) = vbNullString Then
        MsgBox "������������ ���������"
        Call Main

        '���� ������� "�����"
    ElseIf UserForm1.OptionButton2 = True Then
        Call ����������_�����_�����

        '���� ������� "����� � �����" � � ���� "���-�" ����� -��������� ���� "������������ ���������"
        '� ����������� Main
    ElseIf UserForm1.OptionButton3 = True And Trim(UserForm1.TextBox3) = vbNullString Then
        MsgBox "������������ ���������"
        Call Main

        '���� ������� "����� � �����"
'    ElseIf UserForm1.OptionButton3 = True Then
'        Call ���������_�����_�����
    End If
    
End Sub

