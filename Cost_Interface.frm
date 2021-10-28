VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cost_Interface 
   Caption         =   "�����t��-�n�O����"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5820
   OleObjectBlob   =   "Cost_Interface.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Cost_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�n�O��������l�Ƴ]�w
Private Sub UserForm_Initialize()
    '�~���U�Ԧ����
    Dim costYear As Integer
    For costYear = 2016 To 2025
        ComboBox_costYear.AddItem costYear
    Next
    
    '����U�Ԧ����
    Dim costMonth As Integer
    For costMonth = 1 To 12
        ComboBox_costMonth.AddItem costMonth
    Next
    
    ComboBox_costType.AddItem ("�Ͳ�����")
    ComboBox_costType.AddItem ("��������")
    ComboBox_costType.AddItem ("�T�w����")
End Sub
'�n�O��������
Private Sub ComboBox_costType_Change()
    Select Case ComboBox_costType.Value
    Case Is = "�Ͳ�����"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "�ؤl"
        ComboBox_costItem.AddItem "�ή�"
        ComboBox_costItem.AddItem "�A��"
        ComboBox_costItem.AddItem "�c�l"
        ComboBox_costItem.AddItem "��L"
    Case Is = "��������"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "�H�O�O"
        ComboBox_costItem.AddItem "�����]�ƶO"
        ComboBox_costItem.AddItem "���q�O"
        ComboBox_costItem.AddItem "��L"
    Case Is = "�T�w����"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "�g�a����"
        ComboBox_costItem.AddItem "��L"
    Case Is = "�`����"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "�Ͳ�����"
        ComboBox_costItem.AddItem "��������"
        ComboBox_costItem.AddItem "�T�w����"
        ComboBox_costItem.AddItem "�Ҧ�����"
    End Select
End Sub
Private Sub ComboBox_costMonth_Change()
    '����U�Ԧ����
    Dim costDay As Integer
    ComboBox_costDay.Clear
    If ComboBox_costYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�O�������
    ElseIf ComboBox_costYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For costDay = 1 To 31
                ComboBox_costDay.AddItem costDay
            Next
        Case 4, 6, 9, 11
            For costDay = 1 To 30
                ComboBox_costDay.AddItem costDay
            Next
        Case 2
            For costDay = 1 To 29
                ComboBox_costDay.AddItem costDay
            Next
        End Select
    Else
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For costDay = 1 To 31
                ComboBox_costDay.AddItem costDay
            Next
        Case 4, 6, 9, 11
            For costDay = 1 To 30
                ComboBox_costDay.AddItem costDay
            Next
        Case 2
            For costDay = 1 To 28
                ComboBox_costDay.AddItem costDay
            Next
        End Select
    End If
End Sub
'����n�O��������~�u���J�Ʀr
Private Sub TxtBox_costYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���������u���J�Ʀr
Private Sub TxtBox_costMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���������u���J�Ʀr
Private Sub TxtBox_costDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O�����ƶq�u���J�Ʀr
Private Sub TxtBox_costNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O�������B�u���J�Ʀr
Private Sub TxtBox_costPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'�^�n�J���
Private Sub Button_costBacklogin_Click()
    Cost_Interface.Hide '�n�O�����������
    Login_Interface.Show '�}�ҵn�J���
End Sub
'�M�ũҦ�
Private Sub Button_costClear_Click()
    TxtBox_costNum.Text = ""
    TxtBox_costUnit.Text = ""
    TxtBox_costPrice.Text = ""
    TxtBox_costReceipt.Text = ""
    TxtBox_costRemarks.Text = ""
End Sub
'�n�O�����T�{
Private Sub Button_costOk_Click()
    If ComboBox_costYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�O�������
    ElseIf ComboBox_costYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 31 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case 4, 6, 9, 11
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 30 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case 2
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 29 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�O�������
        End Select
    Else
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 31 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case 4, 6, 9, 11
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 30 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case 2
            If ComboBox_costDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            ElseIf ComboBox_costDay.Value > 28 Or ComboBox_costDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O�������
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�O�������
        End Select
    End If
    If ComboBox_costType.Value = Empty Then
        MsgBox "�п�J��������" '�����ϥΪ̿�J��������
        Exit Sub '�j��d�b�n�O�������
    End If
    If ComboBox_costItem.Value = Empty Then
        MsgBox "�п�J��������" '�����ϥΪ̿�J��������
        Exit Sub '�j��d�b�n�O�������
    End If
    If TxtBox_costPrice.Text = Empty Then
        MsgBox "�п�J���B" '�����ϥΪ̿�J���B
        Exit Sub '�j��d�b�n�O�������
    End If
    
    Sheets("����").Select '��w�Ĥ@�i�u�@��(����)
    Dim cCnt As Integer '�U�@����ƿ�J��ҷ|�b���C��
    'MsgBox (ActiveSheet.UsedRange.Rows.Count)
    cCnt = ActiveSheet.UsedRange.Rows.Count + 1 '�p��ثe���X�C��+1
    Cells(cCnt, "A").Value = Login_Interface.ComboBox_loginYear.Value & "/" & Login_Interface.ComboBox_loginMonth.Value & "/" & Login_Interface.ComboBox_loginDay.Value '�N�o�ͤ���^�Ǩ���(yyyy/mm/dd)
    Cells(cCnt, "B").Value = Login_Interface.TxtBox_loginName.Text
    Cells(cCnt, "C").Value = ComboBox_costYear.Value & "/" & ComboBox_costMonth.Value & "/" & ComboBox_costDay.Value '�N�o�ͤ���^�Ǩ���(yyyy/mm/dd)
    Cells(cCnt, "D").Value = ComboBox_costYear.Value
    Cells(cCnt, "E").Value = ComboBox_costMonth.Value
    Cells(cCnt, "F").Value = ComboBox_costDay.Value
    Cells(cCnt, "G").Value = ComboBox_costType.Value '�n�bInitialize��,��select case �� if
    Cells(cCnt, "H").Value = ComboBox_costItem.Value '�n�bInitialize��,��select case �� if
    Cells(cCnt, "I").Value = TxtBox_costNum.Text '�N�ƶq�^��
    Cells(cCnt, "J").Value = TxtBox_costUnit.Text '�N���^��
    Cells(cCnt, "K").Value = TxtBox_costPrice.Text '�N���B�^��
    Cells(cCnt, "L").Value = TxtBox_costReceipt.Text '�Τ@�s���^��
    Cells(cCnt, "M").Value = TxtBox_costRemarks.Text '�N�Ƶ��^��
    '�n�O�H���ܥ����b���U�����o�ӿﶵ�ɴN�i��^�ǡA�]�������ѵn�J�e���i��^��
    
'�վ���e
Range("A1").EntireColumn.AutoFit
Range("B1").EntireColumn.AutoFit
Range("C1").EntireColumn.AutoFit
Range("D1").ColumnWidth = 0
Range("E1").ColumnWidth = 0
Range("F1").ColumnWidth = 0
Range("G1").EntireColumn.AutoFit
Range("H1").EntireColumn.AutoFit
Range("I1").EntireColumn.AutoFit
Range("J1").EntireColumn.AutoFit
Range("K1").EntireColumn.AutoFit
Range("L1").EntireColumn.AutoFit
Range("M1").EntireColumn.AutoFit

'�g�J�������q�����
Dim rCnt_RC, rIdx_RC As Integer
Dim dtrange_RC As Range
Set dtrange_RC = Sheets("�������q���").UsedRange
Sheets("�������q���").Select
rCnt_RC = dtrange_RC.Rows.Count
For rIdx_RC = 2 To rCnt_RC + 1
    If (CInt(Cells(rIdx_RC, "A").Value) <> CInt(ComboBox_costYear.Value)) Then
        If (Cells(rIdx_RC, "A").Value = "") Then
            Cells(rIdx_RC, "A").Value = ComboBox_costYear.Value
            Cells(rIdx_RC, "B").Value = ComboBox_costMonth.Value
            Cells(rIdx_RC, "C").Value = TxtBox_costPrice.Text
        End If
    ElseIf (CInt(Cells(rIdx_RC, "A").Value) = CInt(ComboBox_costYear.Value)) Then
        If (CInt(Cells(rIdx_RC, "B").Value) <> CInt(ComboBox_costMonth.Value)) Then
            If (Cells(rIdx_RC, "B").Value = "") Then
                Cells(rIdx_RC, "A").Value = ComboBox_costYear.Value
                Cells(rIdx_RC, "B").Value = ComboBox_costMonth.Value
                Cells(rIdx_RC, "C").Value = TxtBox_costPrice.Text
            End If
        ElseIf (CInt(Cells(rIdx_RC, "B").Value) = CInt(ComboBox_costMonth.Value)) Then
            Cells(rIdx_RC, "C").Value = Cells(rIdx_RC, "C").Value + TxtBox_costPrice.Text
            rIdx_RC = rCnt_RC + 1
        End If
    End If
Next
'�p��Q��
Set dtrange_RC = Sheets("�������q���").UsedRange
Sheets("�������q���").Select
rCnt_RC = dtrange_RC.Rows.Count
For rIdx_RC = 2 To rCnt_RC
    Cells(rIdx_RC, "E").Value = Cells(rIdx_RC, "D").Value - Cells(rIdx_RC, "C").Value
Next
End Sub

