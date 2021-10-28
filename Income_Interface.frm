VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Income_Interface 
   Caption         =   "�����t��-�n�O���q"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5820
   OleObjectBlob   =   "Income_Interface.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Income_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�n�O���q����l�Ƴ]�w
Private Sub UserForm_Initialize()
    '�~���U�Ԧ����
    Dim incomeYear As Integer
    For incomeYear = 2016 To 2025
        ComboBox_incomeYear.AddItem incomeYear
    Next
    
    '����U�Ԧ����
    Dim incomeMonth As Integer
    For incomeMonth = 1 To 12
        ComboBox_incomeMonth.AddItem incomeMonth
    Next
    
    '�����~�ؤU�Ԧ����
    ComboBox_incomeType.AddItem "�_��"
    ComboBox_incomeType.AddItem "�_�q��"
    ComboBox_incomeType.AddItem "�x������"
End Sub
Private Sub ComboBox_incomeMonth_Change()
    '����U�Ԧ����
    Dim incomeDay As Integer
    ComboBox_incomeDay.Clear
    If ComboBox_incomeYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�O���q���
    ElseIf ComboBox_incomeYear.Value Mod 4 = 0 Then
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For incomeDay = 1 To 31
                ComboBox_incomeDay.AddItem incomeDay
            Next
        Case 4, 6, 9, 11
            For incomeDay = 1 To 30
                ComboBox_incomeDay.AddItem incomeDay
            Next
        Case 2
            For incomeDay = 1 To 29
                ComboBox_incomeDay.AddItem incomeDay
            Next
        End Select
    Else
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For incomeDay = 1 To 31
                ComboBox_incomeDay.AddItem incomeDay
            Next
        Case 4, 6, 9, 11
            For incomeDay = 1 To 30
                ComboBox_incomeDay.AddItem incomeDay
            Next
        Case 2
            For incomeDay = 1 To 28
                ComboBox_incomeDay.AddItem incomeDay
            Next
        End Select
    End If
End Sub
'����n�O���q����~�u���J�Ʀr
Private Sub TxtBox_incomeYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���q�����u���J�Ʀr
Private Sub TxtBox_incomeMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���q�����u���J�Ʀr
Private Sub TxtBox_incomeDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���q�ƶq�u���J�Ʀr
Private Sub TxtBox_incomeNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�O���q���B�u���J�Ʀr
Private Sub TxtBox_incomePrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'�^�n�J���
Private Sub Button_incomeBacklogin_Click()
    Income_Interface.Hide '�n�O���q�������
    Login_Interface.Show '�}�ҵn�J���
End Sub
'�M�ũҦ�
Private Sub Button_incomeClear_Click()
    TxtBox_incomeNum.Text = ""
    TxtBox_incomeUnit.Text = ""
    TxtBox_incomeBuyer.Text = ""
    TxtBox_incomePrice.Text = ""
    TxtBox_incomeReceipt.Text = ""
    TxtBox_incomeRemarks.Text = ""
End Sub
'��J��excel�u�@��
Private Sub Button_incomeOk_Click()

    If ComboBox_incomeYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�O���q���
    ElseIf ComboBox_incomeYear.Value Mod 4 = 0 Then
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 31 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case 4, 6, 9, 11
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 30 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case 2
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 29 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�O���q���
        End Select
    Else
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 31 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case 4, 6, 9, 11
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 30 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case 2
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            ElseIf ComboBox_incomeDay.Value > 28 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�O���q���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�O���q���
        End Select
    End If
    If ComboBox_incomeType.Value = Empty Then
        MsgBox "�п�J��������" '�����ϥΪ̿�J��������
        Exit Sub '�j��d�b�n�O���q���
    End If
    If TxtBox_incomePrice.Text = Empty Then
        MsgBox "�п�J���B" '�����ϥΪ̿�J���B
        Exit Sub '�j��d�b�n�O���q���
    End If

'�g�J���
Dim rCnt_income As Integer
Dim dtRange_income As Range
Set dtRange_income = Sheets("���q").UsedRange
Sheets("���q").Select
rCnt_income = dtRange_income.Rows.Count + 1
'�n�J������
Cells(rCnt_income, "A").Value = Login_Interface.ComboBox_loginYear.Value & "/" & Login_Interface.ComboBox_loginMonth.Value & "/" & Login_Interface.ComboBox_loginDay.Value
'�n�O�H���
Cells(rCnt_income, "B").Value = Login_Interface.TxtBox_loginName.Text
'������
Cells(rCnt_income, "C").Value = ComboBox_incomeYear.Value & "/" & ComboBox_incomeMonth.Value & "/" & ComboBox_incomeDay.Value
'�~���
Cells(rCnt_income, "D").Value = ComboBox_incomeYear.Value
'�����
Cells(rCnt_income, "E").Value = ComboBox_incomeMonth.Value
'�����
Cells(rCnt_income, "F").Value = ComboBox_incomeDay.Value
'�~�����
Cells(rCnt_income, "G").Value = ComboBox_incomeType.Value
'�ƶq���
Cells(rCnt_income, "H").Value = TxtBox_incomeNum.Text
'������
Cells(rCnt_income, "I").Value = TxtBox_incomeUnit.Text
'���B���
Cells(rCnt_income, "J").Value = TxtBox_incomePrice.Text
'������
Cells(rCnt_income, "K").Value = TxtBox_incomeBuyer.Text
'�Τ@�s�����
Cells(rCnt_income, "L").Value = TxtBox_incomeReceipt.Text
'�Ƶ����
Cells(rCnt_income, "M").Value = TxtBox_incomeRemarks.Text

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
    If (CInt(Cells(rIdx_RC, "A").Value) <> CInt(ComboBox_incomeYear.Value)) Then
        If (Cells(rIdx_RC, "A").Value = "") Then
            Cells(rIdx_RC, "A").Value = ComboBox_incomeYear.Value
            Cells(rIdx_RC, "B").Value = ComboBox_incomeMonth.Value
            Cells(rIdx_RC, "D").Value = TxtBox_incomePrice.Text
        End If
    ElseIf (CInt(Cells(rIdx_RC, "A").Value) = CInt(ComboBox_incomeYear.Value)) Then
        If (CInt(Cells(rIdx_RC, "B").Value) <> CInt(ComboBox_incomeMonth.Value)) Then
            If (Cells(rIdx_RC, "B").Value = "") Then
                Cells(rIdx_RC, "A").Value = ComboBox_incomeYear.Value
                Cells(rIdx_RC, "B").Value = ComboBox_incomeMonth.Value
                Cells(rIdx_RC, "D").Value = TxtBox_incomePrice.Text
            End If
        ElseIf (CInt(Cells(rIdx_RC, "B").Value) = CInt(ComboBox_incomeMonth.Value)) Then
            Cells(rIdx_RC, "D").Value = Cells(rIdx_RC, "D").Value + TxtBox_incomePrice.Text
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
