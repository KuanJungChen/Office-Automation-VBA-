VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login_Interface 
   Caption         =   "�����t��-�n�J����"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6360
   OleObjectBlob   =   "Login_Interface.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Login_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�n�J����l�Ƴ]�w
Private Sub UserForm_Initialize()
    '�۰ʶפJ�n�J�ɶ�
    ComboBox_loginYear.Value = year(Date) '�۰ʶפJ���Ѫ��~
    ComboBox_loginMonth.Value = month(Date) '�۰ʶפJ���Ѫ���
    ComboBox_loginDay.Value = day(Date) '�۰ʶפJ���Ѫ���
    
    '�~���U�Ԧ����
    Dim loginYear As Integer
    For loginYear = 2016 To 2025
        ComboBox_loginYear.AddItem loginYear
    Next
    
    '����U�Ԧ����
    Dim loginMonth As Integer
    For loginMonth = 1 To 12
        ComboBox_loginMonth.AddItem loginMonth
    Next
    
End Sub
'����U�Ԧ����
Private Sub ComboBox_loginMonth_Change()
    Dim loginDay As Integer
    ComboBox_loginDay.Clear
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�J���
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Text
        Case 1, 3, 5, 7, 8, 10, 12
            For loginDay = 1 To 31
                ComboBox_loginDay.AddItem loginDay
            Next
        Case 4, 6, 9, 11
            For loginDay = 1 To 30
                ComboBox_loginDay.AddItem loginDay
            Next
        Case 2
            For loginDay = 1 To 29
                ComboBox_loginDay.AddItem loginDay
            Next
        End Select
    Else
        Select Case ComboBox_loginMonth.Text
        Case 1, 3, 5, 7, 8, 10, 12
            For loginDay = 1 To 31
                ComboBox_loginDay.AddItem loginDay
            Next
        Case 4, 6, 9, 11
            For loginDay = 1 To 30
                ComboBox_loginDay.AddItem loginDay
            Next
        Case 2
            For loginDay = 1 To 28
                ComboBox_loginDay.AddItem loginDay
            Next
        End Select
    End If
End Sub
'����n�J����~�u���J�Ʀr
Private Sub ComboBox_loginYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�J�����u���J�Ʀr
Private Sub ComboBox_loginMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����n�J�����u���J�Ʀr
Private Sub ComboBox_loginDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'�n�O����
Private Sub Button_Cost_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "�п�J�n�O�H" '�����ϥΪ̿�J�n�O�H
        Exit Sub '�j��d�b�n�J���
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�J���
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    End If
    Login_Interface.Hide '�n�J�������
    Cost_Interface.Show '�}�ҵn�O�������
End Sub
'�n�O���q
Private Sub Button_Income_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "�п�J�n�O�H" '�����ϥΪ̿�J�n�O�H
        Exit Sub '�j��d�b�n�J���
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�J���
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    End If
    Login_Interface.Hide '�n�J�������
    Income_Interface.Show '�}�ҵn�O���q���
End Sub
'�s�@�Ϫ�
Private Sub Button_makeChart_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "�п�J�n�O�H" '�����ϥΪ̿�J�n�O�H
        Exit Sub '�j��d�b�n�J���
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�n�J���
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�n�J���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�n�J���
        End Select
    End If
    
    Login_Interface.Hide '�n�J�������
    Chart_Interface.Show '�}�һs�@�Ϫ���
End Sub

