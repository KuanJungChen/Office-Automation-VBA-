VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Chart_Interface 
   Caption         =   "�����t��-�s�@�Ϫ�"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12096
   OleObjectBlob   =   "Chart_Interface.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Chart_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_costincomeOk_Click()
    '�������q��
    Sheets("�������q���").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("�������q���!$A$1:$E$9")
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Range("C:D")
    ActiveChart.FullSeriesCollection(1).XValues = "=�������q���!$A$1:$B$9"
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "�������q��"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�������q��"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    '�Q��
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("�������q���!$A$1:$E$9")
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Range("E1:E9")
    ActiveChart.FullSeriesCollection(1).XValues = "=�������q���!$A$1:$B$9"
    ActiveChart.ChartTitle.Select
End Sub

Private Sub Button_incomeOk_Click()
Sheets("�Ϫ�-���q").Select
Dim First_I, Last_I As String
Dim colCnt_I, cIdx_I As Integer
First_I = ComboBox_incomeFirstYear.Value & "/" & ComboBox_incomeFirstMonth.Value & "/" & ComboBox_incomeFirstDay.Value
Last_I = ComboBox_incomeLastYear.Value & "/" & ComboBox_incomeLastMonth.Value & "/" & ComboBox_incomeLastDay.Value
colCnt_I = VBA.DateDiff("d", First_I, Last_I) + 2
For cIdx_I = 2 To colCnt_I
    Cells(3 * (cIdx_I - 2) + 2, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 3, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 4, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 2, "B").Value = "�_��"
    Cells(3 * (cIdx_I - 2) + 3, "B").Value = "�_�q��"
    Cells(3 * (cIdx_I - 2) + 4, "B").Value = "�x������"
    First_I = DateAdd("d", 1, First_I)
Next
Dim colCnt1, beg_data, data_Chart As Integer
Dim dtrange_Chart As Range
Set dtrange_Chart = Sheets("�Ϫ�-���q").UsedRange
data_Chart = dtrange_Chart.Rows.Count
'colCnt1 = colCnt_I * 3 + 2 - 1
Dim data_num As Integer
Dim dtRange As Range
Set dtRange = Sheets("���q").UsedRange
data_num = dtRange.Rows.Count
Dim in_data As Integer

For beg_data = 2 To data_Chart
    For in_data = 2 To data_num
        If (Worksheets("�Ϫ�-���q").Cells(beg_data, "A").Value = Worksheets("���q").Cells(in_data, "C").Value) And (Worksheets("�Ϫ�-���q").Cells(beg_data, "B").Value = Worksheets("���q").Cells(in_data, "G").Value) Then
            Worksheets("�Ϫ�-���q").Cells(beg_data, "C").Value = Sheets("���q").Cells(in_data, "H").Value
        End If
    Next
    If (Worksheets("�Ϫ�-���q").Cells(beg_data, "C").Value = "") Then
        Worksheets("�Ϫ�-���q").Cells(beg_data, "C").Value = 0
    End If
Next

Dim i, b, d, c As Integer
For i = 1 To data_Chart
    Select Case Cells(i, 2).Value
        Case "�_��"
            b = b + Cells(i, 3).Value
        Case "�_�q��"
            d = d + Cells(i, 3).Value
        Case "�x������"
            c = c + Cells(i, 3).Value
    End Select
Next

Cells(2, 5).Value = "�_��"
Cells(3, 5).Value = "�_�q��"
Cells(4, 5).Value = "�x������"
Cells(2, 6).Value = b
Cells(3, 6).Value = d
Cells(4, 6).Value = c
'Range("E2:F4").Select
Dim chart_dtRange1 As Range
Set chart_dtRange1 = Range("E2:F4")
    ActiveSheet.Shapes.AddChart2(262, 5).Select
    ActiveChart.SetSourceData Source:=chart_dtRange1
    'ActiveChart.SetSourceData Source:=Range("�Ϫ�-���q!$E$2:$F$4")
    'ActiveChart.ApplyLayout (6)
End Sub

'�s�@�Ϫ����l�Ƴ]�w
Private Sub UserForm_Initialize()
    '�~���U�Ԧ����
    Dim chartYear As Integer
    For chartYear = 2016 To 2025
        ComboBox_costFirstYear.AddItem chartYear
        ComboBox_costLastYear.AddItem chartYear
        ComboBox_incomeFirstYear.AddItem chartYear
        ComboBox_incomeLastYear.AddItem chartYear
    Next
    
    '����U�Ԧ����
    Dim chartMonth As Integer
    For chartMonth = 1 To 12
        ComboBox_costFirstMonth.AddItem chartMonth
        ComboBox_costLastMonth.AddItem chartMonth
        ComboBox_incomeFirstMonth.AddItem chartMonth
        ComboBox_incomeLastMonth.AddItem chartMonth
    Next
    
    '�s�@�����Ϫ������U�Ԧ����W�[�ﶵ
    ComboBox_costType.AddItem "�Ͳ�����"
    ComboBox_costType.AddItem "��������"
    ComboBox_costType.AddItem "�T�w����"

    '�s�@���q�Ϫ������U�Ԧ����W�[�ﶵ
    ComboBox_incomeType.AddItem "�~��"
    
End Sub
Private Sub ComboBox_costFirstMonth_Change()
    '����U�Ԧ����
    Dim chartDay As Integer
    ComboBox_costFirstDay.Clear
    If ComboBox_costFirstYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_costFirstYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_costFirstDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_costFirstDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 29
                ComboBox_costFirstDay.AddItem chartDay
            Next
        End Select
    Else
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_costFirstDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_costFirstDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 28
                ComboBox_costFirstDay.AddItem chartDay
            Next
        End Select
    End If
End Sub
Private Sub ComboBox_costLastMonth_Change()
    '����U�Ԧ����
    Dim chartDay As Integer
    ComboBox_costLastDay.Clear
    If ComboBox_costLastYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_costLastYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_costLastDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_costLastDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 29
                ComboBox_costLastDay.AddItem chartDay
            Next
        End Select
    Else
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_costLastDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_costLastDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 28
                ComboBox_costLastDay.AddItem chartDay
            Next
        End Select
    End If
End Sub
Private Sub ComboBox_incomeFirstMonth_Change()
    '����U�Ԧ����
    Dim chartDay As Integer
    ComboBox_incomeFirstDay.Clear
    If ComboBox_incomeFirstYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_incomeFirstYear.Value Mod 4 = 0 Then
        Select Case ComboBox_incomeFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 29
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        End Select
    Else
        Select Case ComboBox_incomeFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 28
                ComboBox_incomeFirstDay.AddItem chartDay
            Next
        End Select
    End If
End Sub
Private Sub ComboBox_incomeLastMonth_Change()
    '����U�Ԧ����
    Dim chartDay As Integer
    ComboBox_incomeLastDay.Clear
    If ComboBox_incomeLastYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_incomeLastYear.Value Mod 4 = 0 Then
        Select Case ComboBox_incomeLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 29
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        End Select
    Else
        Select Case ComboBox_incomeLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            For chartDay = 1 To 31
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        Case 4, 6, 9, 11
            For chartDay = 1 To 30
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        Case 2
            For chartDay = 1 To 28
                ComboBox_incomeLastDay.AddItem chartDay
            Next
        End Select
    End If
End Sub

'����s�@�����Ϫ�����~��1�u���J�Ʀr
Private Sub ComboBox_costFirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�����Ϫ�������1�u���J�Ʀr
Private Sub ComboBox_costFirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�����Ϫ�������1�u���J�Ʀr
Private Sub ComboBox_costFirstDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�����Ϫ�����~��2�u���J�Ʀr
Private Sub ComboBox_costLastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�����Ϫ�������2�u���J�Ʀr
Private Sub ComboBox_costLastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�����Ϫ�������2�u���J�Ʀr
Private Sub ComboBox_costLastDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�����~��1�u���J�Ʀr
Private Sub ComboBox_incomeFirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�������1�u���J�Ʀr
Private Sub ComboBox_incomeFirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�������1�u���J�Ʀr
Private Sub ComboBox_incomeFirstDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�����~��2�u���J�Ʀr
Private Sub TxtBox_incomeLastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�������2�u���J�Ʀr
Private Sub ComboBox_incomeLastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@���q�Ϫ�������2�u���J�Ʀr
Private Sub ComboBox_incomeLastDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�������q����Ϫ�����~��1�u���J�Ʀr
Private Sub ComboBox_FirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�������q����Ϫ�������1�u���J�Ʀr
Private Sub ComboBox_FirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�������q����Ϫ�����~��2�u���J�Ʀr
Private Sub ComboBox_LastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'����s�@�������q����Ϫ�������2�u���J�Ʀr
Private Sub ComboBox_LastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'�s�@���q�Ϫ��ؤU�Ԧ����W�[�ﶵ
Private Sub ComboBox_incomeType_Change()
    Select Case ComboBox_incomeType.Text
        Case Is = "�~��"
            ComboBox_incomeItem.AddItem "����"
            'ComboBox_incomeItem.AddItem "�_��"
            'ComboBox_incomeItem.AddItem "�_�q��"
            'ComboBox_incomeItem.AddItem "�x��5��"
    End Select
End Sub
'�^�n�J���
Private Sub Button_chartBacklogin_Click()
    Chart_Interface.Hide '�s�@�Ϫ�������
    Login_Interface.Show '�}�ҵn�J���
End Sub
'�s�@�����Ϫ�
Private Sub Button_costOk_Click()
    '�s�@�����Ϫ��1�Ӯɶ��T�{������
    If ComboBox_costFirstYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_costFirstYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 31 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 4, 6, 9, 11
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 30 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 2
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 29 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�s�@�Ϫ���
        End Select
    Else
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 31 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 4, 6, 9, 11
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 30 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 2
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costFirstDay.Value > 28 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�s�@�Ϫ���
        End Select
    End If
    '�s�@�����Ϫ��2�Ӯɶ��T�{������
    If ComboBox_costLastYear.Value = Empty Then
        MsgBox "�п�J���T���~��" '�����ϥΪ̿�J���T���~��
        Exit Sub '�j��d�b�s�@�Ϫ���
    ElseIf ComboBox_costLastYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 31 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 4, 6, 9, 11
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 30 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 2
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 29 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�s�@�Ϫ���
        End Select
    Else
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 31 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 4, 6, 9, 11
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 30 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case 2
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            ElseIf ComboBox_costLastDay.Value > 28 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
                Exit Sub '�j��d�b�s�@�Ϫ���
            End If
        Case Else
            MsgBox "�п�J���T�����" '�����ϥΪ̿�J���T�����
            Exit Sub '�j��d�b�s�@�Ϫ���
        End Select
    End If
    If ComboBox_costType.Value = Empty Then
        MsgBox "�п�J����" '�����ϥΪ̿�J����
        Exit Sub '�j��d�b�s�@�Ϫ���
    End If
    
    Dim d1, d2 As Date
    Dim costType As String
    Dim dtRange As Range
    Dim rowTotal As Integer
    Dim rowIdx As Integer
    Dim sheetCnt As Integer
    sheetCnt = Sheets.Count
    
    Dim sheetIdx As Integer
    For sheetIdx = 1 To sheetCnt
        If Sheets(sheetIdx).Name = "�Ϫ�" Then
            Application.DisplayAlerts = False
            Sheets("�Ϫ�").Delete
            Application.DisplayAlerts = True
        End If
        sheetCnt = Sheets.Count
    Next
    
    d1 = DateSerial(ComboBox_costFirstYear.Value, ComboBox_costFirstMonth.Value, ComboBox_costFirstDay.Value)
    d2 = DateSerial(ComboBox_costLastYear.Value, ComboBox_costLastMonth.Value, ComboBox_costLastDay.Value)
    costType = ComboBox_costType.Value
    
    Sheets("����").Select
    If ActiveSheet.AutoFilterMode = True Then '�p�G�w�g���z��F
        ActiveSheet.AutoFilterMode = False '��z������
    End If
    
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$11").AutoFilter Field:=3, Criteria1:= _
        ">=" & CStr(d1), Operator:=xlAnd, Criteria2:="<=" & CStr(d2)
    ActiveSheet.Range("$A$1:$J$21").AutoFilter Field:=7, Criteria1:=costType
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�z��᪺��"
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�Ϫ�"
    Sheets("�z��᪺��").Select
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Sheets("�Ϫ�").Range("A1").Value = "����"
    Sheets("�Ϫ�").Range("B1").Value = "���B"
    
    
    Dim rCnt_R, rIdx_R As Integer
    Dim rCnt_Chart, rIdx_Chart As Integer
    Dim dtrange_R As Range
    Set dtrange_R = Sheets("�z��᪺��").UsedRange
    rCnt_R = dtrange_R.Rows.Count

    Dim dtrange_Chart As Range
    Set dtrange_Chart = Sheets("�Ϫ�").UsedRange
    rCnt_Chart = dtrange_Chart.Rows.Count

    Sheets("�Ϫ�").Select

    For rIdx_R = 2 To rCnt_R
        Set dtrange_Chart = Sheets("�Ϫ�").UsedRange
        rCnt_Chart = dtrange_Chart.Rows.Count
        For rIdx_Chart = 2 To rCnt_Chart + 1
            If (Sheets("�z��᪺��").Cells(rIdx_R, "E").Value <> Sheets("�Ϫ�").Cells(rIdx_Chart, "A").Value) Then
                If (Sheets("�Ϫ�").Cells(rIdx_Chart, "A").Value = "") Then
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "A").Value = Sheets("�z��᪺��").Cells(rIdx_R, "E").Value
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "B").Value = Sheets("�z��᪺��").Cells(rIdx_R, "H").Value
                End If
            ElseIf (Sheets("�z��᪺��").Cells(rIdx_R, "E").Value = Sheets("�Ϫ�").Cells(rIdx_Chart, "A").Value) Then
                Sheets("�Ϫ�").Cells(rIdx_Chart, "B").Value = Sheets("�Ϫ�").Cells(rIdx_Chart, "B").Value + Sheets("�z��᪺��").Cells(rIdx_R, "H").Value
                rIdx_Chart = rCnt_Chart + 1
            End If
        Next
    Next
    
    Application.DisplayAlerts = False
    Sheets("�z��᪺��").Delete
    Application.DisplayAlerts = True
    
    Sheets("����").Select
    If ActiveSheet.AutoFilterMode = True Then '�p�G�w�g���z��F
        ActiveSheet.AutoFilterMode = False '��z������
    End If
    
    Sheets("�Ϫ�").Activate
    Dim chart_dtRange As Range '�ŧidtRange���d���ܼ�
    Set chart_dtRange = ActiveSheet.UsedRange '�d��dtRange�]�w���w�ϥΰϰ�
    ActiveSheet.Shapes.AddChart2(201, 5).Select '�s�W�Ϫ�-����
    ActiveChart.SetSourceData Source:=chart_dtRange '�Ϫ�ӷ�
End Sub

