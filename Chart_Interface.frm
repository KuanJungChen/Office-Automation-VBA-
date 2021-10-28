VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Chart_Interface 
   Caption         =   "蕉蕉系統-製作圖表"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12096
   OleObjectBlob   =   "Chart_Interface.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Chart_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_costincomeOk_Click()
    '成本收益表
    Sheets("成本收益比較").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("成本收益比較!$A$1:$E$9")
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Range("C:D")
    ActiveChart.FullSeriesCollection(1).XValues = "=成本收益比較!$A$1:$B$9"
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "成本收益表"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "成本收益表"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    '利潤
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("成本收益比較!$A$1:$E$9")
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Range("E1:E9")
    ActiveChart.FullSeriesCollection(1).XValues = "=成本收益比較!$A$1:$B$9"
    ActiveChart.ChartTitle.Select
End Sub

Private Sub Button_incomeOk_Click()
Sheets("圖表-收益").Select
Dim First_I, Last_I As String
Dim colCnt_I, cIdx_I As Integer
First_I = ComboBox_incomeFirstYear.Value & "/" & ComboBox_incomeFirstMonth.Value & "/" & ComboBox_incomeFirstDay.Value
Last_I = ComboBox_incomeLastYear.Value & "/" & ComboBox_incomeLastMonth.Value & "/" & ComboBox_incomeLastDay.Value
colCnt_I = VBA.DateDiff("d", First_I, Last_I) + 2
For cIdx_I = 2 To colCnt_I
    Cells(3 * (cIdx_I - 2) + 2, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 3, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 4, "A").Value = First_I
    Cells(3 * (cIdx_I - 2) + 2, "B").Value = "北蕉"
    Cells(3 * (cIdx_I - 2) + 3, "B").Value = "寶島蕉"
    Cells(3 * (cIdx_I - 2) + 4, "B").Value = "台蕉五號"
    First_I = DateAdd("d", 1, First_I)
Next
Dim colCnt1, beg_data, data_Chart As Integer
Dim dtrange_Chart As Range
Set dtrange_Chart = Sheets("圖表-收益").UsedRange
data_Chart = dtrange_Chart.Rows.Count
'colCnt1 = colCnt_I * 3 + 2 - 1
Dim data_num As Integer
Dim dtRange As Range
Set dtRange = Sheets("收益").UsedRange
data_num = dtRange.Rows.Count
Dim in_data As Integer

For beg_data = 2 To data_Chart
    For in_data = 2 To data_num
        If (Worksheets("圖表-收益").Cells(beg_data, "A").Value = Worksheets("收益").Cells(in_data, "C").Value) And (Worksheets("圖表-收益").Cells(beg_data, "B").Value = Worksheets("收益").Cells(in_data, "G").Value) Then
            Worksheets("圖表-收益").Cells(beg_data, "C").Value = Sheets("收益").Cells(in_data, "H").Value
        End If
    Next
    If (Worksheets("圖表-收益").Cells(beg_data, "C").Value = "") Then
        Worksheets("圖表-收益").Cells(beg_data, "C").Value = 0
    End If
Next

Dim i, b, d, c As Integer
For i = 1 To data_Chart
    Select Case Cells(i, 2).Value
        Case "北蕉"
            b = b + Cells(i, 3).Value
        Case "寶島蕉"
            d = d + Cells(i, 3).Value
        Case "台蕉五號"
            c = c + Cells(i, 3).Value
    End Select
Next

Cells(2, 5).Value = "北蕉"
Cells(3, 5).Value = "寶島蕉"
Cells(4, 5).Value = "台蕉五號"
Cells(2, 6).Value = b
Cells(3, 6).Value = d
Cells(4, 6).Value = c
'Range("E2:F4").Select
Dim chart_dtRange1 As Range
Set chart_dtRange1 = Range("E2:F4")
    ActiveSheet.Shapes.AddChart2(262, 5).Select
    ActiveChart.SetSourceData Source:=chart_dtRange1
    'ActiveChart.SetSourceData Source:=Range("圖表-收益!$E$2:$F$4")
    'ActiveChart.ApplyLayout (6)
End Sub

'製作圖表表單初始化設定
Private Sub UserForm_Initialize()
    '年份下拉式選單
    Dim chartYear As Integer
    For chartYear = 2016 To 2025
        ComboBox_costFirstYear.AddItem chartYear
        ComboBox_costLastYear.AddItem chartYear
        ComboBox_incomeFirstYear.AddItem chartYear
        ComboBox_incomeLastYear.AddItem chartYear
    Next
    
    '月份下拉式選單
    Dim chartMonth As Integer
    For chartMonth = 1 To 12
        ComboBox_costFirstMonth.AddItem chartMonth
        ComboBox_costLastMonth.AddItem chartMonth
        ComboBox_incomeFirstMonth.AddItem chartMonth
        ComboBox_incomeLastMonth.AddItem chartMonth
    Next
    
    '製作成本圖表類型下拉式表單增加選項
    ComboBox_costType.AddItem "生產成本"
    ComboBox_costType.AddItem "間接成本"
    ComboBox_costType.AddItem "固定成本"

    '製作收益圖表類型下拉式表單增加選項
    ComboBox_incomeType.AddItem "品種"
    
End Sub
Private Sub ComboBox_costFirstMonth_Change()
    '日期下拉式選單
    Dim chartDay As Integer
    ComboBox_costFirstDay.Clear
    If ComboBox_costFirstYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
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
    '日期下拉式選單
    Dim chartDay As Integer
    ComboBox_costLastDay.Clear
    If ComboBox_costLastYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
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
    '日期下拉式選單
    Dim chartDay As Integer
    ComboBox_incomeFirstDay.Clear
    If ComboBox_incomeFirstYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
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
    '日期下拉式選單
    Dim chartDay As Integer
    ComboBox_incomeLastDay.Clear
    If ComboBox_incomeLastYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
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

'限制製作成本圖表期間年份1只能輸入數字
Private Sub ComboBox_costFirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本圖表期間月份1只能輸入數字
Private Sub ComboBox_costFirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本圖表期間日期1只能輸入數字
Private Sub ComboBox_costFirstDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本圖表期間年份2只能輸入數字
Private Sub ComboBox_costLastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本圖表期間月份2只能輸入數字
Private Sub ComboBox_costLastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本圖表期間日期2只能輸入數字
Private Sub ComboBox_costLastDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間年份1只能輸入數字
Private Sub ComboBox_incomeFirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間月份1只能輸入數字
Private Sub ComboBox_incomeFirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間日期1只能輸入數字
Private Sub ComboBox_incomeFirstDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間年份2只能輸入數字
Private Sub TxtBox_incomeLastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間月份2只能輸入數字
Private Sub ComboBox_incomeLastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作收益圖表期間日期2只能輸入數字
Private Sub ComboBox_incomeLastDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本收益比較圖表期間年份1只能輸入數字
Private Sub ComboBox_FirstYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本收益比較圖表期間月份1只能輸入數字
Private Sub ComboBox_FirstMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本收益比較圖表期間年份2只能輸入數字
Private Sub ComboBox_LastYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制製作成本收益比較圖表期間月份2只能輸入數字
Private Sub ComboBox_LastMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'製作收益圖表項目下拉式表單增加選項
Private Sub ComboBox_incomeType_Change()
    Select Case ComboBox_incomeType.Text
        Case Is = "品種"
            ComboBox_incomeItem.AddItem "全選"
            'ComboBox_incomeItem.AddItem "北蕉"
            'ComboBox_incomeItem.AddItem "寶島蕉"
            'ComboBox_incomeItem.AddItem "台蕉5號"
    End Select
End Sub
'回登入表單
Private Sub Button_chartBacklogin_Click()
    Chart_Interface.Hide '製作圖表表單隱藏
    Login_Interface.Show '開啟登入表單
End Sub
'製作成本圖表
Private Sub Button_costOk_Click()
    '製作成本圖表第1個時間確認都有填
    If ComboBox_costFirstYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
    ElseIf ComboBox_costFirstYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 31 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 30 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 2
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 29 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在製作圖表表單
        End Select
    Else
        Select Case ComboBox_costFirstMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 31 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 30 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 2
            If ComboBox_costFirstDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costFirstDay.Value > 28 Or ComboBox_costFirstDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在製作圖表表單
        End Select
    End If
    '製作成本圖表第2個時間確認都有填
    If ComboBox_costLastYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在製作圖表表單
    ElseIf ComboBox_costLastYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 31 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 30 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 2
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 29 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在製作圖表表單
        End Select
    Else
        Select Case ComboBox_costLastMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 31 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 30 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case 2
            If ComboBox_costLastDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            ElseIf ComboBox_costLastDay.Value > 28 Or ComboBox_costLastDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在製作圖表表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在製作圖表表單
        End Select
    End If
    If ComboBox_costType.Value = Empty Then
        MsgBox "請輸入類型" '提醒使用者輸入類型
        Exit Sub '強制停留在製作圖表表單
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
        If Sheets(sheetIdx).Name = "圖表" Then
            Application.DisplayAlerts = False
            Sheets("圖表").Delete
            Application.DisplayAlerts = True
        End If
        sheetCnt = Sheets.Count
    Next
    
    d1 = DateSerial(ComboBox_costFirstYear.Value, ComboBox_costFirstMonth.Value, ComboBox_costFirstDay.Value)
    d2 = DateSerial(ComboBox_costLastYear.Value, ComboBox_costLastMonth.Value, ComboBox_costLastDay.Value)
    costType = ComboBox_costType.Value
    
    Sheets("成本").Select
    If ActiveSheet.AutoFilterMode = True Then '如果已經有篩選了
        ActiveSheet.AutoFilterMode = False '把篩選關掉
    End If
    
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$11").AutoFilter Field:=3, Criteria1:= _
        ">=" & CStr(d1), Operator:=xlAnd, Criteria2:="<=" & CStr(d2)
    ActiveSheet.Range("$A$1:$J$21").AutoFilter Field:=7, Criteria1:=costType
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Worksheets(ActiveWorkbook.Worksheets.Count).Name = "篩選後的值"
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Worksheets(ActiveWorkbook.Worksheets.Count).Name = "圖表"
    Sheets("篩選後的值").Select
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Sheets("圖表").Range("A1").Value = "項目"
    Sheets("圖表").Range("B1").Value = "金額"
    
    
    Dim rCnt_R, rIdx_R As Integer
    Dim rCnt_Chart, rIdx_Chart As Integer
    Dim dtrange_R As Range
    Set dtrange_R = Sheets("篩選後的值").UsedRange
    rCnt_R = dtrange_R.Rows.Count

    Dim dtrange_Chart As Range
    Set dtrange_Chart = Sheets("圖表").UsedRange
    rCnt_Chart = dtrange_Chart.Rows.Count

    Sheets("圖表").Select

    For rIdx_R = 2 To rCnt_R
        Set dtrange_Chart = Sheets("圖表").UsedRange
        rCnt_Chart = dtrange_Chart.Rows.Count
        For rIdx_Chart = 2 To rCnt_Chart + 1
            If (Sheets("篩選後的值").Cells(rIdx_R, "E").Value <> Sheets("圖表").Cells(rIdx_Chart, "A").Value) Then
                If (Sheets("圖表").Cells(rIdx_Chart, "A").Value = "") Then
                    Sheets("圖表").Cells(rIdx_Chart, "A").Value = Sheets("篩選後的值").Cells(rIdx_R, "E").Value
                    Sheets("圖表").Cells(rIdx_Chart, "B").Value = Sheets("篩選後的值").Cells(rIdx_R, "H").Value
                End If
            ElseIf (Sheets("篩選後的值").Cells(rIdx_R, "E").Value = Sheets("圖表").Cells(rIdx_Chart, "A").Value) Then
                Sheets("圖表").Cells(rIdx_Chart, "B").Value = Sheets("圖表").Cells(rIdx_Chart, "B").Value + Sheets("篩選後的值").Cells(rIdx_R, "H").Value
                rIdx_Chart = rCnt_Chart + 1
            End If
        Next
    Next
    
    Application.DisplayAlerts = False
    Sheets("篩選後的值").Delete
    Application.DisplayAlerts = True
    
    Sheets("成本").Select
    If ActiveSheet.AutoFilterMode = True Then '如果已經有篩選了
        ActiveSheet.AutoFilterMode = False '把篩選關掉
    End If
    
    Sheets("圖表").Activate
    Dim chart_dtRange As Range '宣告dtRange為範圍變數
    Set chart_dtRange = ActiveSheet.UsedRange '範圍dtRange設定為已使用區域
    ActiveSheet.Shapes.AddChart2(201, 5).Select '新增圖表-圓餅圖
    ActiveChart.SetSourceData Source:=chart_dtRange '圖表來源
End Sub

