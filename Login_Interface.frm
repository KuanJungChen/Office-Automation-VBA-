VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login_Interface 
   Caption         =   "蕉蕉系統-登入介面"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6360
   OleObjectBlob   =   "Login_Interface.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Login_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'登入表單初始化設定
Private Sub UserForm_Initialize()
    '自動匯入登入時間
    ComboBox_loginYear.Value = year(Date) '自動匯入今天的年
    ComboBox_loginMonth.Value = month(Date) '自動匯入今天的月
    ComboBox_loginDay.Value = day(Date) '自動匯入今天的日
    
    '年份下拉式選單
    Dim loginYear As Integer
    For loginYear = 2016 To 2025
        ComboBox_loginYear.AddItem loginYear
    Next
    
    '月份下拉式選單
    Dim loginMonth As Integer
    For loginMonth = 1 To 12
        ComboBox_loginMonth.AddItem loginMonth
    Next
    
End Sub
'日期下拉式選單
Private Sub ComboBox_loginMonth_Change()
    Dim loginDay As Integer
    ComboBox_loginDay.Clear
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登入表單
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
'限制登入日期年只能輸入數字
Private Sub ComboBox_loginYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登入日期月只能輸入數字
Private Sub ComboBox_loginMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登入日期日只能輸入數字
Private Sub ComboBox_loginDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'登記成本
Private Sub Button_Cost_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "請輸入登記人" '提醒使用者輸入登記人
        Exit Sub '強制停留在登入表單
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登入表單
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    End If
    Login_Interface.Hide '登入表單隱藏
    Cost_Interface.Show '開啟登記成本表單
End Sub
'登記收益
Private Sub Button_Income_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "請輸入登記人" '提醒使用者輸入登記人
        Exit Sub '強制停留在登入表單
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登入表單
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    End If
    Login_Interface.Hide '登入表單隱藏
    Income_Interface.Show '開啟登記收益表單
End Sub
'製作圖表
Private Sub Button_makeChart_Click()
    If TxtBox_loginName.Text = Empty Then
        MsgBox "請輸入登記人" '提醒使用者輸入登記人
        Exit Sub '強制停留在登入表單
    End If
    
    If ComboBox_loginYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登入表單
    ElseIf ComboBox_loginYear.Value Mod 4 = 0 Then
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 29 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    Else
        Select Case ComboBox_loginMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 31 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 30 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case 2
            If ComboBox_loginDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            ElseIf ComboBox_loginDay.Value > 28 Or ComboBox_loginDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登入表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登入表單
        End Select
    End If
    
    Login_Interface.Hide '登入表單隱藏
    Chart_Interface.Show '開啟製作圖表表單
End Sub

