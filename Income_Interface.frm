VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Income_Interface 
   Caption         =   "蕉蕉系統-登記收益"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5820
   OleObjectBlob   =   "Income_Interface.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Income_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'登記收益表單初始化設定
Private Sub UserForm_Initialize()
    '年份下拉式選單
    Dim incomeYear As Integer
    For incomeYear = 2016 To 2025
        ComboBox_incomeYear.AddItem incomeYear
    Next
    
    '月份下拉式選單
    Dim incomeMonth As Integer
    For incomeMonth = 1 To 12
        ComboBox_incomeMonth.AddItem incomeMonth
    Next
    
    '香蕉品種下拉式選單
    ComboBox_incomeType.AddItem "北蕉"
    ComboBox_incomeType.AddItem "寶島蕉"
    ComboBox_incomeType.AddItem "台蕉五號"
End Sub
Private Sub ComboBox_incomeMonth_Change()
    '日期下拉式選單
    Dim incomeDay As Integer
    ComboBox_incomeDay.Clear
    If ComboBox_incomeYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登記收益表單
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
'限制登記收益日期年只能輸入數字
Private Sub TxtBox_incomeYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記收益日期月只能輸入數字
Private Sub TxtBox_incomeMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記收益日期日只能輸入數字
Private Sub TxtBox_incomeDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記收益數量只能輸入數字
Private Sub TxtBox_incomeNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記收益金額只能輸入數字
Private Sub TxtBox_incomePrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'回登入表單
Private Sub Button_incomeBacklogin_Click()
    Income_Interface.Hide '登記收益表單隱藏
    Login_Interface.Show '開啟登入表單
End Sub
'清空所有
Private Sub Button_incomeClear_Click()
    TxtBox_incomeNum.Text = ""
    TxtBox_incomeUnit.Text = ""
    TxtBox_incomeBuyer.Text = ""
    TxtBox_incomePrice.Text = ""
    TxtBox_incomeReceipt.Text = ""
    TxtBox_incomeRemarks.Text = ""
End Sub
'輸入至excel工作表
Private Sub Button_incomeOk_Click()

    If ComboBox_incomeYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登記收益表單
    ElseIf ComboBox_incomeYear.Value Mod 4 = 0 Then
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 31 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 30 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case 2
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 29 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登記收益表單
        End Select
    Else
        Select Case ComboBox_incomeMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 31 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 30 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case 2
            If ComboBox_incomeDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            ElseIf ComboBox_incomeDay.Value > 28 Or ComboBox_incomeDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記收益表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登記收益表單
        End Select
    End If
    If ComboBox_incomeType.Value = Empty Then
        MsgBox "請輸入成本類型" '提醒使用者輸入成本類型
        Exit Sub '強制停留在登記收益表單
    End If
    If TxtBox_incomePrice.Text = Empty Then
        MsgBox "請輸入金額" '提醒使用者輸入金額
        Exit Sub '強制停留在登記收益表單
    End If

'寫入表格
Dim rCnt_income As Integer
Dim dtRange_income As Range
Set dtRange_income = Sheets("收益").UsedRange
Sheets("收益").Select
rCnt_income = dtRange_income.Rows.Count + 1
'登入日期欄位
Cells(rCnt_income, "A").Value = Login_Interface.ComboBox_loginYear.Value & "/" & Login_Interface.ComboBox_loginMonth.Value & "/" & Login_Interface.ComboBox_loginDay.Value
'登記人欄位
Cells(rCnt_income, "B").Value = Login_Interface.TxtBox_loginName.Text
'日期欄位
Cells(rCnt_income, "C").Value = ComboBox_incomeYear.Value & "/" & ComboBox_incomeMonth.Value & "/" & ComboBox_incomeDay.Value
'年欄位
Cells(rCnt_income, "D").Value = ComboBox_incomeYear.Value
'月欄位
Cells(rCnt_income, "E").Value = ComboBox_incomeMonth.Value
'日欄位
Cells(rCnt_income, "F").Value = ComboBox_incomeDay.Value
'品種欄位
Cells(rCnt_income, "G").Value = ComboBox_incomeType.Value
'數量欄位
Cells(rCnt_income, "H").Value = TxtBox_incomeNum.Text
'單位欄位
Cells(rCnt_income, "I").Value = TxtBox_incomeUnit.Text
'金額欄位
Cells(rCnt_income, "J").Value = TxtBox_incomePrice.Text
'交易欄位
Cells(rCnt_income, "K").Value = TxtBox_incomeBuyer.Text
'統一編號欄位
Cells(rCnt_income, "L").Value = TxtBox_incomeReceipt.Text
'備註欄位
Cells(rCnt_income, "M").Value = TxtBox_incomeRemarks.Text

'調整欄寬
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

'寫入成本收益比較表
Dim rCnt_RC, rIdx_RC As Integer
Dim dtrange_RC As Range
Set dtrange_RC = Sheets("成本收益比較").UsedRange
Sheets("成本收益比較").Select
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
'計算利潤
Set dtrange_RC = Sheets("成本收益比較").UsedRange
Sheets("成本收益比較").Select
rCnt_RC = dtrange_RC.Rows.Count
For rIdx_RC = 2 To rCnt_RC
    Cells(rIdx_RC, "E").Value = Cells(rIdx_RC, "D").Value - Cells(rIdx_RC, "C").Value
Next
End Sub
