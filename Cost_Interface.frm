VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cost_Interface 
   Caption         =   "蕉蕉系統-登記成本"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5820
   OleObjectBlob   =   "Cost_Interface.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Cost_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'登記成本表單初始化設定
Private Sub UserForm_Initialize()
    '年份下拉式選單
    Dim costYear As Integer
    For costYear = 2016 To 2025
        ComboBox_costYear.AddItem costYear
    Next
    
    '月份下拉式選單
    Dim costMonth As Integer
    For costMonth = 1 To 12
        ComboBox_costMonth.AddItem costMonth
    Next
    
    ComboBox_costType.AddItem ("生產成本")
    ComboBox_costType.AddItem ("間接成本")
    ComboBox_costType.AddItem ("固定成本")
End Sub
'登記成本項目
Private Sub ComboBox_costType_Change()
    Select Case ComboBox_costType.Value
    Case Is = "生產成本"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "種子"
        ComboBox_costItem.AddItem "肥料"
        ComboBox_costItem.AddItem "農藥"
        ComboBox_costItem.AddItem "箱子"
        ComboBox_costItem.AddItem "其他"
    Case Is = "間接成本"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "人力費"
        ComboBox_costItem.AddItem "機器設備費"
        ComboBox_costItem.AddItem "水電費"
        ComboBox_costItem.AddItem "其他"
    Case Is = "固定成本"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "土地租金"
        ComboBox_costItem.AddItem "其他"
    Case Is = "總成本"
        ComboBox_costItem.Clear
        ComboBox_costItem.AddItem "生產成本"
        ComboBox_costItem.AddItem "間接成本"
        ComboBox_costItem.AddItem "固定成本"
        ComboBox_costItem.AddItem "所有成本"
    End Select
End Sub
Private Sub ComboBox_costMonth_Change()
    '日期下拉式選單
    Dim costDay As Integer
    ComboBox_costDay.Clear
    If ComboBox_costYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登記成本表單
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
'限制登記成本日期年只能輸入數字
Private Sub TxtBox_costYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記成本日期月只能輸入數字
Private Sub TxtBox_costMonth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記成本日期日只能輸入數字
Private Sub TxtBox_costDay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記成本數量只能輸入數字
Private Sub TxtBox_costNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'限制登記成本金額只能輸入數字
Private Sub TxtBox_costPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) = False Then
       KeyAscii = 0
    End If
End Sub
'回登入表單
Private Sub Button_costBacklogin_Click()
    Cost_Interface.Hide '登記成本表單隱藏
    Login_Interface.Show '開啟登入表單
End Sub
'清空所有
Private Sub Button_costClear_Click()
    TxtBox_costNum.Text = ""
    TxtBox_costUnit.Text = ""
    TxtBox_costPrice.Text = ""
    TxtBox_costReceipt.Text = ""
    TxtBox_costRemarks.Text = ""
End Sub
'登記成本確認
Private Sub Button_costOk_Click()
    If ComboBox_costYear.Value = Empty Then
        MsgBox "請輸入正確的年份" '提醒使用者輸入正確的年份
        Exit Sub '強制停留在登記成本表單
    ElseIf ComboBox_costYear.Value Mod 4 = 0 Then
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 31 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 30 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case 2
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 29 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登記成本表單
        End Select
    Else
        Select Case ComboBox_costMonth.Value
        Case 1, 3, 5, 7, 8, 10, 12
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 31 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case 4, 6, 9, 11
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 30 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case 2
            If ComboBox_costDay.Value = Empty Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            ElseIf ComboBox_costDay.Value > 28 Or ComboBox_costDay.Value < 0 Then
                MsgBox "請輸入正確的日期" '提醒使用者輸入正確的日期
                Exit Sub '強制停留在登記成本表單
            End If
        Case Else
            MsgBox "請輸入正確的月份" '提醒使用者輸入正確的月份
            Exit Sub '強制停留在登記成本表單
        End Select
    End If
    If ComboBox_costType.Value = Empty Then
        MsgBox "請輸入成本類型" '提醒使用者輸入成本類型
        Exit Sub '強制停留在登記成本表單
    End If
    If ComboBox_costItem.Value = Empty Then
        MsgBox "請輸入成本項目" '提醒使用者輸入成本項目
        Exit Sub '強制停留在登記成本表單
    End If
    If TxtBox_costPrice.Text = Empty Then
        MsgBox "請輸入金額" '提醒使用者輸入金額
        Exit Sub '強制停留在登記成本表單
    End If
    
    Sheets("成本").Select '鎖定第一張工作表(成本)
    Dim cCnt As Integer '下一筆資料輸入後所會在的列數
    'MsgBox (ActiveSheet.UsedRange.Rows.Count)
    cCnt = ActiveSheet.UsedRange.Rows.Count + 1 '計算目前有幾列並+1
    Cells(cCnt, "A").Value = Login_Interface.ComboBox_loginYear.Value & "/" & Login_Interface.ComboBox_loginMonth.Value & "/" & Login_Interface.ComboBox_loginDay.Value '將發生日期回傳到日期(yyyy/mm/dd)
    Cells(cCnt, "B").Value = Login_Interface.TxtBox_loginName.Text
    Cells(cCnt, "C").Value = ComboBox_costYear.Value & "/" & ComboBox_costMonth.Value & "/" & ComboBox_costDay.Value '將發生日期回傳到日期(yyyy/mm/dd)
    Cells(cCnt, "D").Value = ComboBox_costYear.Value
    Cells(cCnt, "E").Value = ComboBox_costMonth.Value
    Cells(cCnt, "F").Value = ComboBox_costDay.Value
    Cells(cCnt, "G").Value = ComboBox_costType.Value '要在Initialize改,用select case 或 if
    Cells(cCnt, "H").Value = ComboBox_costItem.Value '要在Initialize改,用select case 或 if
    Cells(cCnt, "I").Value = TxtBox_costNum.Text '將數量回傳
    Cells(cCnt, "J").Value = TxtBox_costUnit.Text '將單位回傳
    Cells(cCnt, "K").Value = TxtBox_costPrice.Text '將金額回傳
    Cells(cCnt, "L").Value = TxtBox_costReceipt.Text '統一編號回傳
    Cells(cCnt, "M").Value = TxtBox_costRemarks.Text '將備註回傳
    '登記人的話必須在按下成本這個選項時就進行回傳，因此必須由登入畫面進行回傳
    
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
'計算利潤
Set dtrange_RC = Sheets("成本收益比較").UsedRange
Sheets("成本收益比較").Select
rCnt_RC = dtrange_RC.Rows.Count
For rIdx_RC = 2 To rCnt_RC
    Cells(rIdx_RC, "E").Value = Cells(rIdx_RC, "D").Value - Cells(rIdx_RC, "C").Value
Next
End Sub

