VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 這是全活頁簿通用自動日期系統 (統整最終版)
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' 1. 排除掉自動產出的報表分頁，避免衝突
    If Sh.名字 = "HSE_健檢報件表" Or Sh.名字 = "良民證報件表" Or Sh.名字 = "逾期未交清單" Then Exit Sub
    
    ' 2. 如果一次改太多儲存格則跳過 (避免當機)
    If Target.CountLarge > 100 Then Exit Sub
    
    Dim cell As Range
    On Error GoTo ErrorHandler ' 建立防錯機制，避免程式當掉後永久失去監聽
    
    ' 3. --- 良民證模組：監控 R 欄，日期填入 S 欄 ---
    If Not Intersect(Target, Sh.Range("R:R")) Is Nothing Then
        Application.EnableEvents = False ' 暫停事件監聽，防止無限迴圈
        For Each cell In Intersect(Target, Sh.Range("R:R"))
            If cell.Value = "已收" Then
                cell.Offset(0, 1).Value = Date ' 填入今天日期
            ElseIf cell.Value = "未交" Or cell.Value = "不適用" Or cell.Value = "" Then
                cell.Offset(0, 1).ClearContents ' 結合你原本 Sheet1 的邏輯，這三種情況才清除
            End If
        Next cell
        Application.EnableEvents = True ' 恢復監聽
    End If

    ' 4. --- 健檢報告模組：監控 X 欄，日期填入 Y 欄 ---
    If Not Intersect(Target, Sh.Range("X:X")) Is Nothing Then
        Application.EnableEvents = False ' 暫停事件監聽，防止無限迴圈
        For Each cell In Intersect(Target, Sh.Range("X:X"))
            If cell.Value = "已收" Then
                cell.Offset(0, 1).Value = Date ' 填入今天日期
            ElseIf cell.Value = "未交" Or cell.Value = "不適用" Or cell.Value = "" Then
                cell.Offset(0, 1).ClearContents ' 結合你原本 Sheet1 的邏輯，這三種情況才清除
            End If
        Next cell
        Application.EnableEvents = True ' 恢復監聽
    End If

    Exit Sub

' 防錯處理區：確保就算程式發生預期外的錯誤，也能恢復自動化功能
ErrorHandler:
    Application.EnableEvents = True
End Sub
