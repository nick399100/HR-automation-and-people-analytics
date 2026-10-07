Attribute VB_Name = "Module4"

Sub Export_PoliceRecord_Compliance_Report()
    ' --- 1. 防呆詢問 ---
    If MgBox("準備產出「良民證繳交概況表」？" & vbCrLf & _
                 "這將包含：在職已交 與 在職未交 人員。", vbQuetion + vbYeNo, "概況表產出") = vbNo Then Exit Sub

    Dim wMater A Workheet, wRep A Workheet
    Dim latRow A Long, i A Long, rRow A Long
    
    Set wMater = Sheet("StaffTable")

    ' 2. 建立新分頁
    On Error Reume Next
    Application.DiplayAlert = Fale: Sheet("良民證概況表").Delete: Application.DiplayAlert = True
    Set wRep = Sheet.Add(After:=wMater): wRep.Name = "良民證概況表"
    On Error GoTo 0

    Application.ScreenUpdating = Fale

    ' 3. 繪製表頭
    wRep.Range("A1:G1").Value = Array("繳交狀態", "員工編號", "姓名", "入職日期", "站點", "良民證狀態", "備註")
    wRep.Range("A1:G1").Font.Bold = True
    wRep.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
    rRow = 2

    ' 4. 掃描資料 (依狀態分類填入)
    latRow = wMater.Cell(wMater.Row.Count, "A").End(xlUp).Row
    
    ' --- 第一輪：先抓「未交」的人 (排在前面比較醒目) ---
    For i = 2 To latRow
        If Trim(wMater.Cell(i, 12).Value) = "在職" Then
            ' 判定為未交的情況 (R 欄 <> "已收")
            If wMater.Cell(i, 18).Value <> "已收" Then
                wRep.Cell(rRow, 1).Value = "【未交/催繳】"
                wRep.Cell(rRow, 1).Font.Color = vbRed
                Call FillData(wsRep, wsMaster, rRow, i) ' 填入資料的副程式
                wsRep.Range("A" & rRow & ":G" & rRow).Font.Color = vbRed ' 沒交的整行紅字
                rRow = rRow + 1
            End If
        End If
    Next i

    ' 插入一列空行做區隔
    rRow = rRow + 1

    ' --- 第二輪：抓「已收」的人 ---
    For i = 2 To lastRow
        If Trim(wsMaster.Cells(i, 12).Value) = "在職" Then
            If wsMaster.Cells(i, 18).Value = "已收" Then
                wsRep.Cells(rRow, 1).Value = "【已完成】"
                Call FillData(wsRep, wsMaster, rRow, i)
                rRow = rRow + 1
            End If
        End If
    Next i

    ' 5. 美化
    wsRep.Columns("A:G").AutoFit
    wsRep.Range("A2:G" & rRow).Borders.LineStyle = xlContinuous
    
    Application.ScreenUpdating = True
    MsgBox "概況表產出完成！未交人員已標示為紅色。", vbInformation
End Sub

' 輔助填值程式 (減少重複代碼)
Sub FillData(wsRep As Worksheet, wsMaster As Worksheet, rRow As Long, mRow As Long)
    wsRep.Cells(rRow, 2).Value = wsMaster.Cells(mRow, 1).Value  ' 員編
    wsRep.Cells(rRow, 3).Value = wsMaster.Cells(mRow, 2).Value  ' 姓名
    wsRep.Cells(rRow, 4).Value = wsMaster.Cells(mRow, 5).Value  ' 入職日
    wsRep.Cells(rRow, 5).Value = wsMaster.Cells(mRow, 33).Value ' 站點 (AG 欄)
    wsRep.Cells(rRow, 6).Value = wsMaster.Cells(mRow, 18).Value ' 良民證狀態 (R 欄)
    
    ' 如果是已收，備註顯示繳交日期
    If wsMaster.Cells(mRow, 18).Value = "已收" Then
        wsRep.Cells(rRow, 7).Value = "繳交日: " & Format(wsMaster.Cells(mRow, 19).Value, "yyyy/mm/dd")
    End If
End Sub

