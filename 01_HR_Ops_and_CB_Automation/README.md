# 🚀 HR Onboarding & Compliance Automation System
**全自動化人事營運 (HR Ops) 與合規追蹤系統**

---

## 📖 專案概述 (Project Overview)
本專案透過 **Excel VBA** 打造全自動化的人事營運 (HR Ops) 與合規追蹤系統。核心解決大型企業或多廠區 (Multi-site) 在處理員工入職合規 (Onboarding Compliance)、健康與安全 (HSE) 體檢追蹤、以及人事主資料庫 (Master Data) 同步時面臨的繁瑣手動比對與人工登錄錯誤。

透過此自動化模組，HR 團隊能一鍵完成從資料同步、狀態連動到產出各類稽核報表的端到端 (End-to-End) 流程，大幅提升人事資料的準確度與管理效率。

---

## 🛠️ 系統架構與工作表 (System Architecture)
* **`StaffTable`**：人事主資料表 (Master Data)，存放所有員工的在職狀態、合規文件繳交進度與個人資訊。
* **`Download`**：由 HRIS 系統匯出之原始異動資料檔。
* **VBA 核心模組**：包含 4 個獨立業務報表/處理模組 (`.bas`)，以及 1 個全域事件監聽模組 (`.cls`)。

---

## 💻 技術堆疊 (Tech Stack)
* **核心語言**：VBA (Visual Basic for Applications)
* **核心技術應用**：
  * 工作表事件監聽 (`Worksheet_Change` Events)
  * 動態陣列、迴圈處理與範圍定位 (`Dynamic Ranges`, `Match`, `Intersect`)
  * 字串解析與重組 (`InStr`, `String Concatenation`)
  * 自動化條件式格式與 UI/UX 排版 (`PasteValidation`, `Interior.Color`)
  * 防呆與全域錯誤處理機制 (`ErrorHandler`, `ScreenUpdating = False`)

---

## 🚀 核心功能模組與程式碼展示 (Core Features & Code Highlights)

### 1. 人事主檔智能同步模組 (Staff List Synchronization)
* **對應檔案**：`SyncStaffList.bas` (Module2)
* **功能描述**：比對 `Download` 工作表（HRIS 匯出原始檔）與 `StaffTable`（人事主資料庫），以「員工編號」為 Key 進行全自動分流處理，精準識別「新進員工建檔」與「既有員工異動更新」。
* **自動化亮點**：
  * **智能分流與狀態初始化**：新進員工自動搬移基本資料 (A~K 欄)，批次將 10 項以上合規追蹤欄位 (M~R, U~X, AB~AE 欄) 預設為「未交」，並將報件紀錄 (T, Z 欄) 初始化為「待處理」。
  * **格式與公式動態繼承**：自動複製並延伸關鍵計算公式（AF 欄入職合規、AG 欄站點、AH 欄員編排序），並透過 `xlPasteValidation` 與 `xlPasteFormats` 完整繼承下拉選單資料驗證與條件格式。
  * **在職狀態智能判定**：依據第 11 欄（離職日期）是否存在數值，自動判定並寫入在職狀態（「在職」或「離職」）。
  * **效能優化與防錯設計**：批次處理前關閉螢幕更新 (`ScreenUpdating = False`) 與自動重算 (`xlCalculationManual`)，並加入 `ErrorHandler` 與防呆確認視窗，執行完畢後精確回報新進與更新筆數。

```vba
' 核心片段：SyncStaffList 智能分流與格式自動繼承邏輯
matchRow = Application.Match(wsDown.Cells(i, "A").Value, wsMaster.Columns("A"), 0)

If IsError(matchRow) Then
    ' =================================================================
    ' --- A. 處理新進員工 ---
    ' =================================================================
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
    
    ' 1. 搬運基本資料 (A~K 欄) 並自動判定在職狀態 (L 欄)
    For col = 1 To 11
        wsMaster.Cells(lastRowMaster, col).Value = wsDown.Cells(i, col).Value
    Next col
    wsMaster.Cells(lastRowMaster, 12).Value = IIf(wsDown.Cells(i, 11).Value <> "", "離職", "在職")
    
    ' 2. 批次初始化「未交」與「待處理」狀態
    wsMaster.Range("M" & lastRowMaster & ":R" & lastRowMaster).Value = "未交"
    wsMaster.Range("U" & lastRowMaster & ":X" & lastRowMaster).Value = "未交"
    wsMaster.Range("AB" & lastRowMaster & ":AE" & lastRowMaster).Value = "未交"
    wsMaster.Cells(lastRowMaster, "T").Value = "待處理"
    wsMaster.Cells(lastRowMaster, "Z").Value = "待處理"
    
    ' 3. 動態繼承公式 (AF 入職合規, AG 站點, AH 員編排序)
    wsMaster.Cells(2, "AF").Copy Destination:=wsMaster.Cells(lastRowMaster, "AF")
    wsMaster.Cells(2, "AG").Copy Destination:=wsMaster.Cells(lastRowMaster, "AG")
    wsMaster.Cells(2, "AH").Copy Destination:=wsMaster.Cells(lastRowMaster, "AH")
    
    ' 4. 繼承格式與下拉選單資料驗證 (Data Validation)
    wsMaster.Rows(2).Copy
    wsMaster.Rows(lastRowMaster).PasteSpecial Paste:=xlPasteFormats
    wsMaster.Range("M2:AE2").Copy
    wsMaster.Range("M" & lastRowMaster & ":AE" & lastRowMaster).PasteSpecial Paste:=xlPasteValidation
    
    newCount = newCount + 1
Else
    ' =================================================================
    ' --- B. 處理既有員工 (僅更新異動欄位) ---
    ' =================================================================
    For col = 6 To 11
        wsMaster.Cells(matchRow, col).Value = wsDown.Cells(i, col).Value
    Next col
    
    If wsDown.Cells(i, 11).Value <> "" Then wsMaster.Cells(matchRow, 12).Value = "離職"
    
    updateCount = updateCount + 1
End If
```
---

### 2. 即時狀態連動與日期戳記 (Event-Driven Auto-Date Stamp)
* **對應模組**：`Workbook_SheetChange.cls`
* **功能描述**：透過 Excel 全域事件監聽，實現無感自動化資料登錄，減少 HR 人為輸入負擔。
* **自動化亮點**：
  * **良民證與健檢追蹤**：當 HR 於主檔將狀態下拉更改為「已收」時，系統瞬間於相鄰欄位自動押上「今日日期」(`Date`)。
  * **防錯防呆機制**：若狀態被改回「未交」、「不適用」或清空，系統會自動清除繳交日期。內建 `CountLarge` 門檻，防止大量貼上資料時造成系統當機。

```vba
' 核心片段：監控特定欄位異動並自動寫入/清除時間戳記
If Not Intersect(Target, Sh.Range("R:R")) Is Nothing Then
    Application.EnableEvents = False ' 暫停事件監聽，防止無限迴圈
    For Each cell In Intersect(Target, Sh.Range("R:R"))
        If cell.Value = "已收" Then
            cell.Offset(0, 1).Value = Date ' 自動押上當日時間戳記
        ElseIf cell.Value = "未交" Or cell.Value = "不適用" Or cell.Value = "" Then
            cell.Offset(0, 1).ClearContents
        End If
    Next cell
    Application.EnableEvents = True 
End If
```

---

### 3. 新人入職 30 天合規催繳報表 (30-Days Missing Document Report)
* **對應模組**：`Export_30Days_MissingReport.bas`
* **功能描述**：針對入職滿 30 天且文件未齊全的在職人員，自動產出各站點的催繳清單。
* **自動化亮點**：
  * **動態條件篩選**：結合「在職 (`Trim = 在職`)」、「文件未齊 (`InStr <> 全齊`)」、「年資 ≥ 30 天 (`DateDiff >= 30`)」進行精準攔截。
  * **法律紅線缺件字串重組**：自動掃描 10 項關鍵文件，並將缺交項目重組成結構化字串（如：`"合約, 良民證, 健檢報告"`）。
  * **視覺化分組**：報表依據「站點 (Site)」進行自動排序與視覺化分層（自動插入藍色底色合併儲存格作為站點標題）。

```vba
' 核心片段：法規紅線多重條件判斷與缺件清單重組
If Trim(wsMaster.Cells(i, "L").Value) = "在職" And _
   InStr(1, wsMaster.Cells(i, "AF").Value, "全齊") = 0 And _
   wsMaster.Cells(i, "AF").Value <> "" And _
   DateDiff("d", wsMaster.Cells(i, "E").Value, Date) >= 30 Then
    
    ' 掃描並重組字串
    missingList = ""
    If Trim(wsMaster.Cells(mRow, "M").Value) = "未交" Then missingList = missingList & "合約, "
    If Trim(wsMaster.Cells(mRow, "R").Value) = "未交" Then missingList = missingList & "良民證, "
    If Trim(wsMaster.Cells(mRow, "X").Value) = "未交" Then missingList = missingList & "健檢報告, "
    If Len(missingList) > 0 Then missingList = Left(missingList, Len(missingList) - 2)
    
    wsRep.Cells(rRow, 7).Value = missingList
    wsRep.Cells(rRow, 7).Font.Color = RGB(200, 0, 0)
End If
```

---

### 4. HSE 入職健檢報件系統 (HSE Health Check Dispatch)
* **對應模組**：`Export_HSE_Report.bas`
* **功能描述**：為職安衛 (HSE) 單位自動產出新進員工體檢追蹤名單，符合法規申報格式。
* **自動化亮點**：
  * **高危職務自動標記**：透過字串比對 (`InStr`) 識別職稱是否包含 `"Operator"`，並自動標記 `V / X`。
  * **證照與風險警示**：自動抓取「堆高機證照」並以紅色醒目標示；若員工於申報期間已離職，整列自動標記為黃底警告。
  * **狀態推進**：報表產出後，自動將主檔狀態由「待處理」更新為「已發送」，形成管理閉環。

---

### 5. 良民證繳交稽核概況 (Police Record Compliance Audit)
* **對應模組**：`Export_PoliceRecord_Compliance_Report.bas`
* **功能描述**：一鍵產出全公司背景調查 (Background Check) 文件的合規狀態總覽。
* **自動化亮點**：
  * **優先順序重排**：打破原有人事大表排序，強制將「【未交/催繳】」人員置頂並標示為紅色字體，提升稽核直覺性。
  * **模組化代碼設計**：將填表邏輯獨立拆分為 `FillData` 副程式，提高程式碼重用性與維護性。
```vba
' 核心片段：讓 HR 一眼就能辨識需要催繳的對象。
' --- 第一輪：先抓「未交」的人 (排在前面並標紅字比較醒目) ---
For i = 2 To lastRow
    If Trim(wsMaster.Cells(i, 12).Value) = "在職" And wsMaster.Cells(i, 18).Value <> "已收" Then
        wsRep.Cells(rRow, 1).Value = "【未交/催繳】"
        wsRep.Cells(rRow, 1).Font.Color = vbRed
        Call FillData(wsRep, wsMaster, rRow, i) 
        wsRep.Range("A" & rRow & ":G" & rRow).Font.Color = vbRed 
        rRow = rRow + 1
    End If
Next i
```

---

## 📈 商業價值與效益 (Business Impact)
* **消除人為疏漏 (Error Reduction)**：透過防呆機制與自動化初始化，避免新進員工建檔時漏填合規追蹤欄位，確保 **100%** 的追蹤覆蓋率。
* **法遵風險管控 (Compliance Management)**：精準抓取「滿 30 天未交齊」及 HSE 特殊身分（操作員 / 堆高機），大幅降低企業面臨的勞檢與工安合規風險。
* **工時大幅節省 (Time-Saving)**：原先需耗時數小時跨表比對 (VLOOKUP)、手動上色與篩選的例行公事，縮短至 **單擊按鈕 (1-Click)** 於秒級內完成，釋放 HR 團隊行政量能。
