# 🚀 HR Onboarding & Compliance Automation System (VBA)
**人事報到追蹤與 HSE 跨部門協作自動化系統**

## 📖 專案概述 (Project Overview)
在大型跨國企業的快節奏營運中，新進員工的合規文件（如良民證、健檢報告）繳交進度若僅依賴人工在 Excel 打勾記錄，將面臨極大的管理盲區與勞基法規風險。
本專案透過 **Excel VBA** 與 **事件驅動邏輯 (Event-driven Programming)**，將靜態表單升級為具備「狀態感知」與「時間戳記」的自動化追蹤系統，徹底解決 HRIS 系統名單同步、跨部門交接與合規稽核的營運痛點。

---

## ⚠️ 商業痛點 (Business Pain Points)
1. **合規紅線風險**：過去缺乏文件繳交的「時間維度」，難以精準追蹤「已報到滿 30 天但尚未補齊法規文件」的高風險名單。
2. **跨部門交接無效率**：良民證與健檢報告需交接給 HSE（職安衛）部門，手動篩選耗時且極易遺漏。
3. **資料庫更新繁瑣**：每次有新進員工，HR 需手動核對 HRIS 匯出名單，逐筆建檔並初始化繳交狀態。

---

## 🎯 系統架構與核心功能 (Features & Architecture)

![UI介面](image_9758c3.png)
*(預留說明：若圖片已上傳至同資料夾，上方語法即可直接顯示圖片)*

本系統設計了專屬的 UI 功能區，實現以下四大核心功能：
* **🔄 HRIS 智能同步**：自動比對新舊名單，新進人員自動建檔並預設「未交」，既有員工僅更新異動資訊。
* **⏱️ 隱形時間戳記 (Event Listener)**：運用底層事件監聽，當狀態切換為「已收」時，系統瞬間自動押上日期，完全防呆。
* **📤 一鍵跨部門報件**：自動篩選「待處理」狀態，產出標準化 HSE / 良民證交接報表，並透過關鍵字判定職務屬性（Operator）。
* **🚨 30 天合規稽核**：自動結算滿 30 天未完成繳交的異常清單，並依據廠區 (Site) 自動分群排序。

---

## 💻 核心邏輯展示 (Core Code Snippets)
> **💡 Note:** 此處僅展示最具代表性的系統底層邏輯。完整的原始碼與效能優化細節，請參閱本資料夾內的 `.bas` 模組檔案。

### 1. 狀態改變觸發時間戳記 (Event-Driven Timestamps)
捨棄人工輸入日期，利用 `Workbook_SheetChange` 建立背景監聽，確保合規稽核的時間點絕對準確，並寫入防呆機制避免無限迴圈。

```vba
' 監控良民證狀態 (R 欄)，狀態為"已收"時自動於 S 欄填入日期
If Not Intersect(Target, Sh.Range("R:R")) Is Nothing Then
    Application.EnableEvents = False ' 暫停監聽，防止資源耗盡
    For Each cell In Intersect(Target, Sh.Range("R:R"))
        If cell.Value = "已收" Then
            cell.Offset(0, 1).Value = Date 
        ElseIf cell.Value = "未交" Or cell.Value = "不適用" Or cell.Value = "" Then
            cell.Offset(0, 1).ClearContents
        End If
    Next cell
    Application.EnableEvents = True 
End If
