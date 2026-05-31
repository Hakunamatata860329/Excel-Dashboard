# Estimate Dashboard

從 YouTrack 匯出的工時資料，自動產生「預估工時 vs 實際工時」的 Excel Dashboard。

---

## 你需要準備什麼

將 YouTrack 匯出的 Excel 放在 `data/input/` 資料夾，檔名為：

```
data/input/ESTIMATE.xlsx
```

**必要的 sheet 與欄位：**

| 欄位 | 說明 | 範例 |
|---|---|---|
| `ID` | Issue ID | `TK-123` |
| `Summary` | 任務標題 | `Fix login bug` |
| `Type` | 類型 | `Task` 或 `Issue` |
| `State` | 狀態 | `Done` |
| `Owner` | 負責人（email 格式） | `jason.chen@company.com` |
| `Planned For` | Sprint 名稱 | `Designer Rev1.14.0 Sprint1` |
| `Estimate` | 預估工時（文字格式） | `2 hours 30 mins` |
| `TimeSpend` | 實際花費工時（文字格式） | `3 hours` |

> Sheet 名稱必須是 `in`（YouTrack 匯出預設即為此名稱）

---

## 安裝環境

```bash
pip install pandas openpyxl
```

---

## 執行方式

```bash
# 完整執行（建議第一次使用）
python pipeline.py --all

# 分步執行
python pipeline.py --step data-category        # 步驟一：資料分類
python pipeline.py --step estimate-dashboard   # 步驟二：產生 Dashboard
```

---

## 你會得到什麼

執行後，`data/output/` 資料夾會產出三個檔案：

### `ESTIMATE_with.xlsx`
符合條件的有效資料（Type = Task 且 Estimate、TimeSpend 均有值）。

### `ESTIMATE_without.xlsx`
排除的資料（Issue 類型，或工時欄位有缺漏）。

### `ESTIMATE_dashboard.xlsx`
主要產出，包含以下分頁：

| 分頁 | 內容 |
|---|---|
| **Summary** | 整體 KPI（總預估/實際工時、使用率、任務數）+ 依 Owner 與 Sprint 的比較圖表 |
| **各 Owner 分頁**（共 12 個） | 個人 KPI、Sprint 比較圖表、所有 Task 明細表（含顏色標記：超時=紅、提前=綠） |

---

## Pipeline 流程

```
data/input/ESTIMATE.xlsx
        │
        ▼
  split_estimate.py              ← 依 Type 與工時是否完整分組
        │
        ├── data/output/ESTIMATE_with.xlsx    (有效 Task)
        └── data/output/ESTIMATE_without.xlsx (排除資料)
                │
                ▼
  build_estimate_dashboard.py    ← 建立 Excel Dashboard
                │
                ▼
        data/output/ESTIMATE_dashboard.xlsx
```
