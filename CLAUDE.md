Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

---

## 專案總覽：Estimate Dashboard

本專案從 YouTrack 匯出的工時資料，依 Owner 產出預估工時 vs 實際工時的 Excel Dashboard。

### 資料來源

| 檔案 | 說明 |
|---|---|
| `data/ESTIMATE.xlsx` | 原始資料（sheet: `in`），欄位：Estimate、TimeSpend、Planned For、Type、ID、Summary、Owner、State |

### Pipeline 入口

**`pipeline.py`** — 統一執行入口。

| 步驟名稱 | 對應腳本 | 說明 |
|---|---|---|
| `data-category` | `split_estimate.py` | 將原始資料分成兩份 xlsx |
| `estimate-dashboard` | `build_estimate_dashboard.py` | 產生 Owner Dashboard |

用法：
```
python pipeline.py                                    # 預設執行 data-category
python pipeline.py --step data-category               # 重新分組
python pipeline.py --step estimate-dashboard          # 重建 Dashboard
python pipeline.py --step data-category estimate-dashboard  # 依序執行兩步
python pipeline.py --all                              # 全部步驟
```

### 各腳本職責

**`split_estimate.py`**
- 分組規則：`Estimate` 與 `TimeSpend` 均有值 **且** `Type == "Task"` → `ESTIMATE_with.xlsx`；其餘（含 Issue、任一欄為空）→ `ESTIMATE_without.xlsx`

**`build_estimate_dashboard.py`**
- 讀取 `data/ESTIMATE_with.xlsx`，解析時間字串為小時（float）
- 時間解析：`parse_hours()` 支援 `"1 hour 30 mins"` → `1.5h`、`"30 mins"` → `0.5h` 等格式
- 產出 `data/ESTIMATE_dashboard.xlsx`，包含：
  - **Summary** 頁籤：整體 KPI（Total Estimate/TimeSpend/Utilization/Tasks）+ 兩張圖表（by Owner、by Sprint）
  - **12 個 Owner 頁籤**：個人 KPI、Sprint 比較圖表、明細表
- 圖表資料存於隱藏工作表 `_ChartData`（避免 Excel 不讀隱藏欄的問題）

### 圖表規格

| 項目 | 規格 |
|---|---|
| 類型 | 水平長條圖（type="bar"）|
| X 軸 | Hours（數值軸，底部）|
| Y 軸 | Planned For / Owner（類別軸，左側）|
| Estimate(h) 顏色 | 淺藍 `#9DC3E6` |
| TimeSpend(h) 顏色 | 淺綠 `#A6D260` |
| 無資料的 Sprint | 自動略過，不顯示 |
| 位置（Owner 頁籤）| 右側 col J，不遮蓋 A–H 資料 |
| 位置（Summary）| 右側 col O |

> **openpyxl 水平長條圖注意**：`x_axis` 對應類別軸（左側），`y_axis` 對應數值軸（底部），與直覺相反。設定軸標題時需對調。

### 輸出檔案

| 檔案 | 說明 |
|---|---|
| `data/ESTIMATE_with.xlsx` | 有效 Task 資料（130 筆）|
| `data/ESTIMATE_without.xlsx` | Issue 及工時不完整資料（401 筆）|
| `data/ESTIMATE_dashboard.xlsx` | 13 頁籤 Dashboard（Summary + 12 Owner）|

---

Tradeoff: These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## Pre-Response Checklist

Before every response, explicitly run through these checks:

1. **Assumptions** — List any assumptions I'm making. If any are uncertain, stop and ask instead of guessing.
2. **Scope** — Is everything I'm about to produce explicitly requested? If not, name the extras and ask whether they're wanted.
3. **Clarity** — Is there anything in the request I don't fully understand? If yes, name it and ask — don't invent an interpretation.

If any check flags an issue, surface it to the user before proceeding.

1. Think Before Coding
Don't assume. Don't hide confusion. Surface tradeoffs.

Before implementing:

State your assumptions explicitly. If uncertain, ask.
If multiple interpretations exist, present them - don't pick silently.
If a simpler approach exists, say so. Push back when warranted.
If something is unclear, stop. Name what's confusing. Ask.
2. Simplicity First
Minimum code that solves the problem. Nothing speculative.

No features beyond what was asked.
No abstractions for single-use code.
No "flexibility" or "configurability" that wasn't requested.
No error handling for impossible scenarios.
If you write 200 lines and it could be 50, rewrite it.
Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

3. Surgical Changes
Touch only what you must. Clean up only your own mess.

When editing existing code:

Don't "improve" adjacent code, comments, or formatting.
Don't refactor things that aren't broken.
Match existing style, even if you'd do it differently.
If you notice unrelated dead code, mention it - don't delete it.
When your changes create orphans:

Remove imports/variables/functions that YOUR changes made unused.
Don't remove pre-existing dead code unless asked.
The test: Every changed line should trace directly to the user's request.

4. Goal-Driven Execution
Define success criteria. Loop until verified.

Transform tasks into verifiable goals:

"Add validation" → "Write tests for invalid inputs, then make them pass"
"Fix the bug" → "Write a test that reproduces it, then make it pass"
"Refactor X" → "Ensure tests pass before and after"
For multi-step tasks, state a brief plan:

1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

These guidelines are working if: fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.