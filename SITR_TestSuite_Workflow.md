# SITR Test Suite 標準化工作流程設計

> 討論文件 v0.4｜2026-06-02
> 角色：軟體驗證主管
> 目的：建立新機種加入時，可重複套用的 Test Suite 分類標準

---

## 一、三層架構定義

| 層級 | 對應來源 | 說明 |
|-----|---------|------|
| **SITR** | 機種 | 整份測試報告，一個機種一份 |
| **Test Suite** | Function Check List **功能項目（Level 2）** | 每個 Suite = 一個功能項目；機種支援才建立 |
| **Test Case** | Test Case.csv | Suite 底下的具體測試動作 |

```
SITR（機種 W3A）
├── Test Suite：新增裝置     ← 功能項目 Level 2（底下有 Advance/Basic 等 ●Tag）
│     ├── Test Case 001
│     └── Test Case 002
├── Test Suite：Modbus       ← 功能項目 Level 2（Tag = ●）
│     └── Test Case 010
└── （CANopen 不建立，因 ○）
```

---

## 二、功能 Tag 的角色

功能 Tag 是決定「哪些 Test Case 進入 Test Suite」的篩選鍵。

```
功能項目（Level 2）= Test Suite 名稱
  └── Level 3 子項各有對應 Tag
       ● → 此 Tag 的 Test Case 納入 Suite
       ○ → 略過
       N/A（Level 3 不存在）→ Level 2 的 Tag 直接決定整個 Suite
```

---

## 三、下一步：Excel 輸出（Pilot）

**目標：** 以 W3A 為例，產出第一版 Excel，讓 TE 實際看到 Test Suite 的樣子。

### 輸出規格（data/output/TestSuite_W3A.xlsx）

**Sheet 1：Test Suite 清單**

| 欄位 | 說明 |
|-----|-----|
| 功能分類 | Function Check List Level 1 |
| 功能項目（Test Suite 名稱） | Level 2，= Suite 名稱 |
| 功能Tag | Level 3 或 Level 2 識別碼 |
| W3A 支援狀態 | ●/○/TBC |
| 是否建立 Suite | 是/否（● → 是） |

→ 讓主管確認：W3A 應該有哪些 Test Suite？

**Sheet 2：Mapping Table（待填）**

| 欄位 | 說明 |
|-----|-----|
| 功能Tag | 來自 Sheet 1 中「是否建立 Suite = 是」的 Tag |
| 所屬 Test Suite | 自動帶入 |
| 對應 Test Case IDs | **TE 手動填入**（初次建立） |
| 備註 | Gap 標記或說明 |

→ 這是第一次要 TE 手填的工件，後續可複用。

**Sheet 3：Test Case 參考清單**

| 欄位 | 說明 |
|-----|-----|
| Test Case ID（Name） | 來自 Test Case.csv |
| 功能項 | 供 TE 對照填 Mapping |
| 子項目 | |
| 腳本狀態 | 自動/手動 |
| State | 草稿/檢閱中 等 |

→ 方便 TE 查找 Test Case 時有參考依據，不需要自己開原始 CSV。

---

## 四、工作流程（含 Excel 輸出）

```
步驟 1：RD 更新 Function Check List（新機種欄標記 ●/○/TBC）

步驟 2：執行腳本 → 產出 TestSuite_{機種}.xlsx
  └── Sheet 1 自動填入（來源：Function Check List）
  └── Sheet 2 產出空白模板（待填欄位）
  └── Sheet 3 自動填入（來源：Test Case.csv）

步驟 3：TE 主管審核 Sheet 1（Test Suite 清單）
  └── 確認這些 Suite 符合 SITR 驗證範圍

步驟 4：TE 填寫 Sheet 2（Mapping Table）
  └── 對照 Sheet 3，把 Test Case ID 填入對應 Tag

步驟 5：（後續）腳本讀取 Mapping Table → 產出完整 SITR 範本

步驟 6：Post-SITR Review → 更新 Sheet 2（閉環）
```

---

## 五、可重複優化機制

每次新機種後：
- Sheet 2（Mapping Table）複用、只增不減
- 新機種只需更新 Function Check List 的機種欄 → 重跑步驟 2
- Mapping Table 越來越完整 → 步驟 4 手填量逐次減少

---

## 六、實作計劃

**腳本：`build_test_suite.py`**

| 步驟 | 動作 |
|-----|-----|
| 讀取 | Function Check List.xlsx（DIADesigner Function List 頁籤）|
| 讀取 | Test Case.csv |
| 篩選 | 指定機種欄（如 W3A），取出 ● 的功能Tag |
| 產出 | Sheet 1：Test Suite 清單 |
| 產出 | Sheet 2：Mapping Table 模板（空白待填）|
| 產出 | Sheet 3：Test Case 參考清單 |
| 輸出 | data/output/TestSuite_{機種}.xlsx |

**執行方式（整合進 pipeline.py）：**
```
python pipeline.py --step test-suite --model W3A
```

---

## 七、驗證方式

1. 執行腳本，開啟 `data/output/TestSuite_W3A.xlsx`
2. 確認 Sheet 1 的 Suite 清單與 Function Check List 的 W3A 欄 ● 項目一致
3. TE 試填 Sheet 2 的 2–3 個 Tag，看 Mapping 格式是否好用
4. 主管決定：粒度 OK？欄位夠用？流程順暢？
