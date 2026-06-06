# vlookup_toolset

Python CLI 工具集，提供 Excel / ODS 資料處理功能，包含 VLOOKUP 合併、欄位篩選與文字檔轉換。

---

## 安裝依賴

```bash
pip install pandas odfpy openpyxl
```

---

## 使用方式

所有功能統一透過 `excel_tools.py` 執行，以子命令分開：

### 1. `vlookup` — 跨檔合併

以 `test2` 的 ID 查找 `test1` 的薪資資料（類似 Excel VLOOKUP），輸出合併結果。

```bash
# 使用預設路徑
python excel_tools.py vlookup

# 自訂來源與輸出
python excel_tools.py vlookup --src1 test1.ods --src2 test2.ods --output output_vlookup.xlsx
```

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--src1` | 薪資資料來源（被查找表） | `test1.ods` |
| `--src2` | 查詢條件來源（包含 ID） | `test2.ods` |
| `--output` | 輸出檔案路徑 | `output_vlookup.xlsx` |

---

### 2. `extract` — 欄位篩選

從來源檔案中選取指定欄位，輸出為新的 Excel 檔案。

```bash
# 輸出 ID 與 Net Pay 兩欄
python excel_tools.py extract --columns "ID,Net Pay"

# 輸出多個欄位
python excel_tools.py extract --columns "ID,Base Salary,Tax,Net Pay"

# 自訂來源與輸出
python excel_tools.py extract --src test1.ods --columns "ID,Net Pay" --output my_output.xlsx
```

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--src` | 來源檔案路徑 | `test1.ods` |
| `--columns` | 欄位名稱，逗號分隔 | `ID,Net Pay` |
| `--output` | 輸出檔案路徑 | `output_extract.xlsx` |

---

### 3. `txt2ods` — 文字檔轉 ODS

掃描資料夾內所有 `.txt` / `.csv` / `.tsv`，轉為 ODS 格式並在數字欄位末列加入合計。

```bash
# 使用預設資料夾
python excel_tools.py txt2ods

# 指定資料夾與輸出位置
python excel_tools.py txt2ods --folder "C:\data\input" --output-folder "C:\data\output"
```

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--folder` | 文字檔所在資料夾 | 腳本所在資料夾 |
| `--output-folder` | 輸出資料夾 | 同來源資料夾 |

**自動處理項目：**
- 自動偵測分隔符（`|` / `\t` / `,`）
- 清除尾端多餘分隔符
- 保留前導零的文字欄位（如員工編號 `000001`）
- 數字欄位自動加總，顯示於最後一列（黃底粗體）

---

## 輸出格式

所有輸出檔案統一套用以下樣式：

| 類型 | 格式 |
|------|------|
| 文字欄位 | 左對齊，格式 `@`（保留前導零） |
| 數字欄位 | 右對齊，千分位 `#,##0` |
| 日期欄位 | 置中，`YYYY-MM-DD` |
| 標題列 | 深藍底 `#2F5597`，白字粗體 |
| 資料列 | 隔行淺藍 `#DCE6F1` |
| 合計列 | 黃底 `#FFF2CC`，粗體（僅 txt2ods） |

---

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `excel_tools.py` | 主程式（包含所有功能） |
| `test1.ods` | 範例薪資資料（來源表） |
| `test1.txt` | 範例薪資資料（文字格式） |
| `test2.ods` | 範例查詢條件（含起迄日期） |
| `output_vlookup.xlsx` | vlookup 輸出範例 |
| `output_extract.xlsx` | extract 輸出範例 |
| `test1_converted.ods` | txt2ods 輸出範例 |

---

## 查看說明

```bash
python excel_tools.py --help
python excel_tools.py vlookup --help
python excel_tools.py extract --help
python excel_tools.py txt2ods --help
```
