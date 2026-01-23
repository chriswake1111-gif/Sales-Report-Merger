---
name: legacy-excel-converter
description: 當使用者需要讀取、轉換舊版 Excel (.xls)、處理繁體中文亂碼問題，或將資料匯入至新系統時使用。
---

# Legacy Excel (.xls) to Modern (.xlsx) Converter Skill

## 技術背景
舊版 Excel (`.xls`) 是二進位 OLE2 格式，新版 (`.xlsx`) 是 XML 格式。
在台灣環境，`.xls` 內容常以 `cp950` (Big5 的擴充集) 編碼儲存。若未指定，Python 容易產生亂碼 (Mojibake)。

## 必要的 Python 依賴
確保環境已安裝：
- `pandas`
- `xlrd >= 2.0.1` (用於讀取 .xls)
- `openpyxl` (用於寫入 .xlsx)

## 核心轉換邏輯 (Standard Operating Procedure)

當撰寫轉換程式碼時，請嚴格遵守以下範式：

### 1. 讀取舊檔 (The Safe Way)
使用 `pandas` 讀取時，雖 `xlrd` 通常能自動偵測，但在處理包含特殊繁體字（如人名、藥品名）時，需注意字串型別。

```python
import pandas as pd

def convert_xls_to_xlsx(input_path: str, output_path: str):
    try:
        # 使用 xlrd 引擎讀取
        # 注意：xlrd 2.0+ 僅支援 .xls，不再支援 .xlsx，這正是我們需要的
        df = pd.read_excel(input_path, engine='xlrd')
        
        # 資料清理：移除前後空白 (常見於舊系統匯出的資料)
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        
        # 寫入為 UTF-8 編碼的 .xlsx
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        return {"success": True, "rows": len(df)}
        
    except UnicodeDecodeError:
        # 如果發生編碼錯誤，提示使用者檔案可能損毀或非標準 Big5
        return {"success": False, "error": "Encoding Error: Try opening in Excel and saving as XLSX manually first."}
    except Exception as e:
        return {"success": False, "error": str(e)}
```
