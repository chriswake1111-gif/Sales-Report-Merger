import sys
import pandas as pd
import json
import os
import warnings
import tempfile
import shutil
import subprocess

warnings.filterwarnings("ignore")

if sys.version_info >= (3, 7):
    sys.stdout.reconfigure(encoding='utf-8')

def log(msg):
    print(f"[DIAG] {msg}", file=sys.stderr)


def try_repair_mojibake(text):
    """嘗試修復常見的 Big5/CP950 亂碼"""
    if not isinstance(text, str):
        return text
    # 已經有中文就不處理
    if any('\u4e00' <= c <= '\u9fff' for c in text):
        return text
    try:
        return text.encode('latin-1').decode('cp950')
    except:
        pass
    try:
        return text.encode('cp1252').decode('cp950')
    except:
        pass
    return text


def convert_xls_via_excel_subprocess(xls_path, timeout=5):
    """使用子程序呼叫 Excel COM，5秒超時"""
    log("Trying Excel COM (5s timeout)...")
    
    temp_dir = tempfile.mkdtemp()
    base_name = os.path.splitext(os.path.basename(xls_path))[0]
    xlsx_path = os.path.join(temp_dir, base_name + ".xlsx")
    helper_script = os.path.join(temp_dir, "com_helper.py")
    
    helper_code = f'''
import sys, os
try:
    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(r"{os.path.abspath(xls_path)}", UpdateLinks=0, ReadOnly=True, Local=True)
    wb.SaveAs(r"{xlsx_path}", FileFormat=51)
    wb.Close(False)
    excel.Quit()
    print("SUCCESS:" + r"{xlsx_path}" if os.path.exists(r"{xlsx_path}") else "ERROR:No file")
except Exception as e:
    print("ERROR:" + str(e))
finally:
    try: pythoncom.CoUninitialize()
    except: pass
'''
    
    with open(helper_script, 'w', encoding='utf-8') as f:
        f.write(helper_code)
    
    try:
        result = subprocess.run(
            [sys.executable, helper_script],
            capture_output=True, text=True, timeout=timeout, cwd=temp_dir
        )
        output = result.stdout.strip()
        log(f"COM result: {output[:100]}")
        
        if output.startswith("SUCCESS:"):
            return output.replace("SUCCESS:", ""), None
        return None, output.replace("ERROR:", "") if output.startswith("ERROR:") else output
            
    except subprocess.TimeoutExpired:
        log("COM timed out, killing Excel...")
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], capture_output=True, timeout=3)
        except:
            pass
        return None, "Timeout"
    except Exception as e:
        return None, str(e)
    finally:
        try: os.remove(helper_script)
        except: pass


def smart_convert_value(x):
    """
    Smartly convert strings to numbers:
    - If it starts with '0' and length > 1 (e.g., "0123"), keep as string (likely ID/phone).
    - If it looks like a number (e.g., "123", "10.5"), convert to int/float.
    - Otherwise keep as string.
    """
    if not isinstance(x, str):
        return x
    
    x = x.strip()
    if not x:
        return ""
        
    # Check for leading zero ID
    if x.startswith('0') and len(x) > 1 and x.replace('.', '', 1).isdigit():
        # Exception: "0.5" is a number, not an ID usually. 
        # But "0123" is an ID.
        if '.' in x:
            try:
                val = float(x)
                if val < 1 and val > -1: # It's like 0.5
                    return val
            except:
                pass
        return x

    # Try convert to number
    try:
        if '.' in x:
            return float(x)
        return int(x)
    except:
        return x

def parse_excel(file_path):
    log(f"Parsing: {file_path}")
    
    try:
        if not os.path.exists(file_path):
            return {"success": False, "error": f"File not found: {file_path}"}

        ext = os.path.splitext(file_path)[1].lower()
        df = None
        method = "unknown"
        temp_xlsx_path = None

        if ext == '.xls':
            # 快速嘗試 COM (5秒)
            xlsx_path, error = convert_xls_via_excel_subprocess(file_path, timeout=5)
            
            if xlsx_path and os.path.exists(xlsx_path):
                log("COM success!")
                temp_xlsx_path = xlsx_path
                # FORCE STRING to preserve leading zeros
                df = pd.read_excel(xlsx_path, engine='openpyxl', dtype=str)
                method = "excel_com"
            else:
                log(f"COM failed ({error}), using xlrd...")
                method = "xlrd"
                try:
                    import xlrd
                    wb = xlrd.open_workbook(file_path, encoding_override="cp950")
                    # FORCE STRING
                    df = pd.read_excel(wb, engine='xlrd', dtype=str)
                except:
                    # FORCE STRING
                    df = pd.read_excel(file_path, engine='xlrd', dtype=str)
                
                # 修復亂碼
                df.columns = [try_repair_mojibake(str(c)) for c in df.columns]
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].apply(try_repair_mojibake)

        elif ext == '.csv':
            for enc in ['cp950', 'big5', 'utf-8']:
                try:
                    # FORCE STRING
                    df = pd.read_csv(file_path, encoding=enc, dtype=str)
                    method = f"csv_{enc}"
                    break
                except:
                    pass
        else:
            # FORCE STRING
            df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
            method = "openpyxl"

        if df is None:
            return {"success": False, "error": "Failed to read file"}

        # 清理
        if temp_xlsx_path:
            try: shutil.rmtree(os.path.dirname(temp_xlsx_path), ignore_errors=True)
            except: pass

        df = df.fillna("")
        
        # Apply smart conversion
        # We iterate columns to be safe
        for col in df.columns:
            df[col] = df[col].apply(smart_convert_value)

        log(f"Method: {method}, Rows: {len(df)}, Headers: {list(df.columns)[:3]}")

        return {
            "success": True,
            "data": df.to_dict(orient='records'),
            "headers": list(df.columns),
            "rowCount": len(df)
        }

    except Exception as e:
        log(f"Error: {e}")
        return {"success": False, "error": str(e)}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"success": False, "error": "No file path"}))
        sys.exit(1)
    result = parse_excel(sys.argv[1])
    print(json.dumps(result, ensure_ascii=True))
