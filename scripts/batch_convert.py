"""
批次轉換工具：將資料夾內所有 .xls 轉換為 .xlsx
"""

import sys
import os
import time
import glob

def convert_folder(folder_path):
    try:
        import win32com.client.dynamic
        import pythoncom
    except ImportError:
        print("錯誤：請先安裝 pywin32")
        print("執行：pip install pywin32")
        return

    if not os.path.isdir(folder_path):
        print(f"錯誤：資料夾不存在 - {folder_path}")
        return

    xls_files = glob.glob(os.path.join(folder_path, "*.xls"))
    xls_files = [f for f in xls_files if not f.lower().endswith('.xlsx')]

    if not xls_files:
        print(f"在 {folder_path} 中找不到任何 .xls 檔案")
        return

    print(f"找到 {len(xls_files)} 個 .xls 檔案，開始轉換...")
    print("-" * 50)

    pythoncom.CoInitialize()
    excel = None
    
    try:
        # 使用動態綁定，不需要註冊型別庫
        excel = win32com.client.dynamic.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        success_count = 0
        fail_count = 0

        for i, xls_path in enumerate(xls_files, 1):
            base_name = os.path.splitext(os.path.basename(xls_path))[0]
            xlsx_path = os.path.join(folder_path, base_name + ".xlsx")

            if os.path.exists(xlsx_path):
                print(f"[{i}/{len(xls_files)}] 跳過 (已存在): {base_name}.xlsx")
                continue

            try:
                print(f"[{i}/{len(xls_files)}] 轉換中: {base_name}.xls ... ", end="", flush=True)
                
                wb = excel.Workbooks.Open(os.path.abspath(xls_path))
                wb.SaveAs(xlsx_path, FileFormat=51)
                wb.Close(False)
                
                print("完成")
                success_count += 1
                
            except Exception as e:
                print(f"失敗: {str(e)[:50]}")
                fail_count += 1

        print("-" * 50)
        print(f"轉換完成！成功: {success_count}, 失敗: {fail_count}")
        
    except Exception as e:
        print(f"Excel 啟動失敗: {e}")
        print("\n嘗試備用方案...")
        try_alternative(xls_files, folder_path)
    finally:
        if excel:
            try:
                excel.Quit()
                time.sleep(1)
            except:
                pass
        pythoncom.CoUninitialize()


def try_alternative(xls_files, folder_path):
    """備用方案：使用 subprocess 呼叫 Excel"""
    import subprocess
    
    vbs_code = '''
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(WScript.Arguments(0))
objWorkbook.SaveAs WScript.Arguments(1), 51
objWorkbook.Close False

objExcel.Quit
'''
    
    # 建立臨時 VBS 檔案
    vbs_path = os.path.join(folder_path, "_convert.vbs")
    with open(vbs_path, 'w', encoding='utf-8') as f:
        f.write(vbs_code)
    
    success = 0
    fail = 0
    
    try:
        for i, xls_path in enumerate(xls_files, 1):
            base_name = os.path.splitext(os.path.basename(xls_path))[0]
            xlsx_path = os.path.join(folder_path, base_name + ".xlsx")
            
            if os.path.exists(xlsx_path):
                continue
            
            print(f"[{i}/{len(xls_files)}] {base_name}.xls ... ", end="", flush=True)
            
            try:
                result = subprocess.run(
                    ['cscript', '//Nologo', vbs_path, os.path.abspath(xls_path), xlsx_path],
                    capture_output=True, text=True, timeout=60
                )
                if os.path.exists(xlsx_path):
                    print("完成")
                    success += 1
                else:
                    print("失敗")
                    fail += 1
            except Exception as e:
                print(f"失敗: {e}")
                fail += 1
    finally:
        try:
            os.remove(vbs_path)
        except:
            pass
    
    print("-" * 50)
    print(f"備用方案完成！成功: {success}, 失敗: {fail}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("請提供資料夾路徑")
        print('範例: python batch_convert.py "C:\\資料夾路徑"')
        sys.exit(1)

    convert_folder(sys.argv[1])
