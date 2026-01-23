' 批次轉換工具 - 將 .xls 轉換為 .xlsx
' 使用方式：將此檔案放到包含 .xls 的資料夾，然後雙擊執行

Option Explicit

Dim fso, folder, file, excel, wb
Dim folderPath, xlsFiles, successCount, failCount

Set fso = CreateObject("Scripting.FileSystemObject")

' 取得此腳本所在的資料夾
folderPath = fso.GetParentFolderName(WScript.ScriptFullName)

' 或者讓使用者選擇資料夾
Dim shell
Set shell = CreateObject("Shell.Application")
Dim folderItem
Set folderItem = shell.BrowseForFolder(0, "請選擇包含 .xls 檔案的資料夾", 0, folderPath)

If folderItem Is Nothing Then
    WScript.Echo "已取消操作"
    WScript.Quit
End If

folderPath = folderItem.Self.Path

' 啟動 Excel
On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "錯誤：無法啟動 Excel。請確認已安裝 Microsoft Excel。"
    WScript.Quit
End If
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False

Set folder = fso.GetFolder(folderPath)
successCount = 0
failCount = 0

WScript.Echo "開始轉換資料夾: " & folderPath
WScript.Echo "-------------------------------------------"

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
        Dim baseName, xlsxPath
        baseName = fso.GetBaseName(file.Name)
        xlsxPath = folderPath & "\" & baseName & ".xlsx"
        
        ' 跳過已存在的
        If fso.FileExists(xlsxPath) Then
            WScript.Echo "跳過 (已存在): " & baseName & ".xlsx"
        Else
            On Error Resume Next
            Set wb = excel.Workbooks.Open(file.Path)
            If Err.Number = 0 Then
                wb.SaveAs xlsxPath, 51  ' 51 = xlsx format
                wb.Close False
                WScript.Echo "完成: " & baseName & ".xlsx"
                successCount = successCount + 1
            Else
                WScript.Echo "失敗: " & baseName & ".xls - " & Err.Description
                failCount = failCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
Next

excel.Quit
Set excel = Nothing
Set fso = Nothing

WScript.Echo "-------------------------------------------"
WScript.Echo "轉換完成！成功: " & successCount & ", 失敗: " & failCount
WScript.Echo "按確定關閉此視窗"
