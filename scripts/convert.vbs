Option Explicit

Dim fso, folder, file, excel, wb
Dim folderPath, successCount, failCount

Set fso = CreateObject("Scripting.FileSystemObject")

Dim shell
Set shell = CreateObject("Shell.Application")
Dim folderItem
Set folderItem = shell.BrowseForFolder(0, "Select folder with .xls files", 0)

If folderItem Is Nothing Then
    WScript.Echo "Cancelled"
    WScript.Quit
End If

folderPath = folderItem.Self.Path

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot start Excel"
    WScript.Quit
End If
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False

Set folder = fso.GetFolder(folderPath)
successCount = 0
failCount = 0

WScript.Echo "Converting: " & folderPath

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
        Dim baseName, xlsxPath
        baseName = fso.GetBaseName(file.Name)
        xlsxPath = folderPath & "\" & baseName & ".xlsx"
        
        If fso.FileExists(xlsxPath) Then
            WScript.Echo "Skip: " & baseName
        Else
            On Error Resume Next
            Set wb = excel.Workbooks.Open(file.Path)
            If Err.Number = 0 Then
                wb.SaveAs xlsxPath, 51
                wb.Close False
                WScript.Echo "Done: " & baseName
                successCount = successCount + 1
            Else
                WScript.Echo "Fail: " & baseName
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

WScript.Echo "Complete! Success: " & successCount & ", Failed: " & failCount
