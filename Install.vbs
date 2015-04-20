'=========================================================================================
'SetUp ExcelAddIn
'=========================================================================================

Const FILLE_NAME="VBAFormatter.xlam"

Call Exec

Sub Exec()
    Dim objExcel
    Dim strAdPath
    Dim strMyPath
    Dim strAdCp
    Dim strMyCp
    Dim objFileSys
    Dim oAdd
    
    '-- CreateObject
    Set objExcel = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    '-- Set Path
    strAdPath = objExcel.Application.UserLibraryPath
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAdCp = objFileSys.BuildPath(strAdPath, FILLE_NAME)
    strMyCp = objFileSys.BuildPath(strMyPath, FILLE_NAME)
    '-- CopyFile
    objFileSys.CopyFile strMyCp, strAdCp
    '-- Add to Excel 
    objExcel.Workbooks.Add
    Set oAdd = objExcel.AddIns.Add(strAdCp,True)
    oAdd.Installed = True
    objExcel.Quit
    '-- Free Object
    Set objExcel = Nothing
    Set objFileSys = Nothing
    
    MsgBox "Complete!"
End Sub