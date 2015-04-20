Attribute VB_Name = "INICtrl"
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal AppName As String, _
ByVal KeyName As String, _
ByVal Default As String, _
ByVal ReturnedString As String, _
ByVal MaxSize As Long, _
ByVal FileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName As String, _
ByVal KeyName As Any, _
ByVal lpString As Any, _
ByVal FileName As String) As Long

Public Const INI_NAME                As String = "VBAFormatter.Ini"
Public Const INI_SEC_OPT_FORMAT      As String = "OptFormat"
Public Const INI_KEY_TAB_CNT         As String = "Tab_Cnt"
Public Const INI_KEY_ALL_MODULE_EXEC As String = "AllModuleExec"
Public Const INI_KEY_AS_FORMAT       As String = "AsFormat"
Public Const INI_KEY_COMENT_FORMAT   As String = "CommentFormat"
Public Const INI_KEY_COMENT_EXEC     As String = "CommentExec"

Public Type INI_KEY_LIST
    aTabCnt          As Integer
    aIsAllModuleExec As Boolean
    aIsAsFormat      As Boolean
    aIsCommentFormat As Boolean
    aIsCommentExec   As Boolean
End Type

Public cIniKeyList As INI_KEY_LIST

Public Function GetIniValue(aIniKey As String, aIniSection As String) As String
    Dim wIniVal  As String * 1024
    Dim wRet     As Long
    wRet = GetPrivateProfileString(aIniSection, aIniKey, "", wIniVal, Len(wIniVal), ThisWorkbook.Path & "\" & INI_NAME)
    GetIniValue = Left(wIniVal, InStr(wIniVal, vbNullChar) - 1)
    If GetIniValue = "" Then
        Err.Raise 1000 - vbObjectError, _
        "設定ファイルの取得", _
        "設定ファイルの読み込みに失敗しました。" & vbNewLine & Application.UserLibraryPath & "に存在する「VBAFormatter.Ini」を削除し、再度実行してみてください。" & aIniKey & vbNewLine & aIniSection
    End If
End Function

Public Sub SetIniValue(aIniKey As String, aIniSection As String, aValue As String)
    Dim wRet     As Long
    wRet = WritePrivateProfileString(aIniSection, aIniKey, aValue, ThisWorkbook.Path & "\" & INI_NAME)
End Sub

Public Function IsExistsIni() As Boolean
    IsExistsIni = Dir(ThisWorkbook.Path & "\" & INI_NAME) <> ""
End Function

Public Sub CreateIniFile()
    Dim wNo As Integer
    wNo = FreeFile
    Open ThisWorkbook.Path & "\" & INI_NAME For Output As #wNo
    Print #wNo, "[Info]"
    Print #wNo, "  This file is used by VBAFormatterAddIn"
    Print #wNo, "[OptFormat]"
    Print #wNo, "  Tab_Cnt=4"
    Print #wNo, "  AllModuleExec=True"
    Print #wNo, "  AsFormat=True"
    Print #wNo, "  CommentFormat=True"
    Print #wNo, "  CommentExec=True"
    Close #wNo
End Sub

Public Sub IniWrite()
    Call SetIniValue(INI_KEY_TAB_CNT, INI_SEC_OPT_FORMAT, FOption.TxtTabCnt.Text)
    Call SetIniValue(INI_KEY_ALL_MODULE_EXEC, INI_SEC_OPT_FORMAT, FOption.IsAllModuleExec.Value)
    Call SetIniValue(INI_KEY_AS_FORMAT, INI_SEC_OPT_FORMAT, FOption.IsAsFormat.Value)
    Call SetIniValue(INI_KEY_COMENT_FORMAT, INI_SEC_OPT_FORMAT, FOption.IsCommentFormat.Value)
    Call SetIniValue(INI_KEY_COMENT_EXEC, INI_SEC_OPT_FORMAT, FOption.IsCommentExec.Value)
End Sub

Public Sub IniRead()
    With cIniKeyList
        .aTabCnt = CInt(GetIniValue(INI_KEY_TAB_CNT, INI_SEC_OPT_FORMAT))
        .aIsAllModuleExec = CBool(GetIniValue(INI_KEY_ALL_MODULE_EXEC, INI_SEC_OPT_FORMAT))
        .aIsAsFormat = CBool(GetIniValue(INI_KEY_AS_FORMAT, INI_SEC_OPT_FORMAT))
        .aIsCommentFormat = CBool(GetIniValue(INI_KEY_COMENT_FORMAT, INI_SEC_OPT_FORMAT))
        .aIsCommentExec = CBool(GetIniValue(INI_KEY_COMENT_EXEC, INI_SEC_OPT_FORMAT))
    End With
End Sub
