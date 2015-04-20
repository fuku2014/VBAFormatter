VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FOption 
   Caption         =   "Option"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   OleObjectBlob   =   "FOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub OKBtn_Click()
    Const ERR_MSG1 As String = "有効な値は１～３２の整数です。"
    
    If Not IsNumeric(TxtTabCnt.Text) Or TxtTabCnt.Text = "" Then
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If TxtTabCnt.Text < 1 Or TxtTabCnt.Text > 32 Then
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniWrite
    Unload Me
End Sub

Private Sub TxtTabCnt_Change()
    If Len(TxtTabCnt.Text) = 0 Then
        Exit Sub
    End If
    If IsNumeric(Right(TxtTabCnt.Text, 1)) = True Then
        Exit Sub
    End If
    TxtTabCnt.Text = Left(TxtTabCnt.Text, Len(TxtTabCnt.Text) - 1)
End Sub

Private Sub TxtTabCnt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    With cIniKeyList
        TxtTabCnt.Text = CStr(.aTabCnt)
        IsAllModuleExec.Value = .aIsAllModuleExec
        IsAsFormat.Value = .aIsAsFormat
        IsCommentFormat.Value = .aIsCommentFormat
        IsCommentExec.Value = .aIsCommentExec
    End With
End Sub
