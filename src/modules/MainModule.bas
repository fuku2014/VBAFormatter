Attribute VB_Name = "MainModule"

Option Explicit
Private tabCount         As Integer
Private cBeginList       As New Collection
Private cEndList         As New Collection
Private clsVBAFormatMenu As VBAFormatMenu
Private cBLList          As New Collection


Sub addButton()
    Dim wVBAFormatterMenu As CommandBarControl
    Dim wOptionMenu       As CommandBarControl
    Dim wFormatExecMenu   As CommandBarControl

    Set wVBAFormatterMenu = Application.VBE.CommandBars("メニュー バー").Controls.Add(Type:=msoControlPopup, ID:=1)

    wVBAFormatterMenu.Caption = "VBAFormatter(&Z)"

    Set wFormatExecMenu = wVBAFormatterMenu.Controls.Add(Type:=msoControlButton)

    Set wOptionMenu = wVBAFormatterMenu.Controls.Add(Type:=msoControlButton)
    wFormatExecMenu.Caption = "フォーマット実行(&F)..."

    wOptionMenu.Caption = "オプション(&O)..."
    Set clsVBAFormatMenu = New VBAFormatMenu
    Call clsVBAFormatMenu.InitializeInstance(wFormatExecMenu, wOptionMenu)
End Sub

Sub dellButton()
    Dim wCtrl As CommandBarControl
    Set clsVBAFormatMenu = Nothing
    For Each wCtrl In Application.VBE.CommandBars("メニュー バー").Controls
        If wCtrl.ID = 1 Then
            wCtrl.Delete
            Exit Sub
        End If
    Next
End Sub

Function indentEdit(prev, this) As String
    Dim space As String
    If (Trim(Left(this, InStr(this, " ")) = "End Select") Or this = "End Select" Or Trim(Left(this, InStr(InStr(this, " ") + 1, this, " "))) = "End Select") Then
        tabCount = tabCount - (2 * cIniKeyList.aTabCnt)
    ElseIf (isMemberOfCollection(cEndList, Trim(Left(this, InStr(this, " ")))) Or isMemberOfCollection(cEndList, this)) And this <> "End" Then
        tabCount = tabCount - cIniKeyList.aTabCnt
    End If
    If (Trim(Left(prev, InStr(prev, " ")) = "Select Case") Or prev = "Select Case" Or Trim(Left(prev, InStr(InStr(prev, " ") + 1, prev, " "))) = "Select Case") Then
        tabCount = tabCount + (2 * cIniKeyList.aTabCnt)
    ElseIf (isMemberOfCollection(cBeginList, Trim(Left(prev, InStr(prev, " ")))) Or isMemberOfCollection(cBeginList, prev) Or isMemberOfCollection(cBeginList, Trim(Left(prev, InStr(InStr(prev, " ") + 1, prev, " "))))) And (isOneLineCode(prev) = False) Then
    tabCount = tabCount + cIniKeyList.aTabCnt
End If
If tabCount < 0 Then
    tabCount = 0
End If
While Len(space) < tabCount
    space = space & " "
Wend
indentEdit = space & Trim(this)
End Function

Sub readPrintTxt(aCodeModule As CodeModule)
    Dim prevBuf  As Variant
    Dim thisBuf  As Variant
    Dim i        As Long
    Dim wNewThis As String
    prevBuf = ""
    thisBuf = ""
    For i = 1 To aCodeModule.CountOfLines
        If cIniKeyList.aIsCommentExec = True Or Left(Trim(aCodeModule.Lines(i, 1)), 1) <> "'" Then
            prevBuf = thisBuf
            thisBuf = aCodeModule.Lines(i, 1)
            wNewThis = indentEdit(Trim(prevBuf), Trim(thisBuf))
            aCodeModule.ReplaceLine i, wNewThis
            If Trim(prevBuf) = "" Then
                cBLList.Add New Dictionary
                cBLList.item(cBLList.Count).Add i, wNewThis
            Else
                cBLList.item(cBLList.Count).Add i, wNewThis
            End If
        End If
    Next i
End Sub
Sub init()
    Set cBeginList = New Collection
    Set cEndList = New Collection
    Set cBLList = New Collection
    tabCount = 0
    cBeginList.Add "If"
    cBeginList.Add "Else"
    cBeginList.Add "ElseIf"
    cBeginList.Add "Sub"
    cBeginList.Add "With"
    cBeginList.Add "While"
    cBeginList.Add "For"
    cBeginList.Add "Do"
    cBeginList.Add "Function"
    cBeginList.Add "Public Function"
    cBeginList.Add "Private Function"
    cBeginList.Add "Public Sub"
    cBeginList.Add "Private Sub"
    cBeginList.Add "Property"
    cBeginList.Add "Type"
    cBeginList.Add "Private Type"
    cBeginList.Add "Public Type"
    cBeginList.Add "Public Property"
    cBeginList.Add "Public Enum"
    cBeginList.Add "Case"
    cEndList.Add "End"
    cEndList.Add "Next"
    cEndList.Add "Loop"
    cEndList.Add "Else"
    cEndList.Add "Wend"
    cEndList.Add "ElseIf"
    cEndList.Add "Case"
End Sub

Function isMemberOfCollection(col As Collection, query As Variant) As Boolean
    Dim item
    For Each item In col
        If item = query Then
            isMemberOfCollection = True
            Exit Function
        End If
    Next
    isMemberOfCollection = False
End Function

Function isOneLineCode(str As Variant) As Boolean
Dim buff As Variant
    isOneLineCode = False
    If InStr(str, "'") = 0 Then
    If (InStr(str, "Then") <> 0 And InStr(str, "Then") + 3 < Len(str)) Or InStr(str, "End Function") <> 0 Or InStr(str, "End Sub") Or InStr(str, "End Property") <> 0 Or InStr(str, "End If") <> 0 Then
      isOneLineCode = True
    End If
    Else
    buff = Trim(StrConv(LeftB(StrConv(str, vbFromUnicode), Instr2(1, str, "'") - 1), vbUnicode))
        If (InStr(buff, "Then") <> 0 And InStr(buff, "Then") + 3 < Len(buff)) Or InStr(buff, "End Function") <> 0 Or InStr(buff, "End Sub") Or InStr(str, "End Property") <> 0 Or InStr(buff, "End If") <> 0 Then
      isOneLineCode = True
    End If
    End If
End Function

Sub FormatExecMain()
    Dim mVBComp As VBComponent
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    Call init

    If cIniKeyList.aIsAllModuleExec Then
        For Each mVBComp In ActiveWorkbook.VBProject.VBComponents
            Call Exec(mVBComp.CodeModule)
            Set cBLList = New Collection
            tabCount = 0
        Next mVBComp

    Else
        Call Exec(ActiveWorkbook.Application.VBE.SelectedVBComponent.CodeModule)
    End If
End Sub

Sub Exec(aCodeModule As CodeModule)

    Call readPrintTxt(aCodeModule)

    If cIniKeyList.aIsAsFormat Then
        Call FixAs(aCodeModule)
    End If

    If cIniKeyList.aIsCommentFormat Then
        Call FixCom(aCodeModule)
    End If
End Sub

Sub OptionMain()
    FOption.Show
End Sub


Sub FixAs(aCodeModule As CodeModule)
    Dim i, j  As Integer
    Dim wDic  As Dictionary
    Dim wKeys
    Dim wKey  As Variant
    Dim wStr  As Variant
    Dim wMax  As Integer
    For i = 1 To cBLList.Count
        wMax = 0
        Set wDic = cBLList(i)
        wKeys = wDic.Keys

        For Each wKey In wKeys
            wStr = wDic(wKey)
            If (InStrRev(wStr, """", InStrRev(wStr, " As ") + 1) = 0) And Instr2(1, wStr, " As ") > wMax And (Left(Trim(wStr), 4) = "Dim " Or Left(Trim(wStr), 6) = "Const " Or (aCodeModule.CountOfDeclarationLines >= wKey)) Then
                wMax = Instr2(1, wStr, " As ")
            End If
        Next

        For Each wKey In wKeys
            wStr = wDic(wKey)
            If (InStrRev(wStr, """", InStrRev(wStr, " As ") + 1) = 0) And InStr(wStr, " As ") > 0 And (Left(Trim(wStr), 4) = "Dim " Or Left(Trim(wStr), 6) = "Const " Or (aCodeModule.CountOfDeclarationLines >= wKey)) Then
                wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, " As ")), vbUnicode) & WorksheetFunction.Rept(" ", wMax - Instr2(1, wStr, " As ")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, " As ")), vbUnicode)
                aCodeModule.ReplaceLine wKey, wStr
            End If
            wDic(wKey) = wStr
        Next
    Next i
End Sub

Sub FixCom(aCodeModule As CodeModule)
    Dim i, j    As Integer
    Dim wDic    As Dictionary
    Dim wKeys
    Dim wKey    As Variant
    Dim wStr    As Variant
    Dim wMax    As Integer
    Dim tempStr As Variant
    For i = 1 To cBLList.Count
        wMax = 0
        Set wDic = cBLList(i)
        wKeys = wDic.Keys

        For Each wKey In wKeys
            wStr = wDic(wKey)
            tempStr = wStr
            While InStr(tempStr, """") > 1
                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
            Wend
            If (Left(Trim(wStr), 1) <> "'") And (InStr(tempStr, "'") > 0) And (Instr2(1, wStr, "'") > wMax) Then
                wMax = Instr2(1, wStr, "'")
            End If
        Next

        For Each wKey In wKeys
            wStr = wDic(wKey)
            tempStr = wStr
            While InStr(tempStr, """") > 1
                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
            Wend
            If (Left(Trim(wStr), 1) <> "'") And (InStr(tempStr, "'") > 0) Then
                wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, "'") - 1), vbUnicode) & WorksheetFunction.Rept(" ", wMax - Instr2(1, wStr, "'")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, "'") + 1), vbUnicode)
                aCodeModule.ReplaceLine wKey, wStr
            End If
        Next
    Next i
End Sub

Function Instr2(aStart As Integer, aString1 As Variant, aString2 As String) As Long
    Instr2 = InStrB(aStart, StrConv(aString1, vbFromUnicode), StrConv(aString2, vbFromUnicode))
End Function

