VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAFormatMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cFormatExecEvent As CommandBarEvents
Attribute cFormatExecEvent.VB_VarHelpID = -1
Private WithEvents cOptionEvent     As CommandBarEvents
Attribute cOptionEvent.VB_VarHelpID = -1

Public Sub InitializeInstance(ByVal aFormatExecEvent As CommandBarControl, ByVal aOptionEvent As CommandBarControl)
    Set cFormatExecEvent = Application.VBE.Events.CommandBarEvents(aFormatExecEvent)
    Set cOptionEvent = Application.VBE.Events.CommandBarEvents(aOptionEvent)
End Sub

Private Sub cFormatExecEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo cFormatExecEvent_Click_err
    Call FormatExecMain
    Exit Sub
cFormatExecEvent_Click_err:
    MsgBox Err.Description
End Sub

Private Sub cOptionEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo cOptionEvent_Click_err
    Call OptionMain
    Exit Sub
cOptionEvent_Click_err:
    MsgBox Err.Description
End Sub
