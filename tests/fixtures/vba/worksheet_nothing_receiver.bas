VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "SaveProbe"
' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private receiverCalls As Long
Public Function WidthProbeIdentity() As Long
    WidthProbeIdentity = 2026102019
End Function
Private Function ReceiverRange() As Object
    receiverCalls = receiverCalls + 1
    Set ReceiverRange = Columns("D")
End Function
Private Function ReceiverSheet() As Object
    receiverCalls = receiverCalls + 1
    Set ReceiverSheet = ActiveSheet
End Function
Public Function NothingReceiver(ByVal mode As Long) As String
    Dim value As Variant, failure As Long
    receiverCalls = 0
    Set value = Nothing
    On Error Resume Next
    Err.Clear
    Select Case mode
    Case 0: ReceiverRange().ColumnWidth = value
    Case 1: ReceiverRange().UseStandardWidth = value
    Case 2: ReceiverSheet().StandardWidth = value
    Case 3: ReceiverRange().RowHeight = value
    Case 4: ReceiverRange().Value2 = value
    Case 5: ReceiverRange().Hidden = value
    Case 6: CallByName ReceiverRange(), "ColumnWidth", VbLet, value
    Case 7: CallByName ReceiverSheet(), "StandardWidth", VbLet, value
    Case 8: CallByName ReceiverRange(), "UseStandardWidth", VbLet, value
    Case 9: CallByName ReceiverRange(), "RowHeight", VbLet, value
    Case 10
        With ReceiverRange()
            .ColumnWidth = value
        End With
    Case 11
        With ReceiverSheet()
            .StandardWidth = value
        End With
    End Select
    failure = Err.Number
    On Error GoTo 0
    NothingReceiver = CStr(failure) & "|" & CStr(receiverCalls)
End Function
