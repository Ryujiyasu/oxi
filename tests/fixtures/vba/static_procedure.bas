Option Explicit
Private first As ProcBox
Private second As ProcBox
Public Sub Prime()
Set first = New ProcBox
Set second = New ProcBox
End Sub
Public Static Function StepStandard() As Long
Dim stored As Long
stored = stored + 1
StepStandard = stored
End Function
Public Function ReadBoth() As String
ReadBoth = CStr(StepStandard()) & "|" & CStr(first.StepValue()) & "|" & CStr(second.StepValue())
End Function
Public Sub ResetFirst()
Set first = Nothing
Set first = New ProcBox
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "ProcBox"
Public Static Function StepValue() As Long
Dim stored As Long
stored = stored + 1
StepValue = stored
End Function
