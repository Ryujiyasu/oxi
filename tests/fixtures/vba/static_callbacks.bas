Option Explicit
Private first As CallbackBox
Private second As CallbackBox
Public Sub Prime()
Set first = New CallbackBox
Set second = New CallbackBox
End Sub
Public Function SharedNext(Optional ByVal reenter As Boolean = True) As Long
Static stored As Long
stored = stored + 1
If reenter Then Call first.StepOnce(False)
SharedNext = stored
End Function
Public Function ReadBoth() As String
Dim one As Long, two As Long, three As Long, four As Long
one = SharedNext(True)
two = first.StepOnce(True)
three = second.StepOnce(True)
four = SharedNext(False)
ReadBoth = CStr(one) & "|" & CStr(two) & "|" & CStr(three) & "|" & CStr(four)
End Function
Public Sub ResetFirst()
Set first = Nothing
Set first = New CallbackBox
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CallbackBox"
Public Function StepOnce(Optional ByVal reenter As Boolean = True) As Long
Static stored As Long
stored = stored + 1
If reenter Then Call SharedNext(False)
StepOnce = stored
End Function
