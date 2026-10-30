Option Explicit
Private first As ScopeA
Private second As ScopeA
Private other As ScopeB
Private Function NextCount() As Long
Static stored As Long
stored = stored + 1
NextCount = stored
End Function
Public Sub Prime()
Set first = New ScopeA
Set second = New ScopeA
Set other = New ScopeB
End Sub
Public Function ReadAll() As String
ReadAll = CStr(NextCount()) & "|" & CStr(first.NextCount()) & "|" & CStr(second.NextCount()) & "|" & CStr(other.NextCount())
End Function
Public Sub ResetFirst()
Set first = Nothing
Set first = New ScopeA
End Sub

VERSION 1.0 CLASS
Attribute VB_Name = "ScopeA"
Public Function NextCount() As Long
Static stored As Long
stored = stored + 1
NextCount = stored
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "ScopeB"
Public Function NextCount() As Long
Static stored As Long
stored = stored + 1
NextCount = stored
End Function
