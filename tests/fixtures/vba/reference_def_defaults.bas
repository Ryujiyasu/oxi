Attribute VB_Name = "DriverModule"
DefLng A-A, C-D
DefStr S
DefBool B
Public Function Probe() As String
Probe = TypeName(amount) & ":" & amount & "|" & TypeName(beta) & ":" & CStr(beta) & "|" & TypeName(sunset) & ":" & Len(sunset) & "|" & TypeName(zero) & "|" & TypeName(counter%)
End Function
