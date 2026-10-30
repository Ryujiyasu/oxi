Attribute VB_Name = "DriverModule"
DefLng A-Z
Public Function Probe() As String
Probe = CStr(Len(amount)) & ":" & TypeName(amount)
End Function
