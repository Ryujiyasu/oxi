Attribute VB_Name = "DriverModule"
DefLng A-Z
Public Function Probe() As String
Dim value As Variant
value = 123
Probe = Len(value) & ":" & Len("123")
End Function
