Attribute VB_Name = "DriverModule"
DefLng A-Z
Public Function Probe() As String
Bump amount
Probe = TypeName(amount) & ":" & amount
End Function
Private Sub Bump(ByRef value As Long)
value = value + 1
End Sub
