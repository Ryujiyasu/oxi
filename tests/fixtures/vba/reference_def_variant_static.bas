Attribute VB_Name = "DriverModule"
Public Static Function Probe() As String
amount = amount + 1
Probe = TypeName(amount) & ":" & amount
End Function
