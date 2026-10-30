Attribute VB_Name = "DriverModule"
Public Function Probe() As String
    Probe = CStr(VBA.Len(CLng(123)))
End Function
