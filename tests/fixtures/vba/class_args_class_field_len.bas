Attribute VB_Name = "DriverModule"
Public Function Probe() As String
    Dim worker As New WorkerClass
    Dim result As Variant
    worker.number = 123
    result = Len(worker.number)
    Probe = TypeName(result) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
