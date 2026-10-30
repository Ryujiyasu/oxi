Attribute VB_Name = "DriverModule"
Public Function Probe() As String
    Dim worker As New WorkerClass
    Dim number As Long
    number = 1
    worker.Change number
    Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByRef number As Long)
    number = number + 10
End Sub
