Attribute VB_Name = "DriverModule"
Option Explicit
Public texts() As String * 4
Public Log As String
Public worker As WorkerClass
Public Function Probe() As String
ReDim texts(0)
texts(0) = "abcd"
Log = ""
Set worker = New WorkerClass
On Error Resume Next
Selected.Change texts(0)
Probe = CStr(Err.Number) & ":" & CStr(UBound(texts)) & ":" & Log & ":[" & texts(0) & "]"
End Function
Public Property Get Selected() As WorkerClass
Log = Log & "R"
ReDim texts(0 To 1)
Log = Log & "G"
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByRef value As String)
DriverModule.Log = DriverModule.Log & "C[" & value & "]"
value = "xyz"
End Sub
