Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As IWorker, late As Object
Set worker = New WorkerClass
Set late = worker
late.Value = 7
Probe = CStr(worker.Value)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "ProcessorClass"
Public Sub Change(ByRef value As Variant)
value = 7
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Implements IWorker
Private number As Long
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
number = value + 100
End Property
Public Property Get Value() As Long
Value = number
End Property
Private Property Let IWorker_Value(ByVal value As Long)
number = value + 10
End Property
Private Property Get IWorker_Value() As Long
IWorker_Value = number
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "IWorker"
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
End Property
Public Property Get Value() As Long
End Property
