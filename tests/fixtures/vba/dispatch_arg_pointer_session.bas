Attribute VB_Name = "DriverModule"
Option Explicit
Public held As IWorker
Public previous As LongPtr
Public Function Probe() As String
Dim worker As IWorker, late As WorkerClass
If held Is Nothing Then Set held = New WorkerClass
Set worker = held
Set late = worker

If previous = 0 Then previous = ObjPtr(worker)
Probe = CStr(previous = ObjPtr(worker)) & ":" & CStr(ObjPtr(worker) <> 0) & ":" & CStr(ObjPtr(Nothing) = 0)
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
