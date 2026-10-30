Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim actual As WorkerClass, first As IWorker, second As ISecond
Dim generic As Object, boxed As Variant, again As IWorker
Set actual = New WorkerClass
Set first = actual
Set second = actual
Set generic = first
Set boxed = second
Set again = generic
Probe = CStr(ObjPtr(actual) <> ObjPtr(first)) & ":" & CStr(ObjPtr(first) <> ObjPtr(second)) & ":" & CStr(ObjPtr(first) = ObjPtr(generic)) & ":" & CStr(ObjPtr(second) = ObjPtr(boxed)) & ":" & CStr(ObjPtr(again) = ObjPtr(first)) & ":" & CStr(actual.Pointer = ObjPtr(actual)) & ":" & CStr(first Is second)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "ProcessorClass"
Public Sub Change(ByRef value As Variant)
value = 7
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Implements IWorker
Implements ISecond
Public Function Pointer() As LongPtr
Pointer = ObjPtr(Me)
End Function
Private Sub ISecond_Touch()
End Sub
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
VERSION 1.0 CLASS
Attribute VB_Name = "ISecond"
Public Sub Touch()
End Sub
