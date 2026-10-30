Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = TypeName(ThisWorkbook) & ":" & CStr(ThisWorkbook Is ActiveWorkbook) & ":" & CStr(ObjPtr(ThisWorkbook) = ObjPtr(ActiveWorkbook))
End Function
