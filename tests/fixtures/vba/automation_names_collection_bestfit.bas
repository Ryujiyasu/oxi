VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Dim obj As New Collection
    Dim value As String
    obj.Add "one"
    On Error Resume Next
    value = CStr(CallByName(obj, "C" & ChrW(242) & "unt", VbMethod))
    Probe = CStr(Err.Number) & "|" & value
End Function
