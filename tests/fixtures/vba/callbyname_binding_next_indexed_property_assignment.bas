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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private held(1 To 2) As String
Public Property Let 陰(ByVal index As Long, ByVal value As String)
    held(index) = value
End Property
Public Property Get 陰(ByVal index As Long) As String
    陰 = held(index)
End Property

Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Dim obj As New Item
    Dim value As String
    On Error Resume Next
    CallByName obj, "餌", VbLet, 2, "indexed"
    value = CallByName(obj, "餌", VbGet, 2)
    Probe = CStr(Err.Number) & "|" & value
End Function
