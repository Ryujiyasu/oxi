Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(ThisWorkbook.Tick())
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private count As Long
Public Function Tick() As Long
 count = count + 1
 Tick = count
End Function

