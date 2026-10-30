Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 On Error Resume Next
 Dim s As Object
 Set s = Ledger("A1")
 Probe = CStr(Err.Number) & ":" & TypeName(s)
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function

