Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Set Ledger.Held = Ledger.Range("A1")
 Ledger.Held.Value = 7
 Probe = TypeName(Ledger.Held) & ":" & CStr(Ledger.Held.Value)
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
Public Held As Object
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function

