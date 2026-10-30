Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 If Ledger Is Worksheets(1) Then Ledger.Move After:=Worksheets(2)
 Probe = Ledger.Name & ":" & Ledger.CodeName & ":" & CStr(Ledger.Tick()) & ":" & CStr(Ledger Is Worksheets(2))
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
Private count As Long
Public Function Tick() As Long
 count = count + 1
 Tick = count
End Function

