Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = Ledger.Inside()
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
Public Function Inside() As String
 Cells(1, 1).Value = 7
 Inside = CStr(Me.Range("A1").Value) & ":" & CStr(Worksheets(2).Range("A1").Value)
End Function

