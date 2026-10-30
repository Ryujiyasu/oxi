Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 If Ledger.Index = 1 Then Ledger.Move After:=Journal
 Probe = CStr(Ledger.Tick()) & ":" & CStr(Journal.Tick()) & ":" & CStr(ArchiveHost.Tick()) & ":" & CStr(Ledger.Index) & ":" & CStr(Journal.Index) & ":" & CStr(Ledger Is Worksheets(Ledger.Index))
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Journal"
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArchiveHost"
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

