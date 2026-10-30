Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = PlotHost.Report()
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlotHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private count As Long
Public Function Report() As String
 count = count + 1
 Report = Me.Name & ":" & Me.CodeName & ":" & TypeName(Me) & ":" & CStr(Me.Index) & ":" & CStr(count)
End Function

