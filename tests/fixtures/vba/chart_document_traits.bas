Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = TypeName(Worksheets) & "|" & TypeName(Sheets) & "|" & TypeName(Charts) & "|" & TypeName(ActiveSheet) & "|" & TypeName(PlotHost) & "|" & CStr(Worksheets(1).Index) & "|" & CStr(PlotHost.Index) & "|" & CStr(PlotHost Is Sheets(1)) & "|" & CStr(PlotHost Is Charts(1))
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
Public Function Tick() As Long
 count = count + 1
 Tick = count
End Function

