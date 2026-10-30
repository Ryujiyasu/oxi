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

Attribute VB_Name = "売上"
Option Explicit
Public Function Label() As String
    Label = "売上"
End Function

Attribute VB_Name = "利益"
Option Explicit
Public Function Label() As String
    Label = "利益"
End Function

Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Probe = 売上.Label() & "|" & 利益.Label()
End Function
