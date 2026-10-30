VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "日本語台帳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Function ReadText() As String
    ReadText = TypeName(Me) & "|" & Me.Name
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Probe = 日本語台帳.ReadText() & "|" & CStr(日本語台帳.Worksheets.Count)
End Function
