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
Attribute VB_Name = "売上表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Function ReadText() As String
    ReadText = TypeName(Me) & "|" & Me.Name
End Function

Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Probe = 売上表.ReadText() & "|" & 売上表.CodeName
End Function
