VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

' SPDX-License-Identifier: CC-BY-SA-4.0

Attribute VB_Name = "SaveProbe"
Option Explicit
Private hits As Long
Public Function DirectoryLifetimeIdentity() As Long
    DirectoryLifetimeIdentity = 2026102041
End Function
Public Function DirectoryEnvironment() As Variant
    DirectoryEnvironment = Array(Left$(CurDir$, 1), CurDir$("C"), CurDir$("D"))
End Function
Public Function DirectoryRead() As String
    hits = hits + 1
    DirectoryRead = "plain:" & CurDir$ & "|qualified:" & VBA.CurDir$ & "|C:" & VBA.CurDir$("C") & "|D:" & CurDir$("D") & "|hits:" & CStr(hits)
End Function
Public Function DirectorySet(ByVal folder As String) As String
    ChDrive Left$(folder, 1)
    ChDir folder
    DirectorySet = DirectoryRead()
End Function
Public Function DirectoryRemember(ByVal folder As String) As String
    ChDrive "D"
    ChDir folder
    DirectoryRemember = DirectoryRead()
End Function
Public Function DirectoryReturnC() As String
    ChDrive "C"
    DirectoryReturnC = DirectoryRead()
End Function
Public Sub DirectoryStop(ByVal folder As String)
    ChDrive Left$(folder, 1)
    ChDir folder
    End
End Sub

