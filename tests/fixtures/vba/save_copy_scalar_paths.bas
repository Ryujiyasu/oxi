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
Public Function ScalarPathIdentity() As Long
    ScalarPathIdentity = 2026102037
End Function
Public Function ScalarCopy(ByVal kind As Long, ByVal folder As String) As String
    Dim candidate As Variant, number As Long, origin As String, description As String, help As String, context As Long
    ChDrive Left$(folder, 1)
    ChDir folder
    Select Case kind
    Case 0: candidate = CByte(7)
    Case 1: candidate = CInt(-8)
    Case 2: candidate = CLng(9)
    Case 3: candidate = CSng(1.25)
    Case 4: candidate = CDbl(2.5)
    Case 5: candidate = CCur(3.75)
    Case 6: candidate = CDec(4.125)
    Case 7: candidate = True
    Case 8: candidate = False
    Case 9: candidate = DateSerial(2026, 10, 20)
    Case 10: candidate = TimeSerial(12, 34, 56)
    Case 11: candidate = "relative.xlsm"
    Case 12: candidate = ".\dot.xlsm"
    Case 13: candidate = "sub\child.xlsm"
    Case 14: candidate = "..\parent.xlsm"
    Case 15: candidate = Left$(folder, 2) & "drive.xlsm"
    Case 16: candidate = Mid$(folder, 3) & "\rooted.xlsm"
    Case 17: candidate = Empty
    Case 18: candidate = "space name.xlsm"
    Case 19: candidate = "ends-with-space.xlsm "
    Case 20: candidate = CDate(0)
    Case 21: candidate = CLngLng(10000000001#)
    Case 22: candidate = vbNullString
    End Select
    ThisWorkbook.Saved = True
    On Error Resume Next
    Err.Clear
    ThisWorkbook.SaveCopyAs Filename:=candidate
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    help = Err.HelpFile
    context = Err.HelpContext
    On Error GoTo 0
    ScalarCopy = CStr(number) & "|" & origin & "|" & description & "|" & help & "|" & CStr(context) & "|saved:" & CStr(ThisWorkbook.Saved) & "|kind:" & TypeName(candidate) & "|cwd:" & CurDir$
End Function
Public Function DirectoryProbe(ByVal kind As Long, ByVal folder As String) As String
    Dim number As Long, origin As String, description As String, help As String, context As Long, queried As String
    ChDrive Left$(folder, 1)
    ChDir folder
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 0: ChDir folder & "\sub"
    Case 1: ChDir "sub"
    Case 2: ChDir "missing"
    Case 3: ChDir ".."
    Case 4: ChDir "."
    Case 5: ChDir Replace(folder & "\sub", "\", "/")
    Case 6: ChDir Left$(folder, 2) & "sub"
    Case 7: ChDrive Left$(folder, 1)
    Case 8: ChDrive Left$(folder, 2)
    Case 9: ChDrive "?"
    Case 10: ChDir ""
    Case 11: ChDir Null
    Case 12: ChDir ThisWorkbook.FullName
    Case 13: ChDrive Left$(folder, 2) & "\nonsense"
    Case 14: ChDrive ""
    Case 15: ChDrive Null
    Case 16: queried = CurDir$("?")
    Case 17: queried = CurDir$(Left$(folder, 1))
    Case 18: queried = CurDir$(Left$(folder, 2))
    Case 19: queried = CurDir$("")
    Case 20: queried = CurDir$(Null)
    Case 21: queried = CurDir$(3)
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    help = Err.HelpFile
    context = Err.HelpContext
    On Error GoTo 0
    DirectoryProbe = CStr(number) & "|" & origin & "|" & description & "|" & help & "|" & CStr(context) & "|cwd:" & CurDir$ & "|query:" & queried
End Function

