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
Public Function DirectoryObjectIdentity() As Long
    DirectoryObjectIdentity = 2026102039
End Function
Public Function DirectoryObjectProbe(ByVal kind As Long, ByVal folder As String) As String
    Dim fs As Object, held As Object, result As String, number As Long, origin As String, description As String
    ChDrive Left$(folder, 1)
    ChDir folder
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 0: result = fs.GetAbsolutePathName("held.txt")
    Case 1: result = fs.GetAbsolutePathName(".\held.txt")
    Case 2: result = fs.GetAbsolutePathName("..\parent.txt")
    Case 3: result = fs.GetAbsolutePathName(Mid$(folder, 3) & "\rooted.txt")
    Case 4: result = fs.GetAbsolutePathName(Left$(folder, 2) & "held.txt")
    Case 5: result = fs.GetAbsolutePathName("")
    Case 6
        Set held = fs.GetFile("held.txt")
        result = held.Path & "|" & held.ParentFolder.Path & "|" & CStr(held.Size)
    Case 7
        Set held = fs.GetFolder(".")
        result = held.Path & "|" & held.Name
    Case 8
        Set held = fs.GetFile("held.txt")
        ChDir "sub"
        result = held.Path & "|" & CStr(held.Size) & "|" & CStr(fs.GetFile("held.txt").Size)
    Case 9
        Set held = fs.GetFolder(".")
        ChDir "sub"
        result = held.Path & "|" & held.Name
    Case 10
        Set held = fs.GetFile("..\work\held.txt")
        result = held.Path & "|" & CStr(held.Size)
    Case 11: result = fs.BuildPath("relative", "leaf.txt")
    Case 12: result = fs.GetFolder("sub\..").Path
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    On Error GoTo 0
    DirectoryObjectProbe = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|cwd:" & CurDir$
End Function

