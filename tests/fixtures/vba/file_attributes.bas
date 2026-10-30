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
Private trace As String
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102071
End Function
Public Function DirectoryEnvironment() As Variant
    DirectoryEnvironment = Array(Left$(CurDir$, 1), CurDir$("C"), CurDir$("D"))
End Function
Private Function PathArg(ByVal path As String) As String
    trace = trace & "P"
    PathArg = path
End Function
Private Function FlagsArg(ByVal flags As Variant) As Variant
    trace = trace & "A"
    FlagsArg = flags
End Function
Private Function ErrorPath(ByVal path As String) As String
    trace = trace & "E"
    Err.Raise 6
End Function
Public Function FileAttributesProbe(ByVal kind As Long, ByVal root As String) As String
    Dim fs As Object, held As Object, result As String, fileAttrs As String, folderAttrs As String
    Dim fsFileAttrs As String, fsFolderAttrs As String, number As Long, origin As String, description As String, snapshotError As Long
    ChDrive Left$(root, 1)
    ChDir root
    Set fs = CreateObject("Scripting.FileSystemObject")
    trace = ""
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 0: result = CStr(GetAttr("held.txt"))
    Case 1: SetAttr "held.txt", 0
    Case 2: SetAttr "held.txt", vbReadOnly
    Case 3: SetAttr "held.txt", vbHidden
    Case 4: SetAttr "held.txt", vbSystem
    Case 5: SetAttr "held.txt", vbArchive
    Case 6: SetAttr "held.txt", vbReadOnly Or vbHidden
    Case 7: SetAttr "empty", 0
    Case 8: SetAttr "empty", vbReadOnly
    Case 9: SetAttr "empty", vbHidden
    Case 10: SetAttr "empty", vbArchive
    Case 11
        Set held = fs.GetFile("held.txt")
        held.Attributes = 0
    Case 12
        Set held = fs.GetFile("held.txt")
        held.Attributes = vbReadOnly
    Case 13
        Set held = fs.GetFolder("empty")
        held.Attributes = vbReadOnly
    Case 14: SetAttr "missing.txt", vbReadOnly
    Case 15: SetAttr "absent\missing.txt", vbReadOnly
    Case 16: result = CStr(GetAttr("missing.txt"))
    Case 17: SetAttr "held.txt", vbDirectory
    Case 18: SetAttr "held.txt", vbAlias
    Case 19: SetAttr "held.txt", -1
    Case 20: SetAttr PathArg(root & "\held.txt"), FlagsArg(3)
    Case 21: SetAttr PathArg(root & "\held.txt"), CVErr(2015)
    Case 22: SetAttr ErrorPath(root & "\held.txt"), FlagsArg(3)
    Case 23: SetAttr PathArg(root & "\missing.txt"), FlagsArg(3)
    Case 24: SetAttr "held.txt", Null
    Case 25: SetAttr Null, FlagsArg(3)
    Case 26
        SetAttr "held.txt", vbReadOnly
        Err.Clear
        Open "held.txt" For Output As #1
    Case 27
        SetAttr "held.txt", vbReadOnly
        Err.Clear
        Set held = fs.CreateTextFile("held.txt", True)
    Case 28
        SetAttr "nonempty\leaf.txt", vbReadOnly
        Err.Clear
        fs.DeleteFolder "nonempty", False
    Case 29
        SetAttr "nonempty\leaf.txt", vbReadOnly
        Err.Clear
        fs.DeleteFolder "nonempty", True
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    If kind = 26 Then Close #1
    Err.Clear
    fileAttrs = CStr(GetAttr(root & "\held.txt"))
    folderAttrs = CStr(GetAttr(root & "\empty"))
    fsFileAttrs = CStr(fs.GetFile(root & "\held.txt").Attributes)
    fsFolderAttrs = CStr(fs.GetFolder(root & "\empty").Attributes)
    snapshotError = Err.Number
    FileAttributesProbe = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|trace:" & trace & "|get-file:" & fileAttrs & "|get-folder:" & folderAttrs & "|fso-file:" & fsFileAttrs & "|fso-folder:" & fsFolderAttrs & "|nonempty:" & CStr(fs.FolderExists(root & "\nonempty")) & "|leaf:" & CStr(fs.FileExists(root & "\nonempty\leaf.txt")) & "|snapshot-error:" & CStr(snapshotError)
    On Error GoTo 0
End Function

