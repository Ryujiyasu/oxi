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
    FileAttributesIdentity = 2026102074
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
    Case 30: SetAttr "empty", vbDirectory
    Case 31
        Set held = fs.GetFile("held.txt")
        held.Attributes = vbDirectory
    Case 32
        Set held = fs.GetFolder("empty")
        held.Attributes = vbDirectory
    Case 33: SetAttr "held.txt", 128
    Case 34
        Set held = fs.GetFile("held.txt")
        held.Attributes = 128
    Case 35
        Set held = fs.GetFolder("empty")
        held.Attributes = 0
    Case 36: SetAttr "missing.txt", vbDirectory
    Case 37: SetAttr Null, CVErr(2015)
    Case 38: SetAttr PathArg(root & "\held.txt"), FlagsArg(65536)
    Case 39: SetAttr PathArg(root & "\held.txt"), FlagsArg(1.5)
    Case 40: SetAttr PathArg(root & "\held.txt"), FlagsArg("1")
    Case 41: SetAttr PathArg(root & "\held.txt"), FlagsArg(Empty)
    Case 42: SetAttr PathArg(root & "\held.txt"), FlagsArg(True)
    Case 43
        Set held = fs.GetFile("held.txt")
        held.Attributes = -1
    Case 44
        Set held = fs.GetFolder("empty")
        held.Attributes = -1
    Case 45
        SetAttr "held.txt", vbReadOnly
        FileCopy "held.txt", "copy.txt"
    Case 46
        SetAttr "held.txt", vbReadOnly
        Name "held.txt" As "renamed.txt"
    Case 47
        SetAttr "held.txt", 0
        Open "held.txt" For Append As #1
        Print #1, "after"
        Close #1
    Case 48
        SetAttr "empty", vbReadOnly
        fs.DeleteFolder "empty", False
    Case 49
        SetAttr "held.txt", vbReadOnly
        SetAttr "held.txt", 0
        Set held = fs.CreateTextFile("held.txt", True)
        held.Close
    Case 50
        Open "held.txt" For Input As #1
        SetAttr "held.txt", vbReadOnly
    Case 51
        Open "held.txt" For Output As #1
        SetAttr "held.txt", vbReadOnly
        Print #1, "after"
    Case 52
        SetAttr "nonempty", vbReadOnly
        fs.DeleteFolder "nonempty", False
    Case 53
        SetAttr "nonempty", vbReadOnly
        fs.DeleteFolder "nonempty", True
    Case 54
        SetAttr "nonempty\child", vbReadOnly
        fs.DeleteFolder "nonempty", False
    Case 55
        SetAttr "held.txt", vbReadOnly
        Kill "held.txt"
    Case 56
        SetAttr "held.txt", vbReadOnly
        FileCopy "nonempty\leaf.txt", "held.txt"
    Case 57
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Binary As #1
    Case 58
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Append As #1
    Case 59
        Open "held.txt" For Input As #1
        SetAttr "held.txt", 0
    Case 60
        Open "held.txt" For Append As #1
        SetAttr "held.txt", vbReadOnly
    Case 61
        SetAttr "held.txt", vbReadOnly
        fs.DeleteFile "held.txt", True
    Case 62
        Open "held.txt" For Output As #1
        Set held = fs.GetFile("held.txt")
        held.Attributes = vbReadOnly
    Case 63
        Open "held.txt" For Input As #1
        Set held = fs.GetFile("held.txt")
        held.Attributes = vbReadOnly
    Case 64
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Binary As #1
        result = "X"
        Put #1, 1, result
    Case 65
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Random As #1 Len = 4
        result = "X"
        Put #1, 1, result
    Case 66
        Open "held.txt" For Binary Access Read As #1
        result = "X"
        Put #1, 1, result
    Case 67
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Binary Access Read Write As #1
    Case 68
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Binary Access Write As #1
    Case 69
        SetAttr "held.txt", vbReadOnly
        Open "held.txt" For Random Access Read As #1 Len = 4
        result = "X"
        Put #1, 1, result
    Case 70
        SetAttr "held.txt", vbReadOnly
        Set held = fs.GetFile("held.txt")
        held.Delete False
    Case 71
        SetAttr "held.txt", vbReadOnly
        Set held = fs.GetFile("held.txt")
        held.Delete True
    Case 72
        SetAttr "empty", vbReadOnly
        Set held = fs.GetFolder("empty")
        held.Delete False
    Case 73
        SetAttr "empty", vbReadOnly
        Set held = fs.GetFolder("empty")
        held.Delete True
    Case 74
        SetAttr "empty", vbReadOnly
        RmDir "empty"
    Case 75
        SetAttr "held.txt", vbHidden
        result = Dir("held.txt")
    Case 76
        SetAttr "held.txt", vbHidden
        result = Dir("held.txt", vbHidden)
    Case 77
        SetAttr "held.txt", vbSystem
        result = Dir("held.txt")
    Case 78
        SetAttr "held.txt", vbSystem
        result = Dir("held.txt", vbSystem)
    Case 79
        SetAttr "held.txt", vbReadOnly
        result = Dir("held.txt")
    Case 80
        SetAttr "held.txt", 0
        result = Dir("held.txt", vbReadOnly)
    Case 81
        result = CStr(VarType(GetAttr("held.txt"))) & ":" & CStr(VarType(fs.GetFile("held.txt").Attributes))
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    If kind = 26 Or kind = 50 Or kind = 51 Or (kind >= 57 And kind <= 60) Or kind = 62 Or kind = 63 Or (kind >= 64 And kind <= 69) Then Close #1
    Err.Clear
    fileAttrs = CStr(GetAttr(root & "\held.txt"))
    folderAttrs = CStr(GetAttr(root & "\empty"))
    fsFileAttrs = CStr(fs.GetFile(root & "\held.txt").Attributes)
    fsFolderAttrs = CStr(fs.GetFolder(root & "\empty").Attributes)
    snapshotError = Err.Number
    FileAttributesProbe = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|trace:" & trace & "|get-file:" & fileAttrs & "|get-folder:" & folderAttrs & "|fso-file:" & fsFileAttrs & "|fso-folder:" & fsFolderAttrs & "|nonempty:" & CStr(fs.FolderExists(root & "\nonempty")) & "|leaf:" & CStr(fs.FileExists(root & "\nonempty\leaf.txt")) & "|snapshot-error:" & CStr(snapshotError)
    On Error GoTo 0
End Function

