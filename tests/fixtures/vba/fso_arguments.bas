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
    FileAttributesIdentity = 2026102083
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
Private Function CollectDir(ByVal pattern As Variant, ByVal flags As Variant) As String
    Dim i As Long, item As String, number As Long
    On Error Resume Next
    For i = 0 To 11
        Err.Clear
        If i = 0 Then
            item = Dir(pattern, flags)
        Else
            item = Dir()
        End If
        number = Err.Number
        CollectDir = CollectDir & CStr(i) & ":" & item & ":" & CStr(number) & ";"
    Next i
    Err.Clear
    On Error GoTo 0
End Function
Public Function FileAttributesProbe(ByVal kind As Long, ByVal root As String) As String
    Dim marker As Byte, held2 As Object
    Dim pathValue As Variant, pathNumber As Long, emptyString As String
    Dim fs As Object, held As Object, result As String, fileAttrs As String, folderAttrs As String
    Dim fsFileAttrs As String, fsFolderAttrs As String, number As Long, origin As String, description As String, snapshotError As Long
    ChDrive Left$(root, 1)
    ChDir root
    Set fs = CreateObject("Scripting.FileSystemObject")
    trace = ""
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 301
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("True")
    Case 302
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("False")
    Case 303
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("0")
    Case 304
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("1")
    Case 305
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(" 0 ")
    Case 306
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("yes")
    Case 307
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(0.5)
    Case 308
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(DateSerial(2020, 1, 1))
    Case 309
        fs.CopyFile "held.txt", "copied.txt", Nothing
    Case 310
        fs.CopyFile "held.txt", "copied.txt", Array(1)
    Case 311
        fs.CopyFile "", "copied.txt", True
    Case 312
        fs.CopyFile Null, "copied.txt", True
    Case 313
        fs.CopyFile CVErr(2015), "copied.txt", True
    Case 314
        fs.CopyFile Empty, "copied.txt", True
    Case 315
        fs.CopyFile "missing.txt", "copied.txt", Null
    Case 316
        fs.CopyFile "held.txt", Null, True
    Case 317
        fs.CopyFile "held.txt", Empty, True
    Case 318
        fs.CopyFile "held.txt", CVErr(2015), True
    Case 319
        fs.CopyFile "missing.txt", Null, True
    Case 320
        fs.CopyFile "held.txt", "copied.txt", True, False
    Case 321
        fs.CopyFile "held.txt"
    Case 322
        fs.MoveFile "held.txt", "copied.txt", True
    Case 323
        fs.MoveFile Null, "copied.txt"
    Case 324
        fs.MoveFile "held.txt", Null
    Case 325
        fs.CopyFile "missing.txt", "missing\copied.txt", False
    Case 326
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        SetAttr "copied.txt", 1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt", Null
    Case 327
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        SetAttr "copied.txt", 1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt", "False"
    Case 328
        fs.CopyFile "held.txt", "empty", False
    Case 329
        fs.CopyFile "empty", "copied.txt"
    Case 330
        fs.MoveFile "empty", "copied.txt"
    Case 331
        FileCopy Null, "copied.txt"
    Case 332
        Name Null As "copied.txt"
    Case 333
        FileCopy CVErr(2015), "copied.txt"
    Case 334
        Name CVErr(2015) As "copied.txt"
    Case 335
        fs.CopyFile ErrorPath("held.txt"), PathArg("copied.txt"), FlagsArg(True)
    Case 336
        fs.CopyFile PathArg("held.txt"), ErrorPath("copied.txt"), FlagsArg(True)
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    If kind = 26 Or kind = 50 Or kind = 51 Or (kind >= 57 And kind <= 60) Or kind = 62 Or kind = 63 Or (kind >= 64 And kind <= 69) Then Close #1
    If Not held2 Is Nothing Then held2.Close
    If Not held Is Nothing Then held.Close
    Close #1, #2
    Err.Clear
    fileAttrs = CStr(GetAttr(root & "\held.txt"))
    folderAttrs = CStr(GetAttr(root & "\empty"))
    fsFileAttrs = CStr(fs.GetFile(root & "\held.txt").Attributes)
    fsFolderAttrs = CStr(fs.GetFolder(root & "\empty").Attributes)
    snapshotError = Err.Number
    FileAttributesProbe = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|trace:" & trace & "|get-file:" & fileAttrs & "|get-folder:" & folderAttrs & "|fso-file:" & fsFileAttrs & "|fso-folder:" & fsFolderAttrs & "|nonempty:" & CStr(fs.FolderExists(root & "\nonempty")) & "|leaf:" & CStr(fs.FileExists(root & "\nonempty\leaf.txt")) & "|snapshot-error:" & CStr(snapshotError)
    On Error GoTo 0
End Function

