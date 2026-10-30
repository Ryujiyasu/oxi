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
    FileAttributesIdentity = 2026102084
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
    Case 337
        FileCopy FlagsArg(Null), PathArg("copied.txt")
    Case 338
        Name FlagsArg(Null) As PathArg("copied.txt")
    Case 339
        FileCopy FlagsArg(CVErr(2015)), PathArg("copied.txt")
    Case 340
        Name FlagsArg(CVErr(2015)) As PathArg("copied.txt")
    Case 341
        FileCopy FlagsArg(Null), ErrorPath("copied.txt")
    Case 342
        Name FlagsArg(Null) As ErrorPath("copied.txt")
    Case 343
        FileCopy ErrorPath("held.txt"), PathArg("copied.txt")
    Case 344
        Name ErrorPath("held.txt") As PathArg("copied.txt")
    Case 345
        FileCopy PathArg("held.txt"), FlagsArg(Null)
    Case 346
        Name PathArg("held.txt") As FlagsArg(Null)
    Case 347
        FileCopy FlagsArg(Null), FlagsArg(CVErr(2015))
    Case 348
        Name FlagsArg(Null) As FlagsArg(CVErr(2015))
    Case 349
        FileCopy FlagsArg(CVErr(2015)), FlagsArg(Null)
    Case 350
        Name FlagsArg(CVErr(2015)) As FlagsArg(Null)
    Case 351
        FileCopy FlagsArg(Null), FlagsArg(Null)
    Case 352
        Name FlagsArg(Null) As FlagsArg(Null)
    Case 353
        fs.CopyFile FlagsArg(Null), PathArg("copied.txt"), FlagsArg(True)
    Case 354
        fs.MoveFile FlagsArg(Null), PathArg("copied.txt")
    Case 355
        fs.CopyFile FlagsArg(Null), ErrorPath("copied.txt"), FlagsArg(True)
    Case 356
        fs.MoveFile FlagsArg(Null), ErrorPath("copied.txt")
    Case 357
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(True), ErrorPath("extra")
    Case 358
        fs.MoveFile PathArg("held.txt"), PathArg("copied.txt"), ErrorPath("extra")
    Case 359
        FileCopy FlagsArg(Empty), PathArg("copied.txt")
    Case 360
        Name FlagsArg(Empty) As PathArg("copied.txt")
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

