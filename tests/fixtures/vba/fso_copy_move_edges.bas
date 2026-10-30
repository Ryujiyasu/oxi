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
    FileAttributesIdentity = 2026102082
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
    Case 283
        Open "held.txt" For Binary Shared As #1
        Err.Clear
        FileCopy "held.txt", "held.txt"
    Case 284
        Open "held.txt" For Binary Shared As #1
        Err.Clear
        fs.CopyFile "held.txt", "held.txt"
    Case 285
        Open "held.txt" For Binary Shared As #1
        Err.Clear
        fs.MoveFile "held.txt", "held.txt"
    Case 286
        Open "held.txt" For Binary Shared As #1
        Err.Clear
        Name "held.txt" As "held.txt"
    Case 287
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        SetAttr "copied.txt", 1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt", False
    Case 288
        SetAttr "held.txt", 1
        Err.Clear
        fs.CopyFile "held.txt", "held.txt"
    Case 289
        fs.CopyFile "*.txt", "empty"
    Case 290
        fs.MoveFile "*.txt", "empty"
    Case 291
        FileCopy "held.txt", "empty\"
    Case 292
        FileCopy "held.txt", "empty"
    Case 293
        fs.CopyFile "held.txt", "held.txt", FlagsArg(False)
    Case 294
        Set held = fs.OpenTextFile("held.txt", 2)
        Err.Clear
        fs.MoveFile "held.txt", "held.txt"
    Case 295
        fs.CopyFile "missing.txt", "missing.txt"
    Case 296
        fs.CopyFile "held.txt", "HELD.TXT"
    Case 297
        fs.MoveFile "held.txt", "HELD.TXT"
    Case 298
        Name "held.txt" As "HELD.TXT"
    Case 299
        SetAttr "held.txt", 1
        Err.Clear
        fs.MoveFile "held.txt", "held.txt"
    Case 300
        SetAttr "held.txt", 1
        Err.Clear
        fs.CopyFile "held.txt", "held.txt", FlagsArg(False)
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

