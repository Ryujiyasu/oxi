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
    FileAttributesIdentity = 2026102081
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
    Case 247
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(False)
    Case 248
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(False)
    Case 249
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(True)
    Case 250
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(Empty)
    Case 251
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(Null)
    Case 252
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(CVErr(2015))
    Case 253
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg("bad")
    Case 254
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(1)
    Case 255
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        SetAttr "copied.txt", 1
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(True)
    Case 256
        SetAttr "held.txt", 1
        Err.Clear
        fs.CopyFile PathArg("held.txt"), PathArg("copied.txt"), FlagsArg(True)
    Case 257
        fs.CopyFile "held.txt", "held.txt"
    Case 258
        FileCopy "held.txt", "held.txt"
    Case 259
        fs.CopyFile "missing.txt", "copied.txt"
    Case 260
        fs.CopyFile "held.txt", "missing\copied.txt"
    Case 261
        FileCopy "held.txt", "missing\copied.txt"
    Case 262
        fs.CopyFile "held.txt", "empty"
    Case 263
        fs.CopyFile "held.txt", "empty\"
    Case 264
        fs.MoveFile "held.txt", "empty\"
    Case 265
        fs.MoveFile "held.txt", "held.txt"
    Case 266
        Name "held.txt" As "held.txt"
    Case 267
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        fs.MoveFile "held.txt", "copied.txt"
    Case 268
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        Name "held.txt" As "copied.txt"
    Case 269
        fs.MoveFile "missing.txt", "copied.txt"
    Case 270
        fs.MoveFile "held.txt", "missing\copied.txt"
    Case 271
        Name "held.txt" As "missing\copied.txt"
    Case 272
        fs.CopyFile "*.txt", "empty\"
    Case 273
        fs.CopyFile "*.zz", "empty\"
    Case 274
        fs.MoveFile "*.txt", "empty\"
    Case 275
        fs.CopyFile "held.txt", "empty\"
        Err.Clear
        fs.CopyFile "held.txt", "empty\", False
    Case 276
        SetAttr "held.txt", 1
        Err.Clear
        FileCopy "held.txt", "copied.txt"
    Case 277
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        SetAttr "copied.txt", 1
        Err.Clear
        FileCopy "held.txt", "copied.txt"
    Case 278
        SetAttr "held.txt", 1
        Err.Clear
        Name "held.txt" As "copied.txt"
    Case 279
        SetAttr "held.txt", 1
        Err.Clear
        fs.MoveFile "held.txt", "copied.txt"
    Case 280
        Open "held.txt" For Binary Access Write Lock Read As #1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt"
    Case 281
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        Open "copied.txt" For Input Shared As #1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt"
    Case 282
        FileCopy "held.txt", "copied.txt"
        Err.Clear
        Open "copied.txt" For Binary Shared As #1
        Err.Clear
        fs.CopyFile "held.txt", "copied.txt"
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

