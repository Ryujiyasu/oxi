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
    FileAttributesIdentity = 2026102079
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
    Dim marker As Byte
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
    Case 184
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Lock #1, 1 To 2
        Open "held.txt" For Input Shared As #2
        Err.Clear
        Line Input #2, result
    Case 185
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Input Shared As #2
        Lock #1, 1 To 2
        Err.Clear
        Line Input #2, result
    Case 186
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Input Shared As #2
        result = Input$(1, #2)
        Lock #1, 1 To 2
        Err.Clear
        result = result & Input$(1, #2)
    Case 187
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Input Shared As #2
        Lock #1
        Err.Clear
        result = CStr(EOF(2))
    Case 188
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Lock #1, 64
        Open "held.txt" For Input Shared As #2
        Err.Clear
        Line Input #2, result
    Case 189
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Lock #1, 64
        Open "held.txt" For Input Shared As #2
        Err.Clear
        result = Input$(1, #2)
    Case 190
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Lock #1, 1 To 2
        Set held = fs.OpenTextFile("held.txt", 1)
        Err.Clear
        result = held.ReadLine
    Case 191
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Input Shared As #2
        Lock #1
        Err.Clear
        result = CStr(LOF(2)) & ":" & CStr(FileLen("held.txt"))
    Case 192
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Kill "held.txt"
    Case 193
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Name "held.txt" As "renamed.txt"
    Case 194
        marker = 90
        Open "held.txt" For Binary Shared As #1
        FileCopy "held.txt", "copied.txt"
    Case 195
        marker = 90
        Open "held.txt" For Binary Lock Read Write As #1
        FileCopy "held.txt", "copied.txt"
    Case 196
        marker = 90
        Open "held.txt" For Binary Shared As #1
        fs.CopyFile "held.txt", "copied.txt"
    Case 197
        marker = 90
        Open "held.txt" For Binary Shared As #1
        fs.DeleteFile "held.txt"
    Case 198
        marker = 90
        Open "held.txt" For Binary Shared As #1
        fs.MoveFile "held.txt", "renamed.txt"
    Case 199
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Set held = fs.OpenTextFile("held.txt", 2)
    Case 200
        marker = 90
        Open "held.txt" For Binary Lock Write As #1
        Set held = fs.OpenTextFile("held.txt", 2)
    Case 201
        marker = 90
        Open "held.txt" For Binary Lock Read As #1
        Set held = fs.OpenTextFile("held.txt", 1)
    Case 202
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Binary Shared As #2
        Lock #1, 1 To 2
        Err.Clear
        result = Input$(1, #2)
    Case 203
        marker = 90
        Open "held.txt" For Binary Shared As #1
        Open "held.txt" For Binary Shared As #2
        Lock #1, 1 To 2
        Err.Clear
        result = Input$(0, #2)
    Case 204
        marker = 90
        Open "held.txt" For Binary Access Read Shared As #1
        Lock #1, 1 To 2
    Case 205
        marker = 90
        Open "held.txt" For Binary Access Write Shared As #1
        Open "held.txt" For Binary Access Read Shared As #2
        Lock #1, 1 To 2
        Err.Clear
        Get #2, 1, marker
        result = CStr(marker) & ":seek2:" & CStr(Seek(2)) & ":loc2:" & CStr(Loc(2))
    Case 206
        marker = 90
        Open "held.txt" For Output Shared As #1
        Lock #1, 1 To 2
        Unlock #1, 1 To 2
    Case 207
        marker = 90
        Open "held.txt" For Input Shared As #1
        Lock #1, 1 To 2
        Unlock #1, 1 To 2
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    If kind = 26 Or kind = 50 Or kind = 51 Or (kind >= 57 And kind <= 60) Or kind = 62 Or kind = 63 Or (kind >= 64 And kind <= 69) Then Close #1
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

