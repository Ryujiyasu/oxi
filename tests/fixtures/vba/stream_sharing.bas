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
    FileAttributesIdentity = 2026102080
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
    Case 208
        Open "held.txt" For Binary Access Read Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 209
        Open "held.txt" For Binary Access Read Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 210
        Open "held.txt" For Binary Access Read Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 211
        Open "held.txt" For Binary Access Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 212
        Open "held.txt" For Binary Access Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 213
        Open "held.txt" For Binary Access Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 214
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 215
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 216
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 217
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 218
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 219
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 220
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 221
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 222
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 223
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 224
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 225
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Write Shared As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 226
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 227
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 228
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 229
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 230
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 231
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 232
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 233
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 234
        Set held = fs.OpenTextFile("held.txt", 8)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held2 = fs.OpenTextFile("held.txt", 8)
        result = result & ":second:" & CStr(Err.Number)
    Case 235
        Open "held.txt" For Binary Access Read Lock Write As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 236
        Open "held.txt" For Binary Access Read Lock Write As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 237
        Open "held.txt" For Binary Access Write Lock Read As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 238
        Open "held.txt" For Binary Access Write Lock Read As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 239
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Lock Write As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 240
        Set held = fs.OpenTextFile("held.txt", 1)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Write Lock Read As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 241
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Read Lock Write As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 242
        Set held = fs.OpenTextFile("held.txt", 2)
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Open "held.txt" For Binary Access Write Lock Read As #1
        result = result & ":second:" & CStr(Err.Number)
    Case 243
        Open "held.txt" For Input As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 244
        Open "held.txt" For Input As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
    Case 245
        Open "held.txt" For Output As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 1)
        result = result & ":second:" & CStr(Err.Number)
    Case 246
        Open "held.txt" For Output As #1
        result = "first:" & CStr(Err.Number)
        Err.Clear
        Set held = fs.OpenTextFile("held.txt", 2)
        result = result & ":second:" & CStr(Err.Number)
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

