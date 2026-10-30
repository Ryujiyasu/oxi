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
Public Function FolderMutationIdentity() As Long
    FolderMutationIdentity = 2026102061
End Function
Public Function DirectoryEnvironment() As Variant
    DirectoryEnvironment = Array(Left$(CurDir$, 1), CurDir$("C"), CurDir$("D"))
End Function
Public Function FolderMutation(ByVal kind As Long, ByVal root As String) As String
    Dim fso As Object, held As Object, item As Variant, result As String, snapshot As String
    Dim number As Long, origin As String, description As String, snapshotError As Long
    ChDrive Left$(root, 1)
    ChDir root
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 0: MkDir "missing\child"
    Case 1: MkDir "exists"
    Case 2: MkDir "held.txt"
    Case 3: RmDir "nonempty"
    Case 4: RmDir "absent"
    Case 5: RmDir "."
    Case 6: RmDir "empty"
    Case 7: Set held = fso.CreateFolder("missing\child")
    Case 8: Set held = fso.CreateFolder("exists")
    Case 9: Set held = fso.CreateFolder("held.txt")
    Case 10: fso.DeleteFolder "nonempty"
    Case 11: fso.DeleteFolder "absent"
    Case 12: fso.DeleteFolder "."
    Case 13: fso.DeleteFolder "empty"
    Case 14: RmDir "empty\.."
    Case 15: RmDir "nonempty\empty-child"
    Case 16: MkDir ""
    Case 17: Set held = fso.CreateFolder("")
    Case 18: MkDir "exists\..\made"
    Case 19: Set held = fso.CreateFolder("exists\..\made")
    Case 20: RmDir "nonempty\child"
    Case 21: fso.DeleteFolder "nonempty\child"
    Case 22: RmDir "held.txt"
    Case 23: fso.DeleteFolder "held.txt"
    Case 24: MkDir "exists\"
    Case 25: Set held = fso.CreateFolder("exists\")
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    If number = 0 Then
        If Not held Is Nothing Then result = held.Path
    End If
    Err.Clear
    snapshot = "|root:" & CStr(fso.FolderExists(root))
    For Each item In Array("exists", "empty", "nonempty", "nonempty\child", "nonempty\empty-child", "made", "missing", "missing\child", "held.txt")
        snapshot = snapshot & "|dir:" & item & ":" & CStr(fso.FolderExists(root & "\" & item))
    Next
    For Each item In Array("held.txt", "nonempty\leaf.txt", "nonempty\child\nested.txt")
        snapshot = snapshot & "|file:" & item & ":" & CStr(fso.FileExists(root & "\" & item))
    Next
    snapshotError = Err.Number
    FolderMutation = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|cwd:" & CurDir$ & snapshot & "|snapshot-error:" & CStr(snapshotError)
    On Error GoTo 0
End Function

