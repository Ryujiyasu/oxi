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
Public Function DirectoryOutputsIdentity() As Long
    DirectoryOutputsIdentity = 2026102040
End Function
Private Sub WriteText(ByVal target As String, ByVal content As String)
    Open target For Output As #1
    Print #1, content
    Close #1
End Sub
Public Function DirectoryOutputs(ByVal kind As Long, ByVal folder As String) As String
    Dim fs As Object, listing As Object, held As Object, result As String, number As Long, origin As String, description As String
    ChDrive Left$(folder, 1)
    ChDir folder
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Select Case kind
    Case 0
        WriteText "same.txt", "FIRST"
        ChDir "sub"
        WriteText "same.txt", "SECOND"
        result = CStr(FileLen(folder & "\same.txt")) & "|" & CStr(FileLen("same.txt"))
    Case 1
        Open "Mixed.txt" For Output As #1
        ChDir "sub"
        Open "mixed.txt" For Output As #2
        Print #1, "FIRST"
        Print #2, "SECOND"
        Close #1, #2
        result = CStr(FileLen(folder & "\Mixed.txt")) & "|" & CStr(FileLen("mixed.txt"))
    Case 2, 3
        WriteText "held.txt", "ROOT"
        Set listing = fs.GetFolder(".").Files
        ChDir "sub"
        WriteText "held.txt", "CHILD"
        If kind = 2 Then
            For Each held In listing
                result = held.Path & "|" & CStr(held.Size)
            Next held
        Else
            Set held = listing.Item("held.txt")
            result = held.Path & "|" & CStr(held.Size)
        End If
    Case 4
        Set listing = fs.GetFolder(".").Files
        WriteText "held.txt", "ROOT"
        ChDir "sub"
        WriteText "held.txt", "CHILD"
        result = CStr(listing.Count)
        For Each held In listing
            result = result & "|" & held.Path & "|" & CStr(held.Size)
        Next held
    Case 5, 6, 7
        WriteText "same.txt", "ROOT"
        ChDir "sub"
        WriteText "same.txt", "CHILD"
        ChDir folder
        If kind = 5 Then
            Name "same.txt" As "renamed.txt"
            result = CStr(FileLen("renamed.txt"))
        ElseIf kind = 6 Then
            Kill "same.txt"
            result = Dir$("same.txt")
        Else
            FileCopy "same.txt", "copy.txt"
            result = CStr(FileLen("copy.txt"))
        End If
    Case 8
        Set held = fs.CreateTextFile("same.txt")
        held.WriteLine "ROOT"
        held.Close
        ChDir "sub"
        Set held = fs.CreateTextFile("same.txt")
        held.WriteLine "CHILD"
        held.Close
        result = CStr(FileLen(folder & "\same.txt")) & "|" & CStr(FileLen("same.txt"))
    Case 9
        WriteText "sub\..\same.txt", "ROOT"
        WriteText ".\sub\same.txt", "CHILD"
        result = fs.GetFile("same.txt").Path & "|" & fs.GetFile("sub\same.txt").Path
    End Select
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    On Error GoTo 0
    DirectoryOutputs = CStr(number) & "|" & origin & "|" & description & "|result:" & result & "|cwd:" & CurDir$
End Function

