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
Private hits As Long
Public Function NamespaceIdentity() As Long
    NamespaceIdentity = 2026102051
End Function
Public Function DirectoryEnvironment() As Variant
    DirectoryEnvironment = Array(Left$(CurDir$, 1), CurDir$("C"), CurDir$("D"))
End Function
Public Function NamespaceRead(ByVal root As String) As String
    Dim fso As Object, name As Variant, stream As Object, answer As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    hits = hits + 1
    answer = "hits:" & CStr(hits)
    For Each name In Array("seed-empty", "gone-empty", "made-empty", "parent", "parent\child", "after-end")
        answer = answer & "|folder:" & name & ":" & CStr(fso.FolderExists(root & "\" & name))
    Next
    For Each name In Array("seed.txt", "renamed.txt", "parent\child\child.txt", "after-end\end.txt")
        answer = answer & "|file:" & name & ":" & CStr(fso.FileExists(root & "\" & name))
        If fso.FileExists(root & "\" & name) Then
            Set stream = fso.OpenTextFile(root & "\" & name, 1)
            answer = answer & ":" & stream.ReadAll
            stream.Close
        End If
    Next
    NamespaceRead = answer
End Function
Public Function NamespaceMake(ByVal root As String) As String
    MkDir root & "\made-empty"
    NamespaceMake = NamespaceRead(root)
End Function
Public Function NamespaceKill(ByVal root As String) As String
    Kill root & "\seed.txt"
    NamespaceKill = NamespaceRead(root)
End Function
Public Function NamespaceRemove(ByVal root As String) As String
    RmDir root & "\gone-empty"
    NamespaceRemove = NamespaceRead(root)
End Function
Public Function NamespaceTree(ByVal root As String) As String
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    MkDir root & "\parent"
    MkDir root & "\parent\child"
    Set stream = fso.CreateTextFile(root & "\parent\child\child.txt", True)
    stream.Write "child"
    stream.Close
    NamespaceTree = NamespaceRead(root)
End Function
Public Function NamespaceDeleteTree(ByVal root As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder root & "\parent"
    NamespaceDeleteTree = NamespaceRead(root)
End Function
Public Function NamespaceReplace(ByVal root As String) As String
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(root & "\seed.txt", True)
    stream.Write "replacement"
    stream.Close
    NamespaceReplace = NamespaceRead(root)
End Function
Public Function NamespaceRename(ByVal root As String) As String
    Name root & "\seed.txt" As root & "\renamed.txt"
    NamespaceRename = NamespaceRead(root)
End Function
Public Sub NamespaceStop(ByVal root As String)
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    MkDir root & "\after-end"
    Set stream = fso.CreateTextFile(root & "\after-end\end.txt", True)
    stream.Write "before End"
    stream.Close
    RmDir root & "\made-empty"
    End
End Sub

