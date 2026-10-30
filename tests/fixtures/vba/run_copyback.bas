VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Attribute VB_Name = "Targets"
Option Explicit
Public Sub Scalar(ByRef value As Variant)
value = CLng(99)
End Sub
Public Sub Whole(ByRef value As Long)
value = CLng(99)
End Sub
Public Sub Element(ByRef value As Variant)
value(0) = CLng(99)
End Sub
Public Sub ArrayValue(ByRef value As Variant)
value = Array(9, 8)
End Sub
Public Sub ObjectValue(ByRef value As Object)
value.Tag = "changed"
Set value = Nothing
End Sub
Public Sub ObjectReplace(ByRef value As Object)
Set value = New Item
value.Tag = "new"
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Tag As String


Attribute VB_Name = "Driver"
Option Explicit
Private Function Tryr0k0() As String
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
Application.Run "Targets.Scalar", held
failure = Err.Number
Err.Clear
Tryr0k0 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr0k1() As String
Dim failure As Long
Dim held As Long
held = 1
On Error Resume Next
Err.Clear
Application.Run "Targets.Whole", held
failure = Err.Number
Err.Clear
Tryr0k1 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr0k2() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
Application.Run "Targets.Element", held
failure = Err.Number
Err.Clear
Tryr0k2 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr0k3() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
Application.Run "Targets.ArrayValue", held
failure = Err.Number
Err.Clear
Tryr0k3 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr0k4() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
Application.Run "Targets.ObjectValue", held
failure = Err.Number
Err.Clear
Tryr0k4 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr0k5() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
Application.Run "Targets.ObjectReplace", held
failure = Err.Number
Err.Clear
Tryr0k5 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr1k0() As String
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
With Application
.Run "Targets.Scalar", held
End With
failure = Err.Number
Err.Clear
Tryr1k0 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr1k1() As String
Dim failure As Long
Dim held As Long
held = 1
On Error Resume Next
Err.Clear
With Application
.Run "Targets.Whole", held
End With
failure = Err.Number
Err.Clear
Tryr1k1 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr1k2() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
With Application
.Run "Targets.Element", held
End With
failure = Err.Number
Err.Clear
Tryr1k2 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr1k3() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
With Application
.Run "Targets.ArrayValue", held
End With
failure = Err.Number
Err.Clear
Tryr1k3 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr1k4() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
With Application
.Run "Targets.ObjectValue", held
End With
failure = Err.Number
Err.Clear
Tryr1k4 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr1k5() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
With Application
.Run "Targets.ObjectReplace", held
End With
failure = Err.Number
Err.Clear
Tryr1k5 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr2k0() As String
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
Run "Targets.Scalar", held
failure = Err.Number
Err.Clear
Tryr2k0 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr2k1() As String
Dim failure As Long
Dim held As Long
held = 1
On Error Resume Next
Err.Clear
Run "Targets.Whole", held
failure = Err.Number
Err.Clear
Tryr2k1 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr2k2() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
Run "Targets.Element", held
failure = Err.Number
Err.Clear
Tryr2k2 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr2k3() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
Run "Targets.ArrayValue", held
failure = Err.Number
Err.Clear
Tryr2k3 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr2k4() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
Run "Targets.ObjectValue", held
failure = Err.Number
Err.Clear
Tryr2k4 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr2k5() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
Run "Targets.ObjectReplace", held
failure = Err.Number
Err.Clear
Tryr2k5 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr3k0() As String
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.Scalar", held
failure = Err.Number
Err.Clear
Tryr3k0 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr3k1() As String
Dim failure As Long
Dim held As Long
held = 1
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.Whole", held
failure = Err.Number
Err.Clear
Tryr3k1 = CStr(failure) & "|" & CStr(held)
End Function
Private Function Tryr3k2() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.Element", held
failure = Err.Number
Err.Clear
Tryr3k2 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr3k3() As String
Dim failure As Long
Dim held As Variant
held = Array(1, 2)
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.ArrayValue", held
failure = Err.Number
Err.Clear
Tryr3k3 = CStr(failure) & "|" & CStr(held(0)) & ":" & CStr(held(1))
End Function
Private Function Tryr3k4() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.ObjectValue", held
failure = Err.Number
Err.Clear
Tryr3k4 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function Tryr3k5() As String
Dim failure As Long
Dim held As Object
Set held = New Item
held.Tag = "old"
On Error Resume Next
Err.Clear
CallByName Application, "Run", VbMethod, "Targets.ObjectReplace", held
failure = Err.Number
Err.Clear
Tryr3k5 = CStr(failure) & "|" & ObjectText(held)
End Function
Private Function ObjectText(ByVal value As Object) As String
If value Is Nothing Then
ObjectText = "Nothing"
Else
ObjectText = value.Tag
End If
End Function
Public Function Probe() As String
Dim report As String
report = report & "r0k0=" & Tryr0k0() & vbLf
report = report & "r0k1=" & Tryr0k1() & vbLf
report = report & "r0k2=" & Tryr0k2() & vbLf
report = report & "r0k3=" & Tryr0k3() & vbLf
report = report & "r0k4=" & Tryr0k4() & vbLf
report = report & "r0k5=" & Tryr0k5() & vbLf
report = report & "r1k0=" & Tryr1k0() & vbLf
report = report & "r1k1=" & Tryr1k1() & vbLf
report = report & "r1k2=" & Tryr1k2() & vbLf
report = report & "r1k3=" & Tryr1k3() & vbLf
report = report & "r1k4=" & Tryr1k4() & vbLf
report = report & "r1k5=" & Tryr1k5() & vbLf
report = report & "r2k0=" & Tryr2k0() & vbLf
report = report & "r2k1=" & Tryr2k1() & vbLf
report = report & "r2k2=" & Tryr2k2() & vbLf
report = report & "r2k3=" & Tryr2k3() & vbLf
report = report & "r2k4=" & Tryr2k4() & vbLf
report = report & "r2k5=" & Tryr2k5() & vbLf
report = report & "r3k0=" & Tryr3k0() & vbLf
report = report & "r3k1=" & Tryr3k1() & vbLf
report = report & "r3k2=" & Tryr3k2() & vbLf
report = report & "r3k3=" & Tryr3k3() & vbLf
report = report & "r3k4=" & Tryr3k4() & vbLf
report = report & "r3k5=" & Tryr3k5() & vbLf
Probe = report
End Function

