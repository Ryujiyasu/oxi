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
Public Function Whole(ByRef value As Long) As String
value = 99
Whole = "body"
End Function
Public Function Scalar(ByRef value As Variant) As String
value = 99
Scalar = "body"
End Function
Public Function AliasValue(ByRef first As Variant, ByRef second As Variant) As String
first = 10
second = first + 1
AliasValue = CStr(first) & ":" & CStr(second)
End Function
Public Function Fail(ByRef value As Variant) As String
value = 99
Err.Raise 5
End Function


Attribute VB_Name = "Driver"
Option Explicit
Private Function Try0() As String
Dim result As Variant
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
result = CallByName(Application, "Run", VbMethod, "Targets.Whole", held)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(result) & "|" & CStr(held)
End Function
Private Function Try1() As String
Dim result As Variant
Dim failure As Long
Dim held As Integer
held = 1
On Error Resume Next
Err.Clear
result = CallByName(Application, "Run", VbMethod, "Targets.Whole", held)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(result) & "|" & CStr(held)
End Function
Private Function Try2() As String
Dim result As Variant
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
result = CallByName(Application, "Run", VbMethod, "Targets.Scalar", (held))
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(result) & "|" & CStr(held)
End Function
Private Function Try3() As String
Dim result As Variant
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
result = CallByName(Application, "Run", VbMethod, "Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(result) & "|" & CStr(held)
End Function
Private Function Try4() As String
Dim result As Variant
Dim failure As Long
Dim held As Variant
held = CLng(1)
On Error Resume Next
Err.Clear
result = Application.Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(result) & "|" & CStr(held)
End Function
Public Function Probe() As String
Dim report As String
report = report & "variant_long=" & Try0() & vbLf
report = report & "integer_long=" & Try1() & vbLf
report = report & "forced=" & Try2() & vbLf
report = report & "alias=" & Try3() & vbLf
report = report & "direct_alias=" & Try4() & vbLf
Probe = report
End Function

