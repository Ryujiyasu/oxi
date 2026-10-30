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

Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Private Function GetMacro(ByVal kind As Long) As Variant
Trace = Trace & "N"
Select Case kind
Case 0: GetMacro = "Targets.Echo"
Case 1: GetMacro = Empty
Case 2: GetMacro = Null
Case 3: GetMacro = CByte(1)
Case 4: GetMacro = CLng(0)
Case 5: GetMacro = CLng(-1)
Case 6: GetMacro = CDbl(1.5)
Case 7: GetMacro = True
Case 8: GetMacro = False
Case 9: GetMacro = CVErr(2042)
Case 10: GetMacro = Array("Targets.Echo")
Case 11: Set GetMacro = Nothing
Case 12: Set GetMacro = Range("A1")
Case 13: Set GetMacro = Range("A2")
Case 14: Set GetMacro = Range("A3")
Case 15: Set GetMacro = New NameBox
Case 16: Set GetMacro = New NoDefault
Case 17: GetMacro = ""
Case 18: GetMacro = " "
Case 19: GetMacro = "1"
Case 20: GetMacro = "Targets.Zero"
End Select
End Function
Private Function GetArgument() As Variant
Trace = Trace & "A"
GetArgument = CLng(7)
End Function
Private Function TryRun(ByVal kind As Long, ByVal route As Long) As String
Dim Value As Variant
Dim failure As Long
Dim text As String
Trace = ""
On Error Resume Next
Err.Clear
If kind = -1 Then
Select Case route
Case 0: Value = Application.Run()
Case 1
With Application
Value = .Run()
End With
Case 2: Value = Run()
Case 3: Value = CallByName(Application, "Run", VbMethod)
End Select
Else
Select Case route
Case 0: Value = Application.Run(GetMacro(kind), GetArgument())
Case 1
With Application
Value = .Run(GetMacro(kind), GetArgument())
End With
Case 2: Value = Run(GetMacro(kind), GetArgument())
Case 3: Value = CallByName(Application, "Run", VbMethod, GetMacro(kind), GetArgument())
End Select
End If
failure = Err.Number
Err.Clear
If IsObject(Value) Then
text = "object"
ElseIf IsArray(Value) Then
text = "array"
ElseIf IsError(Value) Then
text = CStr(Value)
ElseIf IsNull(Value) Then
text = "null"
Else
text = CStr(Value)
End If
TryRun = CStr(failure) & "|" & TypeName(Value) & "|" & text & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
Range("A1").Value = "Targets.Echo"
Range("A2").ClearContents
Range("A3").Value = CVErr(2042)
report = report & "r0k-1=" & TryRun(-1, 0) & vbLf
report = report & "r0k0=" & TryRun(0, 0) & vbLf
report = report & "r0k1=" & TryRun(1, 0) & vbLf
report = report & "r0k2=" & TryRun(2, 0) & vbLf
report = report & "r0k3=" & TryRun(3, 0) & vbLf
report = report & "r0k4=" & TryRun(4, 0) & vbLf
report = report & "r0k5=" & TryRun(5, 0) & vbLf
report = report & "r0k6=" & TryRun(6, 0) & vbLf
report = report & "r0k7=" & TryRun(7, 0) & vbLf
report = report & "r0k8=" & TryRun(8, 0) & vbLf
report = report & "r0k9=" & TryRun(9, 0) & vbLf
report = report & "r0k11=" & TryRun(11, 0) & vbLf
report = report & "r0k12=" & TryRun(12, 0) & vbLf
report = report & "r0k13=" & TryRun(13, 0) & vbLf
report = report & "r0k14=" & TryRun(14, 0) & vbLf
report = report & "r0k15=" & TryRun(15, 0) & vbLf
report = report & "r0k16=" & TryRun(16, 0) & vbLf
report = report & "r0k17=" & TryRun(17, 0) & vbLf
report = report & "r0k18=" & TryRun(18, 0) & vbLf
report = report & "r0k19=" & TryRun(19, 0) & vbLf
report = report & "r0k20=" & TryRun(20, 0) & vbLf
report = report & "r1k-1=" & TryRun(-1, 1) & vbLf
report = report & "r1k0=" & TryRun(0, 1) & vbLf
report = report & "r1k1=" & TryRun(1, 1) & vbLf
report = report & "r1k2=" & TryRun(2, 1) & vbLf
report = report & "r1k3=" & TryRun(3, 1) & vbLf
report = report & "r1k4=" & TryRun(4, 1) & vbLf
report = report & "r1k5=" & TryRun(5, 1) & vbLf
report = report & "r1k6=" & TryRun(6, 1) & vbLf
report = report & "r1k7=" & TryRun(7, 1) & vbLf
report = report & "r1k8=" & TryRun(8, 1) & vbLf
report = report & "r1k9=" & TryRun(9, 1) & vbLf
report = report & "r1k11=" & TryRun(11, 1) & vbLf
report = report & "r1k12=" & TryRun(12, 1) & vbLf
report = report & "r1k13=" & TryRun(13, 1) & vbLf
report = report & "r1k14=" & TryRun(14, 1) & vbLf
report = report & "r1k15=" & TryRun(15, 1) & vbLf
report = report & "r1k16=" & TryRun(16, 1) & vbLf
report = report & "r1k17=" & TryRun(17, 1) & vbLf
report = report & "r1k18=" & TryRun(18, 1) & vbLf
report = report & "r1k19=" & TryRun(19, 1) & vbLf
report = report & "r1k20=" & TryRun(20, 1) & vbLf
report = report & "r2k-1=" & TryRun(-1, 2) & vbLf
report = report & "r2k0=" & TryRun(0, 2) & vbLf
report = report & "r2k1=" & TryRun(1, 2) & vbLf
report = report & "r2k2=" & TryRun(2, 2) & vbLf
report = report & "r2k3=" & TryRun(3, 2) & vbLf
report = report & "r2k4=" & TryRun(4, 2) & vbLf
report = report & "r2k5=" & TryRun(5, 2) & vbLf
report = report & "r2k6=" & TryRun(6, 2) & vbLf
report = report & "r2k7=" & TryRun(7, 2) & vbLf
report = report & "r2k8=" & TryRun(8, 2) & vbLf
report = report & "r2k9=" & TryRun(9, 2) & vbLf
report = report & "r2k11=" & TryRun(11, 2) & vbLf
report = report & "r2k12=" & TryRun(12, 2) & vbLf
report = report & "r2k13=" & TryRun(13, 2) & vbLf
report = report & "r2k14=" & TryRun(14, 2) & vbLf
report = report & "r2k15=" & TryRun(15, 2) & vbLf
report = report & "r2k16=" & TryRun(16, 2) & vbLf
report = report & "r2k17=" & TryRun(17, 2) & vbLf
report = report & "r2k18=" & TryRun(18, 2) & vbLf
report = report & "r2k19=" & TryRun(19, 2) & vbLf
report = report & "r2k20=" & TryRun(20, 2) & vbLf
report = report & "r3k-1=" & TryRun(-1, 3) & vbLf
report = report & "r3k0=" & TryRun(0, 3) & vbLf
report = report & "r3k1=" & TryRun(1, 3) & vbLf
report = report & "r3k2=" & TryRun(2, 3) & vbLf
report = report & "r3k3=" & TryRun(3, 3) & vbLf
report = report & "r3k4=" & TryRun(4, 3) & vbLf
report = report & "r3k5=" & TryRun(5, 3) & vbLf
report = report & "r3k6=" & TryRun(6, 3) & vbLf
report = report & "r3k7=" & TryRun(7, 3) & vbLf
report = report & "r3k8=" & TryRun(8, 3) & vbLf
report = report & "r3k9=" & TryRun(9, 3) & vbLf
report = report & "r3k11=" & TryRun(11, 3) & vbLf
report = report & "r3k12=" & TryRun(12, 3) & vbLf
report = report & "r3k13=" & TryRun(13, 3) & vbLf
report = report & "r3k14=" & TryRun(14, 3) & vbLf
report = report & "r3k15=" & TryRun(15, 3) & vbLf
report = report & "r3k16=" & TryRun(16, 3) & vbLf
report = report & "r3k17=" & TryRun(17, 3) & vbLf
report = report & "r3k18=" & TryRun(18, 3) & vbLf
report = report & "r3k19=" & TryRun(19, 3) & vbLf
report = report & "r3k20=" & TryRun(20, 3) & vbLf
Probe = report
End Function
Public Function ProbeOne(ByVal kind As Long, ByVal route As Long) As String
Range("A1").Value = "Targets.Echo"
Range("A2").ClearContents
Range("A3").Value = CVErr(2042)
ProbeOne = TryRun(kind, route)
End Function


Attribute VB_Name = "Targets"
Option Explicit
Public Function Echo(Optional Value As Variant) As String
Driver.Trace = Driver.Trace & "B"
Echo = TypeName(Value) & ":" & CStr(Value)
End Function
Public Function Zero() As String
Driver.Trace = Driver.Trace & "Z"
Zero = "zero"
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NameBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
Driver.Trace = Driver.Trace & "D"
Value = "Targets.Echo"
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

