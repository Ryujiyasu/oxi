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
Case 21: GetMacro = "1.0"
Case 22: GetMacro = " 1 "
Case 23: GetMacro = "-1"
Case 24: GetMacro = "1E2"
Case 25: GetMacro = "=1"
Case 26: GetMacro = "1+1"
Case 27: GetMacro = "TRUE"
Case 28: GetMacro = "False"
Case 29: GetMacro = "A1"
Case 30: GetMacro = "Sheet1!A1"
Case 31: GetMacro = "Targets.Echo()"
Case 32: GetMacro = "Targets.Echo "
Case 33: GetMacro = " Targets.Echo"
Case 34: GetMacro = "   Targets.Echo"
Case 35: GetMacro = "=Targets.Echo"
Case 36: GetMacro = "C1"
Case 37: GetMacro = "#N/A"
Case 38: GetMacro = "1,5"
Case 39: GetMacro = "0x01"
Case 40: GetMacro = "&H1"
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
report = report & "r0k21=" & TryRun(21, 0) & vbLf
report = report & "r0k22=" & TryRun(22, 0) & vbLf
report = report & "r0k23=" & TryRun(23, 0) & vbLf
report = report & "r0k24=" & TryRun(24, 0) & vbLf
report = report & "r0k25=" & TryRun(25, 0) & vbLf
report = report & "r0k26=" & TryRun(26, 0) & vbLf
report = report & "r0k27=" & TryRun(27, 0) & vbLf
report = report & "r0k28=" & TryRun(28, 0) & vbLf
report = report & "r0k29=" & TryRun(29, 0) & vbLf
report = report & "r0k30=" & TryRun(30, 0) & vbLf
report = report & "r0k31=" & TryRun(31, 0) & vbLf
report = report & "r0k32=" & TryRun(32, 0) & vbLf
report = report & "r0k33=" & TryRun(33, 0) & vbLf
report = report & "r0k34=" & TryRun(34, 0) & vbLf
report = report & "r0k35=" & TryRun(35, 0) & vbLf
report = report & "r0k36=" & TryRun(36, 0) & vbLf
report = report & "r0k37=" & TryRun(37, 0) & vbLf
report = report & "r0k38=" & TryRun(38, 0) & vbLf
report = report & "r0k39=" & TryRun(39, 0) & vbLf
report = report & "r0k40=" & TryRun(40, 0) & vbLf
report = report & "r1k21=" & TryRun(21, 1) & vbLf
report = report & "r1k22=" & TryRun(22, 1) & vbLf
report = report & "r1k23=" & TryRun(23, 1) & vbLf
report = report & "r1k24=" & TryRun(24, 1) & vbLf
report = report & "r1k25=" & TryRun(25, 1) & vbLf
report = report & "r1k26=" & TryRun(26, 1) & vbLf
report = report & "r1k27=" & TryRun(27, 1) & vbLf
report = report & "r1k28=" & TryRun(28, 1) & vbLf
report = report & "r1k29=" & TryRun(29, 1) & vbLf
report = report & "r1k30=" & TryRun(30, 1) & vbLf
report = report & "r1k31=" & TryRun(31, 1) & vbLf
report = report & "r1k32=" & TryRun(32, 1) & vbLf
report = report & "r1k33=" & TryRun(33, 1) & vbLf
report = report & "r1k34=" & TryRun(34, 1) & vbLf
report = report & "r1k35=" & TryRun(35, 1) & vbLf
report = report & "r1k36=" & TryRun(36, 1) & vbLf
report = report & "r1k37=" & TryRun(37, 1) & vbLf
report = report & "r1k38=" & TryRun(38, 1) & vbLf
report = report & "r1k39=" & TryRun(39, 1) & vbLf
report = report & "r1k40=" & TryRun(40, 1) & vbLf
report = report & "r2k21=" & TryRun(21, 2) & vbLf
report = report & "r2k22=" & TryRun(22, 2) & vbLf
report = report & "r2k23=" & TryRun(23, 2) & vbLf
report = report & "r2k24=" & TryRun(24, 2) & vbLf
report = report & "r2k25=" & TryRun(25, 2) & vbLf
report = report & "r2k26=" & TryRun(26, 2) & vbLf
report = report & "r2k27=" & TryRun(27, 2) & vbLf
report = report & "r2k28=" & TryRun(28, 2) & vbLf
report = report & "r2k29=" & TryRun(29, 2) & vbLf
report = report & "r2k30=" & TryRun(30, 2) & vbLf
report = report & "r2k31=" & TryRun(31, 2) & vbLf
report = report & "r2k32=" & TryRun(32, 2) & vbLf
report = report & "r2k33=" & TryRun(33, 2) & vbLf
report = report & "r2k34=" & TryRun(34, 2) & vbLf
report = report & "r2k35=" & TryRun(35, 2) & vbLf
report = report & "r2k36=" & TryRun(36, 2) & vbLf
report = report & "r2k37=" & TryRun(37, 2) & vbLf
report = report & "r2k38=" & TryRun(38, 2) & vbLf
report = report & "r2k39=" & TryRun(39, 2) & vbLf
report = report & "r2k40=" & TryRun(40, 2) & vbLf
report = report & "r3k21=" & TryRun(21, 3) & vbLf
report = report & "r3k22=" & TryRun(22, 3) & vbLf
report = report & "r3k23=" & TryRun(23, 3) & vbLf
report = report & "r3k24=" & TryRun(24, 3) & vbLf
report = report & "r3k25=" & TryRun(25, 3) & vbLf
report = report & "r3k26=" & TryRun(26, 3) & vbLf
report = report & "r3k27=" & TryRun(27, 3) & vbLf
report = report & "r3k28=" & TryRun(28, 3) & vbLf
report = report & "r3k29=" & TryRun(29, 3) & vbLf
report = report & "r3k30=" & TryRun(30, 3) & vbLf
report = report & "r3k31=" & TryRun(31, 3) & vbLf
report = report & "r3k32=" & TryRun(32, 3) & vbLf
report = report & "r3k33=" & TryRun(33, 3) & vbLf
report = report & "r3k34=" & TryRun(34, 3) & vbLf
report = report & "r3k35=" & TryRun(35, 3) & vbLf
report = report & "r3k36=" & TryRun(36, 3) & vbLf
report = report & "r3k37=" & TryRun(37, 3) & vbLf
report = report & "r3k38=" & TryRun(38, 3) & vbLf
report = report & "r3k39=" & TryRun(39, 3) & vbLf
report = report & "r3k40=" & TryRun(40, 3) & vbLf
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

