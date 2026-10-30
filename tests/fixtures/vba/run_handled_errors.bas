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
Public Function Handled(ByRef held As Variant) As String
On Error GoTo Caught
held = 99
Driver.Trace = Driver.Trace & "E"
Err.Raise 5, "Inner", "boom", "inner.chm", 27
Exit Function
Caught:
Driver.Trace = Driver.Trace & "H"
Handled = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description & ":" & Err.HelpFile & ":" & CStr(Err.HelpContext)
Err.Clear
End Function
Public Function ResumeBody(ByRef held As Variant) As String
On Error Resume Next
held = 99
Driver.Trace = Driver.Trace & "E"
Err.Raise 5, "Inner", "boom"
ResumeBody = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
Err.Clear
held = 100
Driver.Trace = Driver.Trace & "R"
End Function
Public Function ResumeLabel(ByRef held As Variant) As String
Dim caughtNumber As Long
On Error GoTo Caught
held = 99
Driver.Trace = Driver.Trace & "E"
Err.Raise 5, "Inner", "boom"
Back:
ResumeLabel = CStr(caughtNumber) & ":" & CStr(Err.Number)
Driver.Trace = Driver.Trace & "R"
Exit Function
Caught:
caughtNumber = Err.Number
Resume Back
End Function
Public Function Clean(ByRef held As Variant) As String
held = 99
Driver.Trace = Driver.Trace & "C"
Clean = "body:" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function
Public Function HandledThenClean(ByRef held As Variant) As String
Dim answer As String
answer = Handled(held)
HandledThenClean = answer & ":after:" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function
Public Function NestedRun(ByRef held As Variant) As String
Dim answer As Variant
On Error GoTo Caught
answer = Application.Run("Targets.Handled", held)
NestedRun = CStr(answer) & ":after:" & CStr(Err.Number) & ":" & CStr(held)
Exit Function
Caught:
NestedRun = "caught:" & CStr(Err.Number)
Err.Clear
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Private Function Try0() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Handled", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try1() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ResumeBody", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try2() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ResumeLabel", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try3() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try4() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.HandledThenClean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try5() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.NestedRun", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try6() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.Handled", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try7() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.ResumeBody", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try8() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.ResumeLabel", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try9() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try10() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.HandledThenClean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try11() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.NestedRun", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try12() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Handled", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try13() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ResumeBody", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try14() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ResumeLabel", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try15() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try16() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.HandledThenClean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try17() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.NestedRun", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try18() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.Handled", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try19() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.ResumeBody", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try20() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.ResumeLabel", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try21() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try22() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.HandledThenClean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try23() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.NestedRun", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "run_clear_Handled=" & Try0() & vbLf
report = report & "run_clear_ResumeBody=" & Try1() & vbLf
report = report & "run_clear_ResumeLabel=" & Try2() & vbLf
report = report & "run_clear_Clean=" & Try3() & vbLf
report = report & "run_clear_HandledThenClean=" & Try4() & vbLf
report = report & "run_clear_NestedRun=" & Try5() & vbLf
report = report & "run_prior_Handled=" & Try6() & vbLf
report = report & "run_prior_ResumeBody=" & Try7() & vbLf
report = report & "run_prior_ResumeLabel=" & Try8() & vbLf
report = report & "run_prior_Clean=" & Try9() & vbLf
report = report & "run_prior_HandledThenClean=" & Try10() & vbLf
report = report & "run_prior_NestedRun=" & Try11() & vbLf
report = report & "callbyname_clear_Handled=" & Try12() & vbLf
report = report & "callbyname_clear_ResumeBody=" & Try13() & vbLf
report = report & "callbyname_clear_ResumeLabel=" & Try14() & vbLf
report = report & "callbyname_clear_Clean=" & Try15() & vbLf
report = report & "callbyname_clear_HandledThenClean=" & Try16() & vbLf
report = report & "callbyname_clear_NestedRun=" & Try17() & vbLf
report = report & "callbyname_prior_Handled=" & Try18() & vbLf
report = report & "callbyname_prior_ResumeBody=" & Try19() & vbLf
report = report & "callbyname_prior_ResumeLabel=" & Try20() & vbLf
report = report & "callbyname_prior_Clean=" & Try21() & vbLf
report = report & "callbyname_prior_HandledThenClean=" & Try22() & vbLf
report = report & "callbyname_prior_NestedRun=" & Try23() & vbLf
Probe = report
End Function

