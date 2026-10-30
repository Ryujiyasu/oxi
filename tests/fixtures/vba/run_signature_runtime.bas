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
Public Function Echo(ParamArray args() As Variant) As String
Driver.Trace = Driver.Trace & "B"
Echo = CStr(UBound(args) + 1)
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Private Function Try0() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run("Targets.Echo", 7)
failure = Err.Number
Try0 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try1() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run(Macro:="Targets.Echo", Arg1:=7)
failure = Err.Number
Try1 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try2() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run("Targets.Echo", Arg1:=7)
failure = Err.Number
Try2 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try3() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run(Arg1:=7, Macro:="Targets.Echo")
failure = Err.Number
Try3 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try4() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Run(Macro:="Targets.Echo", Arg1:=7)
failure = Err.Number
Try4 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try5() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
Dim app As Excel.Application
Set app = Application
value = app.Run(Macro:="Targets.Echo", Arg1:=7)
failure = Err.Number
Try5 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try6() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
Dim app As Object
Set app = Application
value = app.Run(Macro:="Targets.Echo", Arg1:=7)
failure = Err.Number
Try6 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try7() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run("Targets.Echo", 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
failure = Err.Number
Try7 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try9() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
Dim app As Object
Set app = Application
value = app.Run("Targets.Echo", 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
failure = Err.Number
Try9 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "positional=" & Try0() & vbLf
report = report & "named_macro=" & Try1() & vbLf
report = report & "named_arg=" & Try2() & vbLf
report = report & "named_reverse=" & Try3() & vbLf
report = report & "named_global=" & Try4() & vbLf
report = report & "named_early=" & Try5() & vbLf
report = report & "named_late=" & Try6() & vbLf
report = report & "positional30=" & Try7() & vbLf
report = report & "late31=" & Try9() & vbLf
Probe = report
End Function

