VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Function Who() As String
    Who = "book"
End Function
Private Function PrivateWho() As String
    PrivateWho = "book-private"
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Function Who() As String
    Who = "sheet"
End Function
Private Function PrivateWho() As String
    PrivateWho = "sheet-private"
End Function


Attribute VB_Name = "Alpha"
Option Explicit
Public Function Who() As String
    Who = "alpha"
End Function
Private Function PrivateWho() As String
    PrivateWho = "alpha-private"
End Function
Public Function Needed(ByVal n As Long) As String
Needed = "need:" & CStr(n)
End Function
Public Function OptionalArg(Optional n As Long = 7) As String
OptionalArg = "optional:" & CStr(n)
End Function
Public Function Many(ParamArray values() As Variant) As String
Many = "many:" & CStr(UBound(values) - LBound(values) + 1)
End Function


Attribute VB_Name = "Beta"
Option Explicit
Public Function Who() As String
    Who = "beta"
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Class_Initialize()
Driver.Trace = Driver.Trace & "P"
End Sub
Public Function Who() As String
    Who = "plain"
End Function
Private Function PrivateWho() As String
    PrivateWho = "plain-private"
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Singleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Class_Initialize()
Driver.Trace = Driver.Trace & "S"
End Sub
Public Function Who() As String
    Who = "singleton"
End Function
Private Function PrivateWho() As String
    PrivateWho = "singleton-private"
End Function


Attribute VB_Name = "Hidden"
Option Explicit
Option Private Module
Public Function VisibleFn() As String
    VisibleFn = "hidden-module"
End Function


Attribute VB_Name = "標準モジュール"
Option Explicit
Public Function 報告() As String
    報告 = "unicode"
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Sub ResetTrace()
Trace = ""
End Sub
Public Function ReadTrace() As String
ReadTrace = Trace
End Function
Private Function TryRun(ByVal entryName As String) As String
Dim value As Variant
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run("'" & ThisWorkbook.Name & "'!" & entryName)
TryRun = CStr(Err.Number) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "Who=" & TryRun("Who") & vbLf
report = report & "Alpha.Who=" & TryRun("Alpha.Who") & vbLf
report = report & "Beta.Who=" & TryRun("Beta.Who") & vbLf
report = report & "Alpha.PrivateWho=" & TryRun("Alpha.PrivateWho") & vbLf
report = report & "Alpha.Needed=" & TryRun("Alpha.Needed") & vbLf
report = report & "Alpha.OptionalArg=" & TryRun("Alpha.OptionalArg") & vbLf
report = report & "Alpha.Many=" & TryRun("Alpha.Many") & vbLf
report = report & "Hidden.VisibleFn=" & TryRun("Hidden.VisibleFn") & vbLf
report = report & "Plain.Who=" & TryRun("Plain.Who") & vbLf
report = report & "Plain.PrivateWho=" & TryRun("Plain.PrivateWho") & vbLf
report = report & "Singleton.Who=" & TryRun("Singleton.Who") & vbLf
report = report & "Sheet1.Who=" & TryRun("Sheet1.Who") & vbLf
report = report & "Sheet1.PrivateWho=" & TryRun("Sheet1.PrivateWho") & vbLf
report = report & "ThisWorkbook.Who=" & TryRun("ThisWorkbook.Who") & vbLf
report = report & "ThisWorkbook.PrivateWho=" & TryRun("ThisWorkbook.PrivateWho") & vbLf
report = report & "標準モジュール.報告=" & TryRun("標準モジュール.報告") & vbLf
report = report & "報告=" & TryRun("報告") & vbLf
report = report & "Alpha.Absent=" & TryRun("Alpha.Absent") & vbLf
Probe = report
End Function

