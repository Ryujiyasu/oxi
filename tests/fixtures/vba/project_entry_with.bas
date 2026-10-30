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
    Driver.Trace = Driver.Trace & "b"
    Who = "book"
End Function
Private Function PrivateWho() As String
    Driver.Trace = Driver.Trace & "bp"
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
    Driver.Trace = Driver.Trace & "s"
    Who = "sheet"
End Function
Private Function PrivateWho() As String
    Driver.Trace = Driver.Trace & "sp"
    PrivateWho = "sheet-private"
End Function
Public Function DocOnly() As String
Driver.Trace = Driver.Trace & "D"
DocOnly = "doc-only"
End Function
Public Function Needed(ByVal n As Long) As String
Driver.Trace = Driver.Trace & "N" & CStr(n)
Needed = "doc-need:" & CStr(n)
End Function
Public Function OptionalArg(Optional n As Long = 7) As String
Driver.Trace = Driver.Trace & "O" & CStr(n)
OptionalArg = "doc-optional:" & CStr(n)
End Function
Public Sub WriteOnly()
Driver.Trace = Driver.Trace & "Z"
End Sub
Public Property Get Label() As String
Driver.Trace = Driver.Trace & "Q"
Label = "sheet-label"
End Property


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
Private Function SharedChoice() As String
    SharedChoice = "alpha-hidden"
End Function
Public Function LocalWho() As String
Dim value As Variant
Driver.Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run("Who")
LocalWho = CStr(Err.Number) & "~" & CStr(value) & "~" & Driver.Trace
End Function
Public Property Get Label() As String
Driver.Trace = Driver.Trace & "L"
Label = "alpha-label"
End Property


Attribute VB_Name = "Beta"
Option Explicit
Public Function Who() As String
    Who = "beta"
End Function
Public Function SharedChoice() As String
    SharedChoice = "beta-public"
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Attribute VB_Exposed = True
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
Dim value As Variant, macroName As String, withArgument As Boolean
If Left(entryName, 5) = "@raw:" Then
    macroName = Mid(entryName, 6)
ElseIf Left(entryName, 5) = "@arg:" Then
    macroName = "'" & ThisWorkbook.Name & "'!" & Mid(entryName, 6)
    withArgument = True
Else
    macroName = "'" & ThisWorkbook.Name & "'!" & entryName
End If
Trace = ""
On Error Resume Next
Err.Clear
If withArgument Then
    With Application
    value = .Run(macroName, 11)
    End With
Else
    With Application
    value = .Run(macroName)
    End With
End If
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
report = report & "PrivateWho=" & TryRun("PrivateWho") & vbLf
report = report & "SharedChoice=" & TryRun("SharedChoice") & vbLf
report = report & "Alpha.SharedChoice=" & TryRun("Alpha.SharedChoice") & vbLf
report = report & "Beta.SharedChoice=" & TryRun("Beta.SharedChoice") & vbLf
report = report & "Alpha.LocalWho=" & TryRun("Alpha.LocalWho") & vbLf
report = report & "@raw:Alpha.LocalWho=" & TryRun("@raw:Alpha.LocalWho") & vbLf
report = report & "VisibleFn=" & TryRun("VisibleFn") & vbLf
report = report & "Sheet1.DocOnly=" & TryRun("Sheet1.DocOnly") & vbLf
report = report & "DocOnly=" & TryRun("DocOnly") & vbLf
report = report & "Sheet1.Needed=" & TryRun("Sheet1.Needed") & vbLf
report = report & "@arg:Sheet1.Needed=" & TryRun("@arg:Sheet1.Needed") & vbLf
report = report & "@arg:Sheet1.Who=" & TryRun("@arg:Sheet1.Who") & vbLf
report = report & "Sheet1.OptionalArg=" & TryRun("Sheet1.OptionalArg") & vbLf
report = report & "Sheet1.WriteOnly=" & TryRun("Sheet1.WriteOnly") & vbLf
report = report & "Alpha.Label=" & TryRun("Alpha.Label") & vbLf
report = report & "Sheet1.Label=" & TryRun("Sheet1.Label") & vbLf
report = report & "@raw:VBAProject.Alpha.Who=" & TryRun("@raw:VBAProject.Alpha.Who") & vbLf
report = report & "@raw:'missing.xlsm'!Alpha.Who=" & TryRun("@raw:'missing.xlsm'!Alpha.Who") & vbLf
report = report & "@raw:Alpha.Who=" & TryRun("@raw:Alpha.Who") & vbLf
report = report & "@raw:Who=" & TryRun("@raw:Who") & vbLf
Probe = report
End Function

