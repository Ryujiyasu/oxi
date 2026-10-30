Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(Worksheets.Count) & ":" & CStr(Sheets.Count) & ":" & CStr(Charts.Count)
End Function
