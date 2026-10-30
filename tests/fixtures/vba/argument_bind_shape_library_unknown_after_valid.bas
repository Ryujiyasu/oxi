Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = VBA.Strings.Mid(String:="abc", Start:=1, NoSuchArg:=1)
End Sub
