Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = SharedHelper(MissingVariable)
End Sub
Attribute VB_Name = "HelperA"
Public Function SharedHelper(ByVal required As Long) As Long
 SharedHelper = required
End Function
Attribute VB_Name = "HelperB"
Public Function SharedHelper(ByVal required As Long) As Long
 SharedHelper = required
End Function
