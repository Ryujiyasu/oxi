Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function NextValue(ByVal text As String) As Long
Log = Log & text
NextValue = 1
End Function
Public Property Let Item(ByVal first As Long, ByVal second As Long, ByVal value As Long)
Log = Log & "C"
End Property
Public Function Probe() As String
Log = ""
Item(second:=NextValue("S"),first:=NextValue("F")) = NextValue("V")
Probe = Log
End Function
