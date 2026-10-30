Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = LeftModule.Inspect() & "|" & RightModule.Inspect()
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
DefLng A-Z
#Const Feature = 1
Private Enum Marker
First = 11
End Enum
Public Function Inspect() As String
Dim amount
#If Feature Then
Inspect = TypeName(amount) & ":" & Marker.First & ":1"
#Else
Inspect = TypeName(amount) & ":" & Marker.First & ":2"
#End If
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Private Enum Marker
First = 22
End Enum
Public Function Inspect() As String
Dim amount
#If Feature Then
Inspect = TypeName(amount) & ":" & Marker.First & ":1"
#Else
Inspect = TypeName(amount) & ":" & Marker.First & ":2"
#End If
End Function
