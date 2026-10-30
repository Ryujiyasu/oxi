Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.Token = 7
RightModule.Token = 9
Probe = LeftModule.Token & "|" & RightModule.Token
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Private stored As Long
Public Property Let Token(ByVal value As Long)
Static count As Long
count = count + 1
stored = value + count
End Property
Public Property Get Token() As Long
Static count As Long
count = count + 1
Token = stored + count
End Property

Attribute VB_Name = "RightModule"
Option Explicit
Private stored As Long
Public Property Let Token(ByVal value As Long)
Static count As Long
count = count + 1
stored = value + count
End Property
Public Property Get Token() As Long
Static count As Long
count = count + 1
Token = stored + count
End Property
