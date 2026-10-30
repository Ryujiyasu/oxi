Option Explicit
Private count As Long
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Property Get Current() As Long
 count = count + 1
 Current = count
End Property
