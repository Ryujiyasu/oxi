Option Explicit
Private count As Long
Public Function Tick() As Long
 count = count + 1
 Tick = count
End Function
Public Sub Bump(ByRef value As Long)
 value = value + 1
End Sub
