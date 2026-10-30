Option Explicit
Private count As Long
Public Function Tick() As Long
 count = count + 1
 Tick = count
End Function
