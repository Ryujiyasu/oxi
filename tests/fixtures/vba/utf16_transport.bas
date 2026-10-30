' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function ReturnHigh() As String
    ReturnHigh = ChrW(&HD83D)
End Function
Public Function Echo(ByVal text As String) As String
    Echo = text
End Function
Public Function Parts() As Variant
    Parts = Array(ChrW(&HD83D), ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00))
End Function
