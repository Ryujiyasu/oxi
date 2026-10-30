' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function LoadedHeightSnapshot() As String
    LoadedHeightSnapshot = CStr(Rows("4").RowHeight) & "|" & CStr(Rows("4").Height) & "|" & CStr(Rows("4").Hidden) & "|" & CStr(Rows("4").UseStandardHeight)
End Function
