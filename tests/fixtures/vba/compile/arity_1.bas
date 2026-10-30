' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit
Private trace As String
Public Function Probe() As String
 On Error Resume Next
 Dim value As Variant
 trace = "entry"
 ' PROBE_ENTRY_MARK
 value = CallByName(Err)
 Probe = Err.Number & "|" & TypeName(value) & "|" & CStr(value)
End Function
Public Function ReadTrace() As String
 ReadTrace = trace
End Function

