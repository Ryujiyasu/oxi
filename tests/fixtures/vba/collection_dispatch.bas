' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit
Public Function DispatchMatrix(ByVal member As String, ByVal kind As Long, ByVal argc As Long) As String
 Dim target As Object, value As Variant
 Set target = New Collection
 target.Add 11, "a"
 target.Add 22, "b"
 Err.Clear
 On Error GoTo caught
 Select Case argc
Case 0
value = CallByName(target, member, kind)
Case 1
value = CallByName(target, member, kind, 1)
Case 2
value = CallByName(target, member, kind, 1, "a")
Case 3
value = CallByName(target, member, kind, 1, "a", 1)
Case 4
value = CallByName(target, member, kind, 1, "a", 1, 2)
Case 5
value = CallByName(target, member, kind, 1, "a", 1, 2, 3)
End Select
 DispatchMatrix = "ok|" & TypeName(value) & "|" & CStr(value) & "|" & Snapshot()
 Exit Function
caught:
 DispatchMatrix = "error|" & Snapshot()
End Function
Private Function Snapshot() As String
 Snapshot = Err.Number & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & Err.HelpContext
End Function
