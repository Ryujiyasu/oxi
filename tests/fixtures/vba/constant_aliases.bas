' SPDX-License-Identifier: CC-BY-SA-4.0
' Intentionally no Option Explicit: distinguish imported constants from implicit Empty variables.
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102122
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function AliasCase(ByVal which As Long) As String
    Dim value As Variant
    Select Case which
        Case 0: value = msoShapeFivePointStar
        Case 1: value = msoShape5pointStar
        Case 2: value = xlDataLabelsPositionCenter
        Case 3: value = xlLabelPositionCenter
        Case 4: value = xlDataLabelsPositionOutsideEnd
        Case 5: value = xlLabelPositionOutsideEnd
        Case 6: value = vbRed
        Case 7: value = vbWhite
        Case 8: value = vbYes
        Case 9: value = UndefinedOxiConstant
    End Select
    AliasCase = CStr(VarType(value)) & ":" & TypeName(value) & ":" & CStr(value)
End Function
