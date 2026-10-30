' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function Driver() As Long
    Driver = 21
End Function
Public Function VEmpty(Optional ByVal value As Variant = Empty) As String
    VEmpty = CStr(VarType(value)) & "|" & CStr(IsMissing(value))
End Function
Public Function VLong(Optional ByVal value As Long = 7) As Long
    VLong = value
End Function
Public Function VText(Optional ByVal value As String = "cache") As String
    VText = value
End Function
Public Function InternalEmpty() As String
    InternalEmpty = VEmpty()
End Function
Public Function RunEmpty() As String
    RunEmpty = Application.Run("VEmpty")
End Function
Public Function ExplicitEmpty() As String
    ExplicitEmpty = Application.Run("VEmpty", Empty)
End Function
Public Function ExplicitNull() As String
    ExplicitNull = Application.Run("VEmpty", Null)
End Function
Public Function ForwardLong(Optional ByVal value As Variant) As Long
    ForwardLong = Application.Run("VLong", value)
End Function
Public Function ForwardText(Optional ByVal value As Variant) As String
    ForwardText = Application.Run("VText", value)
End Function
Public Function ExplicitEmptyLong() As Long
    ExplicitEmptyLong = Application.Run("VLong", Empty)
End Function
