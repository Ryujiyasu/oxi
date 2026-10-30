' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function Driver() As Long
    Driver = 21
End Function
Public Function LiteralCases() As String
    Const literal As String = "cache"
    Dim value As String
    Dim pointer As LongLong
    Dim report As String
    report = CStr(StrPtr("cache") = StrPtr("cache"))
    report = report & "|" & CStr(StrPtr(literal) = StrPtr(literal))
    report = report & "|" & CStr(StrPtr(literal) = StrPtr("cache"))
    value = "cache"
    pointer = StrPtr(value)
    value = value
    report = report & "|" & CStr(StrPtr(value) = pointer)
    pointer = StrPtr(value)
    value = "cache"
    report = report & "|" & CStr(StrPtr(value) = pointer)
    LiteralCases = report
End Function
