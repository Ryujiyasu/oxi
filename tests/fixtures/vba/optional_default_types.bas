' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Enum ProbeMode
    pmOff = 0
    pmOn = 3
End Enum
Public Function Driver() As Long
    Driver = 21
End Function
Public Function DByte(Optional ByVal value As Byte = 255) As Long
    DByte = value
End Function
Public Function DInt(Optional ByVal value As Integer = -7) As Long
    DInt = value
End Function
Public Function DLong(Optional ByVal value As Long = -123456) As Long
    DLong = value
End Function
Public Function DLongLong(Optional ByVal value As LongLong = 9007199254740993^) As String
    DLongLong = CStr(value)
End Function
Public Function DSingle(Optional ByVal value As Single = 1.25!) As Double
    DSingle = value
End Function
Public Function DDouble(Optional ByVal value As Double = 1.125#) As Double
    DDouble = value
End Function
Public Function DCurrency(Optional ByVal value As Currency = -12.3456@) As String
    DCurrency = CStr(value)
End Function
Public Function DDate(Optional ByVal value As Date = #2/3/2004 4:05:06 AM#) As String
    DDate = CStr(Year(value)) & "|" & CStr(Month(value)) & "|" & CStr(Day(value)) & "|" & CStr(Hour(value)) & "|" & CStr(Minute(value)) & "|" & CStr(Second(value))
End Function
Public Function DBool(Optional ByVal value As Boolean = True) As Boolean
    DBool = value
End Function
Public Function DAscii(Optional ByVal value As String = "cache") As String
    DAscii = value
End Function
Public Function DJapanese(Optional ByVal value As String = "日本語") As String
    DJapanese = value
End Function
Public Function DMissing(Optional ByVal value As Variant) As Boolean
    DMissing = IsMissing(value)
End Function
Public Function DEmpty(Optional ByVal value As Variant = Empty) As Long
    DEmpty = VarType(value)
End Function
Public Function DNull(Optional ByVal value As Variant = Null) As Long
    DNull = VarType(value)
End Function
Public Function DObject(Optional ByVal value As Object = Nothing) As Boolean
    DObject = (value Is Nothing)
End Function
Public Function DEnum(Optional ByVal value As ProbeMode = pmOn) As Long
    DEnum = value
End Function
Public Function DPair(Optional ByVal number As Long = 5, Optional ByVal text As String = "cache") As String
    DPair = CStr(number) & "|" & text
End Function
Public Function DTriple(Optional ByVal enabled As Boolean = True, Optional ByVal fraction As Double = 1.25, Optional ByVal number As Long = 7) As Double
    If Not enabled Then Err.Raise 5
    DTriple = fraction + number
End Function
Public Function DLongText(Optional ByVal text As String = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", Optional ByVal number As Long = 7) As String
    DLongText = CStr(Len(text)) & ":" & CStr(number)
End Function
Public Function IEmpty() As Long
    IEmpty = DEmpty()
End Function
Public Function INull() As Long
    INull = DNull()
End Function
Public Function IMissing() As Boolean
    IMissing = DMissing()
End Function
Public Function REmpty() As Long
    REmpty = Application.Run("DEmpty")
End Function
Public Function RNull() As Long
    RNull = Application.Run("DNull")
End Function
Public Function RMissing() As Boolean
    RMissing = Application.Run("DMissing")
End Function
