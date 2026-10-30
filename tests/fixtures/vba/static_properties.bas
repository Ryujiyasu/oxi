Option Explicit
Public Log As String
Private first As PropertyBox
Private second As PropertyBox
Public Sub Prime()
Set first = New PropertyBox
Set second = New PropertyBox
End Sub
Public Function ReadMixed() As String
Dim held As Object
Dim one As String, two As String, three As String
Set held = New Collection
Log = ""
first.Token = 7
first.Token = 8
Set first.Token = held
one = CStr(first.Token)
two = CStr(first.Token)
second.Token = 9
Set second.Token = held
three = CStr(second.Token)
ReadMixed = one & "|" & two & "|" & three & "|" & Log
End Function
Public Sub ResetFirst()
Set first = Nothing
Set first = New PropertyBox
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "PropertyBox"
Public Property Get Token() As Variant
Static stored As Long
stored = stored + 1
Token = CStr(stored)
End Property
Public Property Let Token(ByVal newValue As Variant)
Static stored As Long
stored = stored + 1
Log = Log & "L" & CStr(stored) & ":" & CStr(newValue) & ";"
End Property
Public Property Set Token(ByVal newValue As Variant)
Static stored As Long
stored = stored + 1
Log = Log & "S" & CStr(stored) & ";"
End Property
