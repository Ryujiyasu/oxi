Attribute VB_Name = "DriverModule"
Option Explicit
Public trace As String
Public Function Probe() As String
 On Error GoTo Failed
 Dim value As Variant
 trace = ""
 value = Factory(IndexArg()).ReadValue(7)
 Probe = trace & "|value:" & CStr(value)
 Exit Function
Failed:
 Probe = trace & "|error:" & CStr(Err.Number)
End Function
Private Function TypedFactory() As IndexedReceiver
 trace = trace & "F"
 Set TypedFactory = New IndexedReceiver
End Function
Private Sub ZeroSub()
End Sub
Private Function IndexArg() As Long
 trace = trace & "A"
 IndexArg = 1
End Function
Private Function FailArg() As Long
 trace = trace & "A"
 Err.Raise 5
End Function
Private Function ArrayFactory() As Variant
 trace = trace & "F"
 ArrayFactory = Array(4, 5)
End Function
Private Sub Unused()
 Dim value As Variant
 Dim target As New Receiver
 value = 1
End Sub
Private Function Factory() As Receiver
 trace = trace & "F"
 Set Factory = New Receiver
End Function
Private Function RequiredFactory(ByVal required As Long) As Receiver
 Set RequiredFactory = New Receiver
End Function
Private Function Known(ByVal required As Long) As Long
 Known = required
End Function
Private Function OptionalFactory(Optional ByVal token As Long = 1) As Receiver
 Set OptionalFactory = New Receiver
End Function
Private Function ZeroVariant() As Variant
 Set ZeroVariant = New Receiver
End Function
Private Function ZeroObject() As Object
 Set ZeroObject = New Receiver
End Function
Private Function ZeroScalar() As Long
 ZeroScalar = 7
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Option Explicit
Public Value As Long
Public Property Get Child() As Receiver
 Set Child = Me
End Property
Public Function ReadValue(ByVal required As Long) As Long
 ReadValue = required
End Function
Public Property Let Indexed(ByVal index As Long, ByVal assigned As Long)
 Value = assigned
End Property
Public Property Get Indexed(ByVal index As Long) As Long
 Indexed = Value
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "IndexedReceiver"
Public Property Get Item(ByVal index As Long) As Receiver
Attribute Item.VB_UserMemId = 0
 DriverModule.trace = DriverModule.trace & "D"
 Set Item = New Receiver
End Property
