Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Function TypedFactory() As IndexedReceiver
 Set TypedFactory = New IndexedReceiver
End Function
Private Sub ZeroSub()
End Sub
Private Sub Unused()
 Dim value As Variant
 Dim target As New Receiver
 value = target.Child(1).NoSuch(1)
End Sub
Private Function Factory() As Receiver
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
 Set Item = New Receiver
End Property
