Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 Dim counter As Long
 Dim arr() As Long
 Dim target As New Receiver
 value = 1
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
Public CInt As Variant
Public Function ReadValue(ByVal required As Long) As Long
 ReadValue = required
End Function
Public Property Let Indexed(ByVal index As Long, ByVal assigned As Long)
 Value = assigned
End Property
Public Property Get Indexed(ByVal index As Long) As Long
 Indexed = Value
End Property
