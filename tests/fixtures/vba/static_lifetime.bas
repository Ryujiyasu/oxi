Option Explicit
Public Log As String
Private first As Holder
Private second As Holder
Public Sub Prime()
Log = ""
Set first = New Holder
Set second = New Holder
first.Touch
second.Touch
End Sub
Public Function ReadLog() As String
ReadLog = Log
End Function
Public Function DropFirst() As String
Set first = Nothing
DropFirst = Log
End Function
Public Function DropSecond() As String
Set second = Nothing
DropSecond = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Holder"
Public Sub Touch()
Static saved As Child
If saved Is Nothing Then Set saved = New Child
End Sub
Private Sub Class_Terminate()
Log = Log & "P"
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Child"
Private Sub Class_Initialize()
Log = Log & "I"
End Sub
Private Sub Class_Terminate()
Log = Log & "C"
End Sub
