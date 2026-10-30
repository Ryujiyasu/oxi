Public Log As String
Private held As Collection
Private aliasObject As Object
Private values(1 To 2) As Long
Private target As Object
Public Sub Prime()
Set held = New Collection
Set aliasObject = held
held.Add New Counter
Set target = Range("A1")
values(1) = 41
Open "state.txt" For Output As #1
Print #1, "retained"
Close #1
End Sub
Public Function ReadState() As String
Static count As Long
count = count + 1
values(1) = values(1) + 1
ReadState = CStr(count) & "|" & CStr(values(1)) & "|" & CStr(held Is aliasObject) & "|" & CStr(held(1).NextValue()) & "|" & CStr(target.Value) & "|" & Log
End Function
Public Function NextRandom() As Single
NextRandom = Rnd()
End Function
Public Function ReadFile() As String
Dim text As String
Open "state.txt" For Input As #1
Line Input #1, text
Close #1
ReadFile = text
End Function
Public Function Release() As String
Set held = Nothing
Set aliasObject = Nothing
Release = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Counter"
Private value As Long
Private Sub Class_Initialize()
Log = Log & "I"
End Sub
Private Sub Class_Terminate()
Log = Log & "T"
End Sub
Public Function NextValue() As Long
value = value + 1
NextValue = value
End Function
