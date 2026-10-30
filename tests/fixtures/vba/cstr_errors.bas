VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Attribute VB_Name = "Driver"
Option Explicit
Private Function MissingText(Optional value As Variant) As String
MissingText = CStr(value)
End Function
Public Function Probe() As String
Dim report As String
report = "missing=" & MissingText() & vbLf
report = report & "error0=" & CStr(CVErr(0)) & vbLf
report = report & "error1=" & CStr(CVErr(1)) & vbLf
report = report & "error13=" & CStr(CVErr(13)) & vbLf
report = report & "error448=" & CStr(CVErr(448)) & vbLf
report = report & "error2015=" & CStr(CVErr(2015)) & vbLf
report = report & "error2042=" & CStr(CVErr(2042)) & vbLf
report = report & "error65535=" & CStr(CVErr(65535)) & vbLf
Probe = report
End Function

