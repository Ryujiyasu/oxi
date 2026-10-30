' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private bodyHits As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102302
End Function

Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function

Private Function AcceptObject(ByVal target As Object) As String
    bodyHits = bodyHits + 1
    AcceptObject = TypeName(target)
End Function

Private Function AcceptRange(ByVal target As Range) As String
    bodyHits = bodyHits + 1
    AcceptRange = TypeName(target)
End Function

Private Function AcceptWorksheet(ByVal target As Worksheet) As String
    bodyHits = bodyHits + 1
    AcceptWorksheet = TypeName(target)
End Function

Private Function AcceptWorkbook(ByVal target As Workbook) As String
    bodyHits = bodyHits + 1
    AcceptWorkbook = TypeName(target)
End Function

Private Function AcceptCollection(ByVal target As Collection) As String
    bodyHits = bodyHits + 1
    AcceptCollection = TypeName(target)
End Function

Private Function AcceptApplication(ByVal target As Application) As String
    bodyHits = bodyHits + 1
    AcceptApplication = TypeName(target)
End Function

Private Function AcceptQualifiedRange(ByVal target As Excel.Range) As String
    bodyHits = bodyHits + 1
    AcceptQualifiedRange = TypeName(target)
End Function

Private Function AcceptQualifiedWorksheet(ByVal target As Excel.Worksheet) As String
    bodyHits = bodyHits + 1
    AcceptQualifiedWorksheet = TypeName(target)
End Function

Private Function AcceptQualifiedWorkbook(ByVal target As Excel.Workbook) As String
    bodyHits = bodyHits + 1
    AcceptQualifiedWorkbook = TypeName(target)
End Function

Private Function AcceptQualifiedCollection(ByVal target As VBA.Collection) As String
    bodyHits = bodyHits + 1
    AcceptQualifiedCollection = TypeName(target)
End Function

Private Function AcceptQualifiedApplication(ByVal target As Excel.Application) As String
    bodyHits = bodyHits + 1
    AcceptQualifiedApplication = TypeName(target)
End Function

Public Function BoundaryCase(ByVal scenario As Long) As String
    Dim value As Variant, answer As String, number As Long, targetKind As Long
    If scenario >= 220 Then
        bodyHits = 0
        BoundaryCase = QualifiedObjectFormsCase(scenario - 220)
        Exit Function
    End If
    bodyHits = 0
    targetKind = scenario \ 20
    scenario = scenario Mod 20
    Select Case scenario
        Case 0: value = Empty
        Case 1: value = Null
        Case 2: value = CVErr(2042)
        Case 3: value = 0
        Case 4: value = 1
        Case 5: value = -1
        Case 6: value = 1.5
        Case 7: value = False
        Case 8: value = True
        Case 9: value = ""
        Case 10: value = "x"
        Case 11: value = CDate(45292)
        Case 12: value = CCur(1.25)
        Case 13: value = Array(1, 2)
        Case 14: Set value = Nothing
        Case 15: Set value = Range("A1")
        Case 16: Set value = ThisWorkbook
        Case 17: Set value = New Collection
        Case 18: Set value = ActiveSheet
        Case 19: Set value = Application
    End Select
    On Error Resume Next
    Err.Clear
    Select Case targetKind
        Case 0: answer = AcceptObject(value)
        Case 1: answer = AcceptRange(value)
        Case 2: answer = AcceptWorksheet(value)
        Case 3: answer = AcceptWorkbook(value)
        Case 4: answer = AcceptCollection(value)
        Case 5: answer = AcceptApplication(value)
        Case 6: answer = AcceptQualifiedRange(value)
        Case 7: answer = AcceptQualifiedWorksheet(value)
        Case 8: answer = AcceptQualifiedWorkbook(value)
        Case 9: answer = AcceptQualifiedCollection(value)
        Case 10: answer = AcceptQualifiedApplication(value)
    End Select
    number = Err.Number
    BoundaryCase = CStr(number) & "|" & CStr(bodyHits) & "|" & answer
End Function

Private Function LocalBareRange(ByVal mode As Long) As String
    Dim target As Range, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = Range("A1")
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalBareRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalQualifiedRange(ByVal mode As Long) As String
    Dim target As Excel.Range, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = Range("A1")
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalQualifiedRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalBareWorksheet(ByVal mode As Long) As String
    Dim target As Worksheet, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = ActiveSheet
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalBareWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalQualifiedWorksheet(ByVal mode As Long) As String
    Dim target As Excel.Worksheet, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = ActiveSheet
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalQualifiedWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalBareWorkbook(ByVal mode As Long) As String
    Dim target As Workbook, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = ThisWorkbook
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalBareWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalQualifiedWorkbook(ByVal mode As Long) As String
    Dim target As Excel.Workbook, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = ThisWorkbook
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalQualifiedWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalBareCollection(ByVal mode As Long) As String
    Dim target As Collection, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = New Collection
        Case 2: Set incoming = Range("A1"): Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalBareCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalQualifiedCollection(ByVal mode As Long) As String
    Dim target As VBA.Collection, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = New Collection
        Case 2: Set incoming = Range("A1"): Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalQualifiedCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalBareApplication(ByVal mode As Long) As String
    Dim target As Application, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = Application
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalBareApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function LocalQualifiedApplication(ByVal mode As Long) As String
    Dim target As Excel.Application, incoming As Variant, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Select Case mode
        Case 1: Set target = Application
        Case 2: Set incoming = New Collection: Set target = incoming
        Case 3: incoming = 1: Set target = incoming
    End Select
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    LocalQualifiedApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedBareRange() As Range
End Function

Private Function ReturnBareRange() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedBareRange()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnBareRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedQualifiedRange() As Excel.Range
End Function

Private Function ReturnQualifiedRange() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedQualifiedRange()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnQualifiedRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedBareWorksheet() As Worksheet
End Function

Private Function ReturnBareWorksheet() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedBareWorksheet()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnBareWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedQualifiedWorksheet() As Excel.Worksheet
End Function

Private Function ReturnQualifiedWorksheet() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedQualifiedWorksheet()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnQualifiedWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedBareWorkbook() As Workbook
End Function

Private Function ReturnBareWorkbook() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedBareWorkbook()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnBareWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedQualifiedWorkbook() As Excel.Workbook
End Function

Private Function ReturnQualifiedWorkbook() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedQualifiedWorkbook()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnQualifiedWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedBareCollection() As Collection
End Function

Private Function ReturnBareCollection() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedBareCollection()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnBareCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedQualifiedCollection() As VBA.Collection
End Function

Private Function ReturnQualifiedCollection() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedQualifiedCollection()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnQualifiedCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedBareApplication() As Application
End Function

Private Function ReturnBareApplication() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedBareApplication()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnBareApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function UnassignedQualifiedApplication() As Excel.Application
End Function

Private Function ReturnQualifiedApplication() As String
    Dim target As Object, made As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = UnassignedQualifiedApplication()
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ReturnQualifiedApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|check:" & CStr(Err.Number)
End Function

Private Function MakeBareCollection() As String
    Dim target As Object, made As Long, count As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = New Collection
    made = Err.Number
    Err.Clear
    count = -1
    If Not (target Is Nothing) Then
        target.Add "one"
        count = target.Count
    End If
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    MakeBareCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|count:" & CStr(count) & "|check:" & CStr(Err.Number)
End Function

Private Function MakeQualifiedCollection() As String
    Dim target As Object, made As Long, count As Long
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = New VBA.Collection
    made = Err.Number
    Err.Clear
    count = -1
    If Not (target Is Nothing) Then
        target.Add "one"
        count = target.Count
    End If
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    MakeQualifiedCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|count:" & CStr(count) & "|check:" & CStr(Err.Number)
End Function

Private Function QualifiedTypeOf(ByVal mode As Long) As String
    Dim target As Object
    On Error Resume Next
    Err.Clear
    If mode < 2 Then Set target = Range("A1") Else Set target = New Collection
    Select Case mode
        Case 0: QualifiedTypeOf = CStr(TypeOf target Is Excel.Range)
        Case 1: QualifiedTypeOf = CStr(TypeOf target Is Excel.Worksheet)
        Case 2: QualifiedTypeOf = CStr(TypeOf target Is VBA.Collection)
        Case 3: QualifiedTypeOf = CStr(TypeOf target Is Excel.Range)
    End Select
    QualifiedTypeOf = QualifiedTypeOf & "|check:" & CStr(Err.Number)
End Function

Private Function BorrowBareRange(ByRef target As Range) As String
    bodyHits = bodyHits + 1
    BorrowBareRange = TypeName(target)
    Set target = Nothing
End Function

Private Function BorrowQualifiedRange(ByRef target As Excel.Range) As String
    bodyHits = bodyHits + 1
    BorrowQualifiedRange = TypeName(target)
    Set target = Nothing
End Function

Private Function ProbeQualifiedToBareRange() As String
    Dim target As Excel.Range, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = Range("A1")
    answer = BorrowBareRange(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeQualifiedToBareRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function ProbeBareToQualifiedRange() As String
    Dim target As Range, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = Range("A1")
    answer = BorrowQualifiedRange(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeBareToQualifiedRange = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function BorrowBareWorksheet(ByRef target As Worksheet) As String
    bodyHits = bodyHits + 1
    BorrowBareWorksheet = TypeName(target)
    Set target = Nothing
End Function

Private Function BorrowQualifiedWorksheet(ByRef target As Excel.Worksheet) As String
    bodyHits = bodyHits + 1
    BorrowQualifiedWorksheet = TypeName(target)
    Set target = Nothing
End Function

Private Function ProbeQualifiedToBareWorksheet() As String
    Dim target As Excel.Worksheet, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = ActiveSheet
    answer = BorrowBareWorksheet(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeQualifiedToBareWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function ProbeBareToQualifiedWorksheet() As String
    Dim target As Worksheet, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = ActiveSheet
    answer = BorrowQualifiedWorksheet(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeBareToQualifiedWorksheet = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function BorrowBareWorkbook(ByRef target As Workbook) As String
    bodyHits = bodyHits + 1
    BorrowBareWorkbook = TypeName(target)
    Set target = Nothing
End Function

Private Function BorrowQualifiedWorkbook(ByRef target As Excel.Workbook) As String
    bodyHits = bodyHits + 1
    BorrowQualifiedWorkbook = TypeName(target)
    Set target = Nothing
End Function

Private Function ProbeQualifiedToBareWorkbook() As String
    Dim target As Excel.Workbook, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = ThisWorkbook
    answer = BorrowBareWorkbook(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeQualifiedToBareWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function ProbeBareToQualifiedWorkbook() As String
    Dim target As Workbook, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = ThisWorkbook
    answer = BorrowQualifiedWorkbook(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeBareToQualifiedWorkbook = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function BorrowBareCollection(ByRef target As Collection) As String
    bodyHits = bodyHits + 1
    BorrowBareCollection = TypeName(target)
    Set target = Nothing
End Function

Private Function BorrowQualifiedCollection(ByRef target As VBA.Collection) As String
    bodyHits = bodyHits + 1
    BorrowQualifiedCollection = TypeName(target)
    Set target = Nothing
End Function

Private Function ProbeQualifiedToBareCollection() As String
    Dim target As VBA.Collection, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = New Collection
    answer = BorrowBareCollection(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeQualifiedToBareCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function ProbeBareToQualifiedCollection() As String
    Dim target As Collection, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = New Collection
    answer = BorrowQualifiedCollection(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeBareToQualifiedCollection = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function BorrowBareApplication(ByRef target As Application) As String
    bodyHits = bodyHits + 1
    BorrowBareApplication = TypeName(target)
    Set target = Nothing
End Function

Private Function BorrowQualifiedApplication(ByRef target As Excel.Application) As String
    bodyHits = bodyHits + 1
    BorrowQualifiedApplication = TypeName(target)
    Set target = Nothing
End Function

Private Function ProbeQualifiedToBareApplication() As String
    Dim target As Excel.Application, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = Application
    answer = BorrowBareApplication(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeQualifiedToBareApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function ProbeBareToQualifiedApplication() As String
    Dim target As Application, made As Long, answer As String
    Dim kind As String, objectFlag As Boolean, nothingFlag As Boolean
    On Error Resume Next
    Err.Clear
    Set target = Application
    answer = BorrowQualifiedApplication(target)
    made = Err.Number
    Err.Clear
    kind = TypeName(target)
    objectFlag = IsObject(target)
    nothingFlag = (target Is Nothing)
    ProbeBareToQualifiedApplication = CStr(made) & "|" & kind & "|" & CStr(objectFlag) & "|" & CStr(nothingFlag) & "|answer:" & answer & "|hits:" & CStr(bodyHits) & "|check:" & CStr(Err.Number)
End Function

Private Function QualifiedObjectFormsCase(ByVal form As Long) As String
    Select Case form
        Case 0: QualifiedObjectFormsCase = LocalBareRange(0)
        Case 1: QualifiedObjectFormsCase = LocalBareRange(1)
        Case 2: QualifiedObjectFormsCase = LocalBareRange(2)
        Case 3: QualifiedObjectFormsCase = LocalBareRange(3)
        Case 4: QualifiedObjectFormsCase = LocalQualifiedRange(0)
        Case 5: QualifiedObjectFormsCase = LocalQualifiedRange(1)
        Case 6: QualifiedObjectFormsCase = LocalQualifiedRange(2)
        Case 7: QualifiedObjectFormsCase = LocalQualifiedRange(3)
        Case 8: QualifiedObjectFormsCase = LocalBareWorksheet(0)
        Case 9: QualifiedObjectFormsCase = LocalBareWorksheet(1)
        Case 10: QualifiedObjectFormsCase = LocalBareWorksheet(2)
        Case 11: QualifiedObjectFormsCase = LocalBareWorksheet(3)
        Case 12: QualifiedObjectFormsCase = LocalQualifiedWorksheet(0)
        Case 13: QualifiedObjectFormsCase = LocalQualifiedWorksheet(1)
        Case 14: QualifiedObjectFormsCase = LocalQualifiedWorksheet(2)
        Case 15: QualifiedObjectFormsCase = LocalQualifiedWorksheet(3)
        Case 16: QualifiedObjectFormsCase = LocalBareWorkbook(0)
        Case 17: QualifiedObjectFormsCase = LocalBareWorkbook(1)
        Case 18: QualifiedObjectFormsCase = LocalBareWorkbook(2)
        Case 19: QualifiedObjectFormsCase = LocalBareWorkbook(3)
        Case 20: QualifiedObjectFormsCase = LocalQualifiedWorkbook(0)
        Case 21: QualifiedObjectFormsCase = LocalQualifiedWorkbook(1)
        Case 22: QualifiedObjectFormsCase = LocalQualifiedWorkbook(2)
        Case 23: QualifiedObjectFormsCase = LocalQualifiedWorkbook(3)
        Case 24: QualifiedObjectFormsCase = LocalBareCollection(0)
        Case 25: QualifiedObjectFormsCase = LocalBareCollection(1)
        Case 26: QualifiedObjectFormsCase = LocalBareCollection(2)
        Case 27: QualifiedObjectFormsCase = LocalBareCollection(3)
        Case 28: QualifiedObjectFormsCase = LocalQualifiedCollection(0)
        Case 29: QualifiedObjectFormsCase = LocalQualifiedCollection(1)
        Case 30: QualifiedObjectFormsCase = LocalQualifiedCollection(2)
        Case 31: QualifiedObjectFormsCase = LocalQualifiedCollection(3)
        Case 32: QualifiedObjectFormsCase = LocalBareApplication(0)
        Case 33: QualifiedObjectFormsCase = LocalBareApplication(1)
        Case 34: QualifiedObjectFormsCase = LocalBareApplication(2)
        Case 35: QualifiedObjectFormsCase = LocalBareApplication(3)
        Case 36: QualifiedObjectFormsCase = LocalQualifiedApplication(0)
        Case 37: QualifiedObjectFormsCase = LocalQualifiedApplication(1)
        Case 38: QualifiedObjectFormsCase = LocalQualifiedApplication(2)
        Case 39: QualifiedObjectFormsCase = LocalQualifiedApplication(3)
        Case 40: QualifiedObjectFormsCase = ReturnBareRange()
        Case 41: QualifiedObjectFormsCase = ReturnQualifiedRange()
        Case 42: QualifiedObjectFormsCase = ReturnBareWorksheet()
        Case 43: QualifiedObjectFormsCase = ReturnQualifiedWorksheet()
        Case 44: QualifiedObjectFormsCase = ReturnBareWorkbook()
        Case 45: QualifiedObjectFormsCase = ReturnQualifiedWorkbook()
        Case 46: QualifiedObjectFormsCase = ReturnBareCollection()
        Case 47: QualifiedObjectFormsCase = ReturnQualifiedCollection()
        Case 48: QualifiedObjectFormsCase = ReturnBareApplication()
        Case 49: QualifiedObjectFormsCase = ReturnQualifiedApplication()
        Case 50: QualifiedObjectFormsCase = MakeBareCollection()
        Case 51: QualifiedObjectFormsCase = MakeQualifiedCollection()
        Case 52: QualifiedObjectFormsCase = QualifiedTypeOf(0)
        Case 53: QualifiedObjectFormsCase = QualifiedTypeOf(1)
        Case 54: QualifiedObjectFormsCase = QualifiedTypeOf(2)
        Case 55: QualifiedObjectFormsCase = QualifiedTypeOf(3)
        Case 56: QualifiedObjectFormsCase = ProbeQualifiedToBareRange()
        Case 57: QualifiedObjectFormsCase = ProbeBareToQualifiedRange()
        Case 58: QualifiedObjectFormsCase = ProbeQualifiedToBareWorksheet()
        Case 59: QualifiedObjectFormsCase = ProbeBareToQualifiedWorksheet()
        Case 60: QualifiedObjectFormsCase = ProbeQualifiedToBareWorkbook()
        Case 61: QualifiedObjectFormsCase = ProbeBareToQualifiedWorkbook()
        Case 62: QualifiedObjectFormsCase = ProbeQualifiedToBareCollection()
        Case 63: QualifiedObjectFormsCase = ProbeBareToQualifiedCollection()
        Case 64: QualifiedObjectFormsCase = ProbeQualifiedToBareApplication()
        Case 65: QualifiedObjectFormsCase = ProbeBareToQualifiedApplication()
    End Select
End Function

