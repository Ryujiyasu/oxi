' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit

Private Function ErrorSnapshot() As String
    ErrorSnapshot = CStr(Err.Number) & vbTab & Err.Source & vbTab & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & vbTab & Err.HelpFile & vbTab & CStr(Err.HelpContext)
End Function

Private Sub NestedLocalizedRaise()
    Err.Raise 513
End Sub

Public Function VbaDefaultHelpFile() As String
    On Error Resume Next
    Err.Clear
    Err.Raise 6
    VbaDefaultHelpFile = Err.HelpFile
End Function

Public Function Cleared() As String
    Dim value As Double
    On Error Resume Next
    Err.Clear
    Cleared = ErrorSnapshot()
End Function

Public Function NumberOnly() As String
    Dim value As Double
    On Error Resume Next
    Err.Number = 6
    NumberOnly = ErrorSnapshot()
End Function

Public Function DivideZero() As String
    Dim value As Double
    On Error Resume Next
    value = 1 / 0
    DivideZero = ErrorSnapshot()
End Function

Public Function Overflow() As String
    Dim value As Double
    On Error Resume Next
    value = CInt(32768)
    Overflow = ErrorSnapshot()
End Function

Public Function TypeMismatch() As String
    Dim value As Double
    On Error Resume Next
    value = CLng("bad")
    TypeMismatch = ErrorSnapshot()
End Function

Public Function InvalidNull() As String
    Dim value As Double
    On Error Resume Next
    value = CLng(Null)
    InvalidNull = ErrorSnapshot()
End Function

Public Function Domain() As String
    Dim value As Double
    On Error Resume Next
    value = Sqr(-1)
    Domain = ErrorSnapshot()
End Function

Public Function ArrayBounds() As String
    Dim value As Double
    On Error Resume Next
    Dim items(1 To 1) As Long
    value = items(2)
    ArrayBounds = ErrorSnapshot()
End Function

Public Function CollectionDuplicate() As String
    Dim value As Double
    On Error Resume Next
    Dim items As New Collection
    items.Add 1, "key"
    items.Add 2, "key"
    CollectionDuplicate = ErrorSnapshot()
End Function

Public Function RaiseBuiltin() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6
    RaiseBuiltin = ErrorSnapshot()
End Function

Public Function RaiseUnknown() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513
    RaiseUnknown = ErrorSnapshot()
End Function

Public Function RaiseNegative() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise -5
    RaiseNegative = ErrorSnapshot()
End Function

Public Function RaiseZero() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 0
    RaiseZero = ErrorSnapshot()
End Function

Public Function RaiseHelpOverflow() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, , , , 2147483648#
    RaiseHelpOverflow = ErrorSnapshot()
End Function

Public Function RaiseNamed() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise Description:="独自の説明", Number:=513
    RaiseNamed = ErrorSnapshot()
End Function

Public Function RaiseEnglishWords() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "custom", "Overflow"
    RaiseEnglishWords = ErrorSnapshot()
End Function

Public Function RaiseEmptyWords() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, , ""
    RaiseEmptyWords = ErrorSnapshot()
End Function

Public Function RaiseEnglishGenericWords() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, , "Application-defined or object-defined error"
    RaiseEnglishGenericWords = ErrorSnapshot()
End Function

Public Function RaiseCatalogWords() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, , Error(513)
    RaiseCatalogWords = ErrorSnapshot()
End Function

Public Function RaiseExplicitHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "custom", "description", "guide.chm", 42
    RaiseExplicitHelp = ErrorSnapshot()
End Function

Public Function RaiseExplicitEmptyHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6, , , ""
    RaiseExplicitEmptyHelp = ErrorSnapshot()
End Function

Public Function RaiseAssignedEmptyHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpFile = ""
    Err.Raise 6
    RaiseAssignedEmptyHelp = ErrorSnapshot()
End Function

Public Function RaiseAssignedCustomHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpFile = "manual.chm"
    Err.Raise 6
    RaiseAssignedCustomHelp = ErrorSnapshot()
End Function

Public Function RaiseContextOnly() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpContext = 42
    Err.Raise 513
    RaiseContextOnly = ErrorSnapshot()
End Function

Public Function RaiseAssignedEmptyDescription() As String
    Dim value As Double
    On Error Resume Next
    Err.Description = ""
    Err.Raise 513
    RaiseAssignedEmptyDescription = ErrorSnapshot()
End Function

Public Function RaiseAssignedEmptySource() As String
    Dim value As Double
    On Error Resume Next
    Err.Source = ""
    Err.Raise 513
    RaiseAssignedEmptySource = ErrorSnapshot()
End Function

Public Function RaiseAfterClear() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpFile = "manual.chm"
    Err.Clear
    Err.Raise 6
    RaiseAfterClear = ErrorSnapshot()
End Function

Public Function InternalAfterCustomHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpFile = "manual.chm"
    value = 1 / 0
    InternalAfterCustomHelp = ErrorSnapshot()
End Function

Public Function ErrorStatement() As String
    Dim value As Double
    On Error Resume Next
    Error 11
    ErrorStatement = ErrorSnapshot()
End Function

Public Function ErrorStatementAfterCustom() As String
    Dim value As Double
    On Error Resume Next
    Err.Source = "custom": Err.Description = "custom": Err.HelpFile = "manual.chm": Err.HelpContext = 42
    Error 513
    ErrorStatementAfterCustom = ErrorSnapshot()
End Function

Public Function ClearAfterRaise() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6
    Err.Clear
    ClearAfterRaise = ErrorSnapshot()
End Function

Public Function ClearFileAfterRaise() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6
    Err.HelpFile = ""
    ClearFileAfterRaise = ErrorSnapshot()
End Function

Public Function ClearContextAfterRaise() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6
    Err.HelpContext = 0
    ClearContextAfterRaise = ErrorSnapshot()
End Function

Public Function CallByNameBuiltin() As String
    Dim value As Double
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, 6, "custom", "custom", "guide.chm", 42
    CallByNameBuiltin = ErrorSnapshot()
End Function

Public Function CallByNameUnknown() As String
    Dim value As Double
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, 513, "custom", "custom", "guide.chm", 42
    CallByNameUnknown = ErrorSnapshot()
End Function

Public Function NestedRaise() As String
    Dim value As Double
    On Error Resume Next
    NestedLocalizedRaise
    NestedRaise = ErrorSnapshot()
End Function

Public Function NestedEmptyHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.HelpFile = ""
    NestedLocalizedRaise
    NestedEmptyHelp = ErrorSnapshot()
End Function

Public Function HostRangeOffset() As String
    Dim value As Double
    On Error Resume Next
    value = Range("A1").Offset(-1, 0).Value
    HostRangeOffset = ErrorSnapshot()
End Function

Public Function HostWorksheetIndex() As String
    Dim value As Double
    On Error Resume Next
    Dim sheet As Object
    Set sheet = Worksheets(999)
    HostWorksheetIndex = ErrorSnapshot()
End Function

Public Function HostWorksheetFunction() As String
    Dim value As Double
    On Error Resume Next
    value = Application.WorksheetFunction.Match("absent", Range("A1:A2"), 0)
    HostWorksheetFunction = ErrorSnapshot()
End Function

Private Sub NestedSourceRaise()
    Err.Raise 513, "Collection", "custom"
End Sub

Public Function SourceVba() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "VBA", "custom"
    SourceVba = ErrorSnapshot()
End Function

Public Function SourceCollection() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "Collection", "custom"
    SourceCollection = ErrorSnapshot()
End Function

Public Function SourceExcel() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "Microsoft Excel", "custom"
    SourceExcel = ErrorSnapshot()
End Function

Public Function SourceLowercase() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "vba", "custom"
    SourceLowercase = ErrorSnapshot()
End Function

Public Function SourceBuiltinDefault() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 6, "VBA"
    SourceBuiltinDefault = ErrorSnapshot()
End Function

Public Function SourceAssigned() As String
    Dim value As Double
    On Error Resume Next
    Err.Source = "VBA"
    Err.Description = "custom"
    Err.Raise 513
    SourceAssigned = ErrorSnapshot()
End Function

Public Function SourceNamed() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise Source:="Collection", Description:="custom", Number:=513
    SourceNamed = ErrorSnapshot()
End Function

Public Function SourceEmptyWords() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "VBA", ""
    SourceEmptyWords = ErrorSnapshot()
End Function

Public Function SourceNested() As String
    Dim value As Double
    On Error Resume Next
    NestedSourceRaise
    SourceNested = ErrorSnapshot()
End Function

Public Function SourceWithHelp() As String
    Dim value As Double
    On Error Resume Next
    Err.Raise 513, "Collection", "custom", "guide.chm", 42
    SourceWithHelp = ErrorSnapshot()
End Function

Public Function CollectionMissingIndex() As String
    Dim value As Double
    On Error Resume Next
    Dim items As New Collection
    items.Add 1
    value = items(2)
    CollectionMissingIndex = ErrorSnapshot()
End Function

Public Function CollectionMissingKey() As String
    Dim value As Double
    On Error Resume Next
    Dim items As New Collection
    items.Add 1
    value = items("absent")
    CollectionMissingKey = ErrorSnapshot()
End Function

Public Function FileMissing() As String
    Dim value As Double
    On Error Resume Next
    Open "__oxi_vba_absent_3547b203cffb4d6f.txt" For Input As #1
    FileMissing = ErrorSnapshot()
End Function

Public Function DictionaryDuplicate() As String
    Dim value As Double
    On Error Resume Next
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    items.Add "key", 1
    items.Add "key", 2
    DictionaryDuplicate = ErrorSnapshot()
End Function

Public Function DictionaryMissingKey() As String
    Dim value As Double
    On Error Resume Next
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    items.Remove "absent"
    DictionaryMissingKey = ErrorSnapshot()
End Function

Public Function SourceControlDescription() As String
    On Error Resume Next
    Err.Raise 513, "VBA", "before" & Chr(1) & "after"
    SourceControlDescription = ErrorSnapshot()
End Function

Public Function CustomControlDescription() As String
    On Error Resume Next
    Err.Raise 513, "custom", "before" & Chr(1) & "after"
    CustomControlDescription = ErrorSnapshot()
End Function

Public Function CallByNameVariantNumber() As String
    Dim value As Double
    On Error Resume Next
    Dim target As Variant
    target = 1
    value = CallByName(target, "Foo", VbGet)
    CallByNameVariantNumber = ErrorSnapshot()
End Function

Public Function DictionaryUnknownNamed() As String
    Dim value As Double
    On Error Resume Next
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    items.Add Unknown:=1
    DictionaryUnknownNamed = ErrorSnapshot()
End Function

Public Function CollectionUnknownNamed() As String
    Dim value As Double
    On Error Resume Next
    Dim items As Object
    Set items = New Collection
    items.Add Unknown:=1
    CollectionUnknownNamed = ErrorSnapshot()
End Function
