' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit

Public Function HelpFile() As String
    On Error Resume Next
    Err.HelpFile = "guide.chm"
    HelpFile = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function HelpContext() As String
    On Error Resume Next
    Err.HelpContext = 42
    HelpContext = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function HelpContextRounded() As String
    On Error Resume Next
    Err.HelpContext = 2.5
    HelpContextRounded = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function HelpContextNegative() As String
    On Error Resume Next
    Err.HelpContext = -1
    HelpContextNegative = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function HelpContextOverflow() As String
    On Error Resume Next
    Err.HelpContext = 2147483648#
    HelpContextOverflow = CStr(Err.Number)
End Function

Public Function HelpContextNull() As String
    On Error Resume Next
    Err.HelpContext = Null
    HelpContextNull = CStr(Err.Number)
End Function

Public Function HelpContextBad() As String
    On Error Resume Next
    Err.HelpContext = "abc"
    HelpContextBad = CStr(Err.Number)
End Function

Public Function HelpFileNull() As String
    On Error Resume Next
    Err.HelpFile = Null
    HelpFileNull = CStr(Err.Number)
End Function

Public Function SourceNull() As String
    On Error Resume Next
    Err.Source = Null
    SourceNull = CStr(Err.Number)
End Function

Public Function DescriptionNull() As String
    On Error Resume Next
    Err.Description = Null
    DescriptionNull = CStr(Err.Number)
End Function

Public Function NumberNull() As String
    On Error Resume Next
    Err.Number = Null
    NumberNull = CStr(Err.Number)
End Function

Public Function NumberBad() As String
    On Error Resume Next
    Err.Number = "abc"
    NumberBad = CStr(Err.Number)
End Function

Public Function NumberZero() As String
    On Error Resume Next
    Err.Number = 0
    NumberZero = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NumberNegative() As String
    On Error Resume Next
    Err.Number = -5
    NumberNegative = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function Number65535() As String
    On Error Resume Next
    Err.Number = 65535
    Number65535 = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function Number65536() As String
    On Error Resume Next
    Err.Number = 65536
    Number65536 = CStr(Err.Number)
End Function

Public Function NumberOverflow() As String
    On Error Resume Next
    Err.Number = 2147483648#
    NumberOverflow = CStr(Err.Number)
End Function

Public Function NumberUnderflow() As String
    On Error Resume Next
    Err.Number = -2147483649#
    NumberUnderflow = CStr(Err.Number)
End Function

Public Function NumberRounded() As String
    On Error Resume Next
    Err.Number = 2.5
    NumberRounded = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NumberTrue() As String
    On Error Resume Next
    Err.Number = True
    NumberTrue = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NumberDefault() As String
    On Error Resume Next
    Err = 13
    NumberDefault = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NumberKeepsWords() As String
    On Error Resume Next
    Err.Source = "src": Err.Description = "desc": Err.HelpFile = "guide": Err.HelpContext = 42: Err.Number = 13
    NumberKeepsWords = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function ClearHelp() As String
    On Error Resume Next
    Err.HelpFile = "guide": Err.HelpContext = 42: Err.Clear
    ClearHelp = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseHelp() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42
    RaiseHelp = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseKeepsHelp() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42: Err.Raise 514
    RaiseKeepsHelp = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseEmptyHelp() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42: Err.Raise 514, , , "", 0
    RaiseEmptyHelp = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseNull() As String
    On Error Resume Next
    Err.Raise Null
    RaiseNull = CStr(Err.Number)
End Function

Public Function RaiseOverflow() As String
    On Error Resume Next
    Err.Raise 2147483648#
    RaiseOverflow = CStr(Err.Number)
End Function

Public Function RaiseUnderflow() As String
    On Error Resume Next
    Err.Raise -2147483649#
    RaiseUnderflow = CStr(Err.Number)
End Function

Public Function Raise65536() As String
    On Error Resume Next
    Err.Raise 65536
    Raise65536 = CStr(Err.Number)
End Function

Public Function RaiseNegative() As String
    On Error Resume Next
    Err.Raise -5
    RaiseNegative = CStr(Err.Number)
End Function

Public Function RaiseHelpOverflow() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 2147483648#
    RaiseHelpOverflow = CStr(Err.Number)
End Function

Public Function RaiseHelpNull() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", Null
    RaiseHelpNull = CStr(Err.Number)
End Function

Public Function RaiseSourceNull() As String
    On Error Resume Next
    Err.Raise 513, Null, "desc"
    RaiseSourceNull = CStr(Err.Number)
End Function

Public Function RaiseDescriptionNull() As String
    On Error Resume Next
    Err.Raise 513, "src", Null
    RaiseDescriptionNull = CStr(Err.Number)
End Function

Public Function RaiseHelpFileNull() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", Null
    RaiseHelpFileNull = CStr(Err.Number)
End Function

Public Function RaiseNamed() As String
    On Error Resume Next
    Err.Raise Description:="desc", Number:=513, HelpContext:=42, HelpFile:="guide", Source:="src"
    RaiseNamed = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function WithProperties() As String
    On Error Resume Next
    With Err
        .Number = 13
        .Source = "src"
        .Description = "desc"
        .HelpFile = "guide"
        .HelpContext = 42
    End With
    WithProperties = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function WithRaise() As String
    On Error Resume Next
    With Err
        .Raise 513, "src", "desc", "guide", 42
    End With
    WithRaise = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function WithRaiseNamed() As String
    On Error Resume Next
    With Err
        .Raise Description:="desc", Number:=513, HelpContext:=42, HelpFile:="guide", Source:="src"
    End With
    WithRaiseNamed = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function WithRead() As String
    On Error Resume Next
    With Err
        .Number = 13
        WithRead = CStr(.Number)
    End With
    Exit Function
    WithRead = ""
End Function

Public Function WithClear() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42
    With Err
        .Clear
    End With
    WithClear = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function ObjectAlias() As String
    On Error Resume Next
    Dim e As Object
    Set e = Err
    e.Number = 13
    e.Source = "src"
    e.Description = "desc"
    e.HelpFile = "guide"
    e.HelpContext = 42
    ObjectAlias = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function ObjectAliasNamed() As String
    On Error Resume Next
    Dim e As Object
    Set e = Err
    e.Raise Description:="desc", Number:=513, HelpContext:=42, HelpFile:="guide", Source:="src"
    ObjectAliasNamed = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function ObjectAliasClear() As String
    On Error Resume Next
    Dim e As Object
    Set e = Err
    Err.Raise 513, "src", "desc", "guide", 42
    e.Clear
    ObjectAliasClear = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseEmptyNumber() As String
    On Error Resume Next
    Err.Raise Empty
    RaiseEmptyNumber = CStr(Err.Number)
End Function

Public Function RaiseZeroNumber() As String
    On Error Resume Next
    Err.Raise 0
    RaiseZeroNumber = CStr(Err.Number)
End Function

Public Function RaiseNumericText() As String
    On Error Resume Next
    Err.Raise "513", "src", "desc", "guide", 42
    RaiseNumericText = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseContextFraction() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 2.5
    RaiseContextFraction = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseContextNegative() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", -1
    RaiseContextNegative = CStr(Err.Number)
End Function

Public Function RaiseExplicitEmptySource() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42
    Err.Raise 514, ""
    RaiseExplicitEmptySource = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function RaiseExplicitEmptyDescription() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42
    Err.Raise 514, , ""
    RaiseExplicitEmptyDescription = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function SourceFromNumber() As String
    On Error Resume Next
    Err.Source = 42
    SourceFromNumber = Err.Source
End Function

Public Function HelpFileFromNumber() As String
    On Error Resume Next
    Err.HelpFile = 42
    HelpFileFromNumber = Err.HelpFile
End Function

Public Function DescriptionFromNumber() As String
    On Error Resume Next
    Err.Description = 42
    DescriptionFromNumber = Err.Description
End Function

Public Function RaiseEmptySource() As String
    On Error Resume Next
    Err.Raise 513, Empty, "desc", "guide", 42
    RaiseEmptySource = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseEmptyDescription() As String
    On Error Resume Next
    Err.Raise 513, "src", Empty, "guide", 42
    RaiseEmptyDescription = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseEmptyHelpFile() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", Empty, 42
    RaiseEmptyHelpFile = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseManuallySetMetadata() As String
    On Error Resume Next
    Err.Source = "src": Err.Description = "desc": Err.HelpFile = "guide": Err.HelpContext = 42: Err.Raise 513
    RaiseManuallySetMetadata = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseMissingSource() As String
    On Error Resume Next
    Err.Raise 513, , "desc", "guide", 42
    RaiseMissingSource = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseEmptyContext() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", Empty
    RaiseEmptyContext = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function RaiseNegativeContextValue() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", -1
    RaiseNegativeContextValue = CStr(Err.HelpContext)
End Function

Private Sub ThrowNested()
    Err.Raise 513, "src", "desc", "guide", 42
End Sub

Private Sub RethrowNested()
    On Error GoTo Handler
    ThrowNested
    Exit Sub
Handler:
    Err.Raise 514
End Sub

Private Sub RaiseInsideHandler()
    On Error GoTo Handler
    ThrowNested
    Exit Sub
Handler:
    Err.Raise 514, "src2", "desc2", "guide2", 84
End Sub

Public Function ErrIsObject() As String
    On Error Resume Next
    
    ErrIsObject = CStr(IsObject(Err))
End Function

Public Function ErrTypeName() As String
    On Error Resume Next
    
    ErrTypeName = TypeName(Err)
End Function

Public Function ErrVarType() As String
    On Error Resume Next
    
    ErrVarType = CStr(VarType(Err))
End Function

Public Function ErrIdentity() As String
    On Error Resume Next
    
    ErrIdentity = CStr(Err Is Err)
End Function

Public Function ErrDefaultAsString() As String
    On Error Resume Next
    Err.Number = 13
    ErrDefaultAsString = CStr(Err)
End Function

Public Function ErrFunctionAsString() As String
    On Error Resume Next
    Err.Number = 13
    ErrFunctionAsString = CStr(Err())
End Function

Public Function AliasAsString() As String
    On Error Resume Next
    Dim e As Object
    Set e = Err
    Err.Number = 13
    AliasAsString = CStr(e)
End Function

Public Function AliasLetObject() As String
    On Error Resume Next
    Dim e As Object
    Set e = Err
    e = 13
    AliasLetObject = CStr(Err.Number)
End Function

Public Function AliasLetVariant() As String
    On Error Resume Next
    Dim e As Variant
    Set e = Err
    e = 13
    AliasLetVariant = CStr(VarType(e)) & "|" & CStr(Err.Number)
End Function

Public Function ErrCallByNameGet() As String
    On Error Resume Next
    Err.Number = 13
    ErrCallByNameGet = CStr(CallByName(Err, "Number", VbGet))
End Function

Public Function ErrCallByNameLet() As String
    On Error Resume Next
    CallByName Err, "Number", VbLet, 13
    ErrCallByNameLet = CStr(Err.Number)
End Function

Public Function ErrCallByNameHelp() As String
    On Error Resume Next
    CallByName Err, "HelpFile", VbLet, "guide"
    CallByName Err, "HelpContext", VbLet, 42
    ErrCallByNameHelp = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function ErrCallByNameRaise() As String
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, 513, "src", "desc", "guide", 42
    ErrCallByNameRaise = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(Err.Number)) & "|" & CStr(Err.HelpContext)
End Function

Public Function ErrCallByNameClear() As String
    On Error Resume Next
    Err.Raise 513, "src", "desc", "guide", 42
    CallByName Err, "Clear", VbMethod
    ErrCallByNameClear = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NestedRaiseResumeNext() As String
    On Error Resume Next
    ThrowNested
    NestedRaiseResumeNext = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NestedRethrow() As String
    On Error Resume Next
    RethrowNested
    NestedRethrow = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function NestedActiveHandler() As String
    On Error Resume Next
    RaiseInsideHandler
    NestedActiveHandler = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function

Public Function CallByNameRaiseBuiltin() As String
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, 6, "src", "desc", "guide", 42
    CallByNameRaiseBuiltin = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(Err.Number)) & "|" & CStr(Err.HelpContext)
End Function

Public Function CallByNameRaiseUser() As String
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, 514, "src", "desc", "guide", 42
    CallByNameRaiseUser = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(Err.Number)) & "|" & CStr(Err.HelpContext)
End Function

Public Function CallByNameRaiseNegative() As String
    On Error Resume Next
    CallByName Err, "Raise", VbMethod, -5, "src", "desc", "guide", 42
    CallByNameRaiseNegative = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.HelpContext)
End Function

Private Function LongParameterType(ByVal value As Long) As String
    LongParameterType = TypeName(value)
End Function

Private Function VariantParameterType(ByVal value As Variant) As String
    VariantParameterType = TypeName(value)
End Function

Public Function ErrCondition() As String
    On Error Resume Next
    Err.Number = 13
    If Err Then
        ErrCondition = "yes"
    Else
        ErrCondition = "no"
    End If
    Exit Function
    ErrCondition = ""
End Function

Public Function ErrSelectCase() As String
    On Error Resume Next
    Err.Number = 13
    Select Case Err
    Case 13
        ErrSelectCase = "matched"
    Case Else
        ErrSelectCase = "missed"
    End Select
    Exit Function
    ErrSelectCase = ""
End Function

Public Function ErrForBound() As String
    On Error Resume Next
    Dim i As Long, total As Long
    Err.Number = 3
    For i = 1 To Err
        total = total + i
    Next i
    ErrForBound = CStr(total)
End Function

Public Function ErrDoCondition() As String
    On Error Resume Next
    Dim total As Long
    Err.Number = 3
    Do While Err
        total = total + 1
        Err.Number = Err.Number - 1
    Loop
    ErrDoCondition = CStr(total)
End Function

Public Function ErrArrayIndex() As String
    On Error Resume Next
    Dim values(1 To 3) As String
    values(2) = "two"
    Err.Number = 2
    ErrArrayIndex = values(Err)
End Function

Public Function ErrMidStart() As String
    On Error Resume Next
    Dim text As String
    text = "abcd"
    Err.Number = 2
    Mid(text, Err, 1) = "X"
    ErrMidStart = text
End Function

Public Function ErrLongParameter() As String
    On Error Resume Next
    Err.Number = 13
    ErrLongParameter = LongParameterType(Err)
End Function

Public Function ErrVariantParameter() As String
    On Error Resume Next
    Err.Number = 13
    ErrVariantParameter = VariantParameterType(Err)
End Function

Public Function ErrErrorArgument() As String
    On Error Resume Next
    Err.Number = 13
    ErrErrorArgument = CStr(Error(Err) = Error(13))
End Function

Private Function ByRefLongParameterType(ByRef value As Long) As String
    ByRefLongParameterType = TypeName(value)
End Function

Public Function ErrByRefLongParameter() As String
    On Error Resume Next
    Err.Number = 13
    ErrByRefLongParameter = ByRefLongParameterType(Err)
End Function

Public Function ContextUnknownDefault() As String
    On Error Resume Next
    Err.Raise 513
    ContextUnknownDefault = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownCustom() As String
    On Error Resume Next
    Err.Raise 513, , "custom"
    ContextUnknownCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownEmpty() As String
    On Error Resume Next
    Err.Raise 513, , ""
    ContextUnknownEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownEmptyVariant() As String
    On Error Resume Next
    Err.Raise 513, , Empty
    ContextUnknownEmptyVariant = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownCatalogText() As String
    On Error Resume Next
    Err.Raise 513, , Error(513)
    ContextUnknownCatalogText = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownOtherCatalogText() As String
    On Error Resume Next
    Err.Raise 513, , Error(6)
    ContextUnknownOtherCatalogText = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextBuiltinCustom() As String
    On Error Resume Next
    Err.Raise 6, , "custom"
    ContextBuiltinCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextBuiltinEmpty() As String
    On Error Resume Next
    Err.Raise 6, , ""
    ContextBuiltinEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownContextZero() As String
    On Error Resume Next
    Err.Raise 513, , , , 0
    ContextUnknownContextZero = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownContextCustom() As String
    On Error Resume Next
    Err.Raise 513, , , , 42
    ContextUnknownContextCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedDescription() As String
    On Error Resume Next
    Err.Description = "inherited"
    Err.Raise 513
    ContextUnknownInheritedDescription = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedEmpty() As String
    On Error Resume Next
    Err.Description = ""
    Err.Raise 513
    ContextUnknownInheritedEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedGeneric() As String
    On Error Resume Next
    Err.Description = Error(513)
    Err.Raise 513
    ContextUnknownInheritedGeneric = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedBuiltin() As String
    On Error Resume Next
    Err.Description = Error(6)
    Err.Raise 513
    ContextUnknownInheritedBuiltin = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedContext() As String
    On Error Resume Next
    Err.HelpContext = 42
    Err.Raise 513
    ContextUnknownInheritedContext = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedFile() As String
    On Error Resume Next
    Err.HelpFile = "guide"
    Err.Raise 513
    ContextUnknownInheritedFile = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownFileEmpty() As String
    On Error Resume Next
    Err.Raise 513, , , ""
    ContextUnknownFileEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownSourceEmpty() As String
    On Error Resume Next
    Err.Raise 513, ""
    ContextUnknownSourceEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedSourceEmpty() As String
    On Error Resume Next
    Err.Source = ""
    Err.Raise 513
    ContextUnknownInheritedSourceEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownInheritedSource() As String
    On Error Resume Next
    Err.Source = "inherited"
    Err.Raise 513
    ContextUnknownInheritedSource = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownDescriptionRead() As String
    On Error Resume Next
    Dim prior As String
    prior = Err.Description
    Err.Raise 513
    ContextUnknownDescriptionRead = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownAfterEmptyRaise() As String
    On Error Resume Next
    Err.Raise 6, , ""
    Err.Raise 513
    ContextUnknownAfterEmptyRaise = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownAfterEmptyRaiseContextReset() As String
    On Error Resume Next
    Err.Raise 6, , ""
    Err.HelpContext = 0
    Err.Raise 513
    ContextUnknownAfterEmptyRaiseContextReset = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextUnknownAfterClear() As String
    On Error Resume Next
    Err.Description = ""
    Err.Clear
    Err.Raise 513
    ContextUnknownAfterClear = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextErrorAfterCustom() As String
    On Error Resume Next
    Err.Source = "src": Err.Description = "desc": Err.HelpFile = "guide": Err.HelpContext = 42
    Error 513
    ContextErrorAfterCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextInternalAfterCustom() As String
    On Error Resume Next
    Dim value As Double
    Err.Source = "src": Err.Description = "desc": Err.HelpFile = "guide": Err.HelpContext = 42
    value = 1 / 0
    ContextInternalAfterCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNegativeDefault() As String
    On Error Resume Next
    Err.Raise -5
    ContextNegativeDefault = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNegativeCustom() As String
    On Error Resume Next
    Err.Raise -5, , "custom"
    ContextNegativeCustom = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNegativeEmpty() As String
    On Error Resume Next
    Err.Raise -5, , ""
    ContextNegativeEmpty = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNegativeExplicitContextZero() As String
    On Error Resume Next
    Err.Raise -5, , "custom", , 0
    ContextNegativeExplicitContextZero = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextDivideZero() As String
    On Error Resume Next
    Dim value As Double
    value = 1 / 0
    ContextDivideZero = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextOverflow() As String
    On Error Resume Next
    Dim value As Integer
    value = CInt(32768)
    ContextOverflow = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNull() As String
    On Error Resume Next
    Dim value As Long
    value = CLng(Null)
    ContextNull = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextTypeMismatch() As String
    On Error Resume Next
    Dim value As Long
    value = CLng("bad")
    ContextTypeMismatch = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextSubscript() As String
    On Error Resume Next
    Dim values(1 To 1) As Long
    Dim value As Long
    value = values(2)
    ContextSubscript = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextDomain() As String
    On Error Resume Next
    Dim value As Double
    value = Sqr(-1)
    ContextDomain = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextClearedAfterRaise() As String
    On Error Resume Next
    Err.Raise 6
    Err.Clear
    ContextClearedAfterRaise = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextZeroAfterRaise() As String
    On Error Resume Next
    Err.Raise 6
    Err.HelpContext = 0
    ContextZeroAfterRaise = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Public Function ContextNestedRaise() As String
    On Error Resume Next
    Err.Description = ""
    ContextNestedHelper
    ContextNestedRaise = CStr(Err.Number) & "|" & Err.Source & "|" & CStr(Err.Description = Error(95)) & "|" & CStr(Err.Description = "") & "|" & CStr(Err.HelpContext)
End Function

Private Sub ContextNestedHelper()
    Err.Raise 513
End Sub

Public Function RaiseContextCatalog() As String
    On Error Resume Next
    Dim code As Variant, result As String
    For Each code In Array(3, 5, 6, 7, 9, 10, 11, 13, 14, 16, 17, 18, 20, 28, 35, 47, 48, 49, 51, 52, 53, 54, 55, 57, 58, 59, 61, 62, 63, 67, 68, 70, 71, 74, 75, 76, 91, 92, 93, 94, 95, 96, 97, 98, 321, 322, 325, 380, 381, 382, 383, 385, 387, 393, 394, 422, 423, 424, 429, 430, 432, 438, 440, 442, 443, 445, 446, 447, 448, 449, 450, 451, 452, 453, 454, 455, 457, 458, 459, 460, 461, 462, 463, 481, 482, 735, 744, 746)
        Err.Clear
        Err.Raise code
        result = result & CStr(code) & ":" & CStr(Err.HelpContext) & ";"
    Next code
    RaiseContextCatalog = result
End Function

Public Function ErrorContextCatalog() As String
    On Error Resume Next
    Dim code As Variant, result As String
    For Each code In Array(3, 5, 6, 7, 9, 10, 11, 13, 14, 16, 17, 18, 20, 28, 35, 47, 48, 49, 51, 52, 53, 54, 55, 57, 58, 59, 61, 62, 63, 67, 68, 70, 71, 74, 75, 76, 91, 92, 93, 94, 95, 96, 97, 98, 321, 322, 325, 380, 381, 382, 383, 385, 387, 393, 394, 422, 423, 424, 429, 430, 432, 438, 440, 442, 443, 445, 446, 447, 448, 449, 450, 451, 452, 453, 454, 455, 457, 458, 459, 460, 461, 462, 463, 481, 482, 735, 744, 746)
        Err.Clear
        Error code
        result = result & CStr(code) & ":" & CStr(Err.HelpContext) & ";"
    Next code
    ErrorContextCatalog = result
End Function

Public Function RaiseContextDefaultSweep() As String
    On Error Resume Next
    Dim code As Long, generic As Long, builtin As Long, other As Long
    For code = 1 To 65535
        Err.Clear
        Err.Raise code
        If Err.HelpContext = 1000095 Then
            generic = generic + 1
        ElseIf Err.HelpContext = 1000000 + code Then
            builtin = builtin + 1
        Else
            other = other + 1
        End If
    Next code
    RaiseContextDefaultSweep = CStr(generic) & "|" & CStr(builtin) & "|" & CStr(other)
End Function

Public Function RaiseContextCustomSweep() As String
    On Error Resume Next
    Dim code As Long, generic As Long, builtin As Long, other As Long
    For code = 1 To 65535
        Err.Clear
        Err.Raise code, , "custom"
        If Err.HelpContext = 1000095 Then
            generic = generic + 1
        ElseIf Err.HelpContext = 1000000 + code Then
            builtin = builtin + 1
        Else
            other = other + 1
        End If
    Next code
    RaiseContextCustomSweep = CStr(generic) & "|" & CStr(builtin) & "|" & CStr(other)
End Function

Public Function RaiseContextEmptySweep() As String
    On Error Resume Next
    Dim code As Long, generic As Long, builtin As Long, other As Long
    For code = 1 To 65535
        Err.Clear
        Err.Raise code, , ""
        If Err.HelpContext = 1000095 Then
            generic = generic + 1
        ElseIf Err.HelpContext = 1000000 + code Then
            builtin = builtin + 1
        Else
            other = other + 1
        End If
    Next code
    RaiseContextEmptySweep = CStr(generic) & "|" & CStr(builtin) & "|" & CStr(other)
End Function
