// SPDX-License-Identifier: CC-BY-SA-4.0
// Self-authored components, each measured twice in opposite orders in Excel.
// Excel 16.0 build 20527.0, Japanese locale 1041. CRLF normalized to LF.
// These expectations cover the component compiled on entry; they do not assert
// that the unrelated chart project prelude compiles without Excel references.
struct Case { id: &'static str, program: &'static str, native_kind: Option<&'static str>, description: Option<&'static str> }
const CASES: &[Case] = &[
    // Native full-project source SHA256: 89429f6dcc03d1a77a259c9fa947c09bab450de579ec4d929d96dd545f870899
    Case { id: "for_before_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
For counter = 1 To 1
Next counter
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("VariableNotDefined"), description: Some("変数が定義されていません。") },
    // Native full-project source SHA256: 6a04b01f436819b042d491ddd4cbb7d7f50abcd99850801baaacdbf620dd107c
    Case { id: "for_after_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long
For counter = 1 To 1
Next counter
Probe = counter
End Function

"########, native_kind: None, description: None },
    // Native full-project source SHA256: d2a8f3f686423a6dabedb6c663904e499b744a06ea255665ee8e93a28e893dd5
    Case { id: "assignment_before_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
counter = 7
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("VariableNotDefined"), description: Some("変数が定義されていません。") },
    // Native full-project source SHA256: c9fa57606998b52a7eb1451b46c70f2b502ed7efe81ea957c2c1c0d6c27fc571
    Case { id: "read_before_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Probe = TypeName(counter)
Dim counter As Long
End Function

"########, native_kind: Some("VariableNotDefined"), description: Some("変数が定義されていません。") },
    // Native full-project source SHA256: 277da8efc4bcdb6f85e951d589d90d57aa17ffe0f217792111fe26d40eeaafdf
    Case { id: "for_before_local_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
For counter = 1 To 1
Next counter
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 634b9ea40ac0089787ec227f694eb5090e8bed60b154f17fbb2635e631a5c214
    Case { id: "assignment_before_local_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
counter = 7
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: b31645e20d8a6b642909377266cbbe6186dfb6be49d1a1bb7cd1523c8542cc2c
    Case { id: "for_before_local_module_shadow", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Public Function Probe() As Variant
On Error Resume Next
For counter = 1 To 1
Next counter
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 54e5d10f8d75c894dd492cee2441ecbe911f5074c6170d18c5969bf7cda62a9c
    Case { id: "for_each_before_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
For Each counter In Array(1)
Next counter
Dim counter As Variant
Probe = TypeName(counter)
End Function

"########, native_kind: Some("VariableNotDefined"), description: Some("変数が定義されていません。") },
    // Native full-project source SHA256: c18205c2c930bc2631188c4bfd610845e783f5d435d42e9af89f1c92c85ad71b
    Case { id: "for_each_after_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Variant
For Each counter In Array(1)
Next counter
Probe = TypeName(counter)
End Function

"########, native_kind: None, description: None },
    // Native full-project source SHA256: e28d366ae5b57cc97b6bd61d6be7b1973c87bfba8ee3b3d1149e4acabd993666
    Case { id: "callbyname_before_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Probe = CallByName(target, "Count", VbGet)
Dim target As Object
End Function

"########, native_kind: Some("VariableNotDefined"), description: Some("変数が定義されていません。") },
    // Native full-project source SHA256: 4fd822ee02030c3558a23025fcb71f2148f0dd870aa397323cdf3e48b6dbb074
    Case { id: "read_before_local_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Probe = TypeName(counter)
Dim counter As Long
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: c6f90c531af318b885bb5b8ed2a5573916d28582e46b6ab1b3474a29b94f11a1
    Case { id: "read_before_local_module_shadow_explicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Private Sub InitializeCounter()
counter = 40
End Sub
Public Function Probe() As Variant
On Error Resume Next
InitializeCounter
Probe = counter
Dim counter As Long
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: dab15f74f6ba2318b2c66e96bc97900ba409b0297f49c1850080b296c0b3d493
    Case { id: "read_before_local_module_shadow_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Private Sub InitializeCounter()
counter = 40
End Sub
Public Function Probe() As Variant
On Error Resume Next
InitializeCounter
Probe = counter
Dim counter As Long
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 723fc7edcf1569d5440be614465f90aff7a5a3444c26bbf938c26335a4b96d57
    Case { id: "declared_before_module_read", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Private Sub InitializeCounter()
counter = 40
End Sub
Public Function Probe() As Variant
On Error Resume Next
InitializeCounter
Dim counter As Long
Probe = counter
End Function

"########, native_kind: None, description: None },
    // Native full-project source SHA256: 7d0d1fdf9b8e5ff8739a661be612cac4da4b07d61a978e80cca108191f47bd9f
    Case { id: "duplicate_local_declaration", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long
Dim counter As Long
Probe = counter
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 2f8192a363beb3119e1244d2fe59f6fb9c660a843647f1c571e4c11e33d8f047
    Case { id: "module_for_without_local", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Public Function Probe() As Variant
On Error Resume Next
For counter = 1 To 1
Next counter
Probe = counter
End Function

"########, native_kind: None, description: None },
    // Native full-project source SHA256: 6f468ffea9eabe41632a8bbf2bf08089dd5c786375228846f586e341d66c8ffc
    Case { id: "typed_read_before_local_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Probe = TypeName(counter&)
Dim counter As Long
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 0de16d3d5feeea0bd271d7d2e05b92acd572d39c72655f9a321fb2237e8c4f31
    Case { id: "exact_existing_runtime_scope_test", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Private value As Long
Private Sub SetModuleValue()
value = 40
End Sub
Private Function LateDeclaration() As Long
LateDeclaration = value
Dim value As Long
End Function
Private Function NestedDeclaration() As Long
If False Then
Dim hidden As Long
End If
NestedDeclaration = hidden
End Function
Public Function ProcedureScopeProbe() As Long
SetModuleValue
ProcedureScopeProbe = LateDeclaration() + NestedDeclaration()
End Function

Public Function Probe() As Variant
On Error Resume Next
Probe = ProcedureScopeProbe()
End Function

"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 587fec64fc8bf7d2082fb64c2e45a819ef4342d1a366268bab84ac4a90ee597a
    Case { id: "duplicate_dim_same_statement", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long, counter As Long
Probe = counter
End Function
"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 5534449e2a4bd1f37f87d650a0ccb92453007e3d5580666cc7cd497df7331e82
    Case { id: "duplicate_dim_same_statement_implicit", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long, counter As Long
Probe = counter
End Function
"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 0bf1b0fa0546f0519e409623a0f20dddb28ff5c1b26c5585ed25367a366bfc28
    Case { id: "duplicate_dim_case_insensitive", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long, COUNTER As Long
Probe = counter
End Function
"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: a6f31bd5a9474f78ffac62065825eacf044e5070ca1b59f0e66551102d1a4f45
    Case { id: "distinct_dim_same_statement", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long, other As Long
counter = 3
other = 4
Probe = counter + other
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: a42b2231ff01144a68ed8aa71cbbb77b7b12a4b2f6e12d9dc82dc8661ca72df4
    Case { id: "duplicate_const_same_statement", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Const counter As Long = 1, counter As Long = 2
Probe = counter
End Function
"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: fee74eb01814592db45ac7b1c00c77360d3189c3dcdaa17dd6a41b50841cbe20
    Case { id: "distinct_const_same_statement", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Const counter As Long = 1, other As Long = 2
Probe = counter + other
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 0bf3105847293b96583211cee4585ac7de227c9b3bf4c1e8542b3b86e87239be
    Case { id: "duplicate_dim_colon", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long: Dim counter As Long
Probe = counter
End Function
"########, native_kind: Some("DuplicateDeclaration"), description: Some("同じ適用範囲内で宣言が重複しています。") },
    // Native full-project source SHA256: 9a56731023cd84d16222fad9bd168ee7d76010e4793df2bc8cdd1a992a440d16
    Case { id: "nested_unexecuted_declaration_before_read", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
If False Then
Dim counter As Long
End If
Probe = counter
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 5a6c5df20047a3ac706fa607d30770d55166b394f42feb47c64507a0a9219572
    Case { id: "module_const_distinct_chain", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private Const first As Long = 1, second As Long = first + 2
Public Function Probe() As Variant
On Error Resume Next
Probe = second
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 45ab30986d06ee511cced4f369215740f88d842a03719f58a02b2256fb599c64
    Case { id: "local_const_distinct_chain", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Public Function Probe() As Variant
On Error Resume Next
Const first As Long = 1, second As Long = first + 2
Probe = second
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 297cfb990b93acb302749507d713a0c6a67d7448dd8789b759b2117c602615e0
    Case { id: "module_distinct_dim_control", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long, other As Long
Public Function Probe() As Variant
On Error Resume Next
Probe = counter + other
End Function
"########, native_kind: None, description: None },
];

const MODULE_COLLISION_CASES: &[Case] = &[
    // Native full-project source SHA256: 13392f82d5debdf5e8dc78278e9e5b52d94040d03bca3d058d2a11b5bd1fa3e1
    Case { id: "dim_same_application_run", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long, counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = Application.Run("DeclarationOrderTarget.GetValue")
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 5f29999fde465759cc764c225cd409db2dddb53d1d4227f7ebb5f11790be4e2a
    Case { id: "dim_same_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long, counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: f8d99c8ed0cde095f8d3d9657bba5c0eeac7c63c47e2e204a22642f2d647f4bd
    Case { id: "dim_same_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long, counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 5064dcf0a34cee2a71c346f53ffc4551207b0fb3146945ddd7cae5082a3ce006
    Case { id: "dim_separate_application_run", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = Application.Run("DeclarationOrderTarget.GetValue")
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: b9465e9dc461f9dd8dbf90b89b7439f5148ba8cd0dfefdd85aef8ed11a8f516a
    Case { id: "dim_separate_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: a250dcec38ad05c99a9df71ca02104758cabc0f55a25b43320b2645524130d7f
    Case { id: "dim_separate_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 57825c9312f94de593801775c8466ac39c974784a7e3f4c4f04f3ab59cfde6fb
    Case { id: "const_same_application_run", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1, counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = Application.Run("DeclarationOrderTarget.GetValue")
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: e7adc228802a5b68b4fbc8b536468095dc79ef2bcd034952f6c2b237c81b32c9
    Case { id: "const_same_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1, counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 2541aa401838dcf2072a7a0b36e68bf06a6a4c71cb76ffe42a029837e80d8fb5
    Case { id: "const_same_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1, counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 81459bd82642d99df260815e37f08028b28fc5c5c9886f573f4033529ad4e061
    Case { id: "const_separate_application_run", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = Application.Run("DeclarationOrderTarget.GetValue")
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 441a87b67d266d4c54e20c468b5257ad1e7a58784208b4ffd92d16d38abeea5b
    Case { id: "const_separate_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 6db218a80d4631512a190831266ff64638118166e2a73a40b7e2aca17d111d38
    Case { id: "const_separate_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: c810e7932bc069266877f0afc5412d7e02d2074a134629ddf97e0380c2586a93
    Case { id: "var_const_application_run", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = Application.Run("DeclarationOrderTarget.GetValue")
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 167f65a8b4090c430e48e39e3f6e0731a26d02b1b4a774764b5af77b48585d8f
    Case { id: "var_const_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: 6676ed4ad9252741f81318fd955951f730daf79063589143e3cc498925a317d9
    Case { id: "var_const_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private Const counter As Long = 2
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: counter"########) },
    // Native full-project source SHA256: b7a4e3314ff0659057b76b053fb2cb7b4a7eb633764e79bd0c23fedf3994f57a
    Case { id: "valid_direct_call", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 84c1aba6d9b51ce2866cbd4dcad3983faa498fbde2c72aeedad6e32a4a315d64
    Case { id: "valid_unused", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Public Function GetValue() As Variant
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 45b1309da778c646cdd758fa3dce69c91c7de917d706d30e87bda97f18cc8683
    Case { id: "private_dim_separate_modules", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Public Function GetValue() As Long
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOther"
Option Explicit
Private counter As Long
Public Function GetValue() As Long
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue + DeclarationOrderOther.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: bbac8cc0abac4de393adb61f5f3a694fd8c00f4752cedd618ca0df6004ae0b43
    Case { id: "private_const_separate_modules", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private Const counter As Long = 1
Public Function GetValue() As Long
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOther"
Option Explicit
Private Const counter As Long = 2
Public Function GetValue() As Long
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = DeclarationOrderTarget.GetValue + DeclarationOrderOther.GetValue
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 0b2268dc4c021f2f7cb67cd44fea60cf1537909022d890f34e8712648b0b45c7
    Case { id: "local_shadows_module", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Private counter As Long
Public Function Probe() As Variant
On Error Resume Next
Dim counter As Long
counter = 7
Probe = Array(Err.Number, Err.Source, TypeName(counter), counter)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: b7b13f9deb854485e48b84bb7fed8ca4158f71c83e28597ef95cd8e3e03fd515
    Case { id: "separate_record_fields", program: r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Private Type FirstRecord
counter As Long
End Type
Private Type SecondRecord
counter As Long
End Type
Public Function Probe() As Variant
On Error Resume Next
Dim result As Long
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: None, description: None },
    // Native full-project source SHA256: 9b57743e0497b9c465bc96f42aff13958f78f3b6c87a9db1cfd5a8907c169707
    Case { id: "casefold_duplicate_module", program: r########"Attribute VB_Name = "DeclarationOrderTarget"
Option Explicit
Private counter As Long
Private COUNTER As Long
Public Function GetValue() As Long
GetValue = counter
End Function

Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
Public Function Probe() As Variant
On Error Resume Next
Dim result As Variant
Err.Clear
result = 7
Probe = Array(Err.Number, Err.Source, TypeName(result), result)
End Function
"########, native_kind: Some("AmbiguousName"), description: Some(r########"名前が適切ではありません: COUNTER"########) },
];

const MODULE_UNAVAILABLE_CASES: &[(&str, &str)] = &[
    // Native full-project source SHA256: 3fd148a25c1d0ad85adbce518e04783cd4b8fafca5d8981d43db289d21719168
    ("module_dim_duplicate_same_used", r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long, counter As Long
Public Function Probe() As Variant
On Error Resume Next
Probe = counter
End Function
"########),
    // Native full-project source SHA256: dfd04a199803283e5407180063497cf0f8258dc863b0648cd23c4a91d8b5b7dd
    ("module_dim_duplicate_separate_used", r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Private counter As Long
Public Function Probe() As Variant
On Error Resume Next
Probe = counter
End Function
"########),
    // Native full-project source SHA256: f8aa6902d8b15b4c73a0782f8a7bc73427ff2ac7596bbb13dcddcd430a519cd9
    ("module_const_duplicate_same_used", r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private Const counter As Long = 1, counter As Long = 2
Public Function Probe() As Variant
On Error Resume Next
Probe = counter
End Function
"########),
    // Native full-project source SHA256: 72c0d2315e6a32de0bda7503ffd8d5d7c0f583f11e0cf2c1834c452bc19d0ecf
    ("module_const_duplicate_separate_used", r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private Const counter As Long = 1
Private Const counter As Long = 2
Public Function Probe() As Variant
On Error Resume Next
Probe = counter
End Function
"########),
    // Native full-project source SHA256: fab6cf679b76d086234588452d8a70f9c944c224125758959467dc1765d1e1a4
    ("module_var_and_const_duplicate_used", r########"Attribute VB_Name = "DeclarationOrderOracle"
Option Explicit
' Self-authored conformance probe, CC BY-SA 4.0.
Private counter As Long
Private Const counter As Long = 2
Public Function Probe() As Variant
On Error Resume Next
Probe = counter
End Function
"########),
];
