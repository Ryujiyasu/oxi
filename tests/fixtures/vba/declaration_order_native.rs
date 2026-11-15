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
