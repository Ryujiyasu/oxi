// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::parser::parse_project;
use oxivba_core::runtime::{Runtime, Value};

// Native showed an unhandled error-5 dialog with Continue disabled. The
// physical trace stopped at handler_before despite caller/emitter catches.
// End returned a generic COM error, not a complete original Err tuple.
#[test]
fn unhandled_event_error_bypasses_raising_procedure_catches() {
    let (module, classes) = parse_project(SOURCE).unwrap();
    let mut runtime = Runtime::new(&module).with_classes(&classes)
        .with_project_name("VBAProject");
    runtime.file_system_mut().mount_folder(FOLDER).unwrap();
    let result = runtime.call("Probe", vec![Value::Integer(0), Value::String(TARGET.into())]);
    let trace = runtime.file_system_mut().bytes(TARGET).unwrap();
    assert_eq!(trace, b"caller_before\r\nemitter_before\r\nhandler_before\r\n",
        "Native stops without entering emitter/caller handlers; result={result:?}");
    let failure = result.expect_err("Native exposes an unhandled event error");
    assert_eq!(failure.vba_number, Some(5));
    assert_eq!(failure.message, "explicit-event-error-4750");
}

const TARGET: &str = r####"E:\oxi-vba-public-artifacts-20261115\types4750\owned-data-4751-0-0.bin"####;
const FOLDER: &str = r####"E:\oxi-vba-public-artifacts-20261115\types4750"####;
const SOURCE: &str = r####"Attribute VB_Name = "FatalCases4750"
Option Explicit
Private target As String
Private source As Emitter4750
Private sink As Sink4750
Public Sub Mark(ByVal text As String)
 Dim handle As Integer
 handle=FreeFile
 Open target For Append As #handle
 Print #handle,text
 Close #handle
End Sub
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 target=pathname
 On Error GoTo CallerCaught
 Set source=New Emitter4750
 Set sink=New Sink4750
 sink.Attach source
 Mark "caller_before"
 source.Emit
 Mark "caller_after"
 GoTo Cleanup
CallerCaught:
 Mark "caller_caught:" & Err.Number
Cleanup:
 sink.Detach
 Mark "cleanup"
 Set sink=Nothing
 Set source=Nothing
 Probe=Array("returned",Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext))
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Emitter4750"
Option Explicit
Public Event Ping()
Public Sub Emit()
 On Error GoTo EmitterCaught
 FatalCases4750.Mark "emitter_before"
 RaiseEvent Ping
 FatalCases4750.Mark "emitter_after"
 Exit Sub
EmitterCaught:
 FatalCases4750.Mark "emitter_caught:" & Err.Number
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Sink4750"
Option Explicit
Private WithEvents origin As Emitter4750
Public Sub Attach(ByVal value As Emitter4750)
 Set origin=value
End Sub
Public Sub Detach()
 Set origin=Nothing
End Sub
Private Sub origin_Ping()
 FatalCases4750.Mark "handler_before"
 Err.Raise 5,"FatalSink4750","explicit-event-error-4750"
 FatalCases4750.Mark "handler_after"
End Sub
"####;
