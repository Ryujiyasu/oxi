// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What a macro would do if it ran, read before running it.
//!
//! This answers one question: *is it safe to enable the macros in this file?*
//! It is read from the source text alone — the file is never opened in Excel,
//! nothing is executed, and no decision is made here. What comes back is
//! **evidence with line numbers**, so a person decides.
//!
//! # Why it gives no verdict
//!
//! `Shell` is how a legitimate macro opens a PDF. `MSXML2.XMLHTTP` is how one
//! fetches a rate table. Neither is a finding on its own, and a tool that
//! shouts "dangerous" at both teaches its reader to click through. So this
//! reports *capabilities* — what the code can reach — and leaves the judgement
//! where the context is.
//!
//! The one thing it states plainly is whether the code runs **without anyone
//! pressing anything**, because that changes the question from "should I run
//! this?" to "have I already run it?".
//!
//! # What it cannot see, and says so
//!
//! Late binding hides its target: `CreateObject(name)` where `name` is
//! computed is reported as unresolved rather than guessed at. Lines the parser
//! could not read are carried through as a count, because a line nobody could
//! read is a line nobody has cleared. Neither is silently dropped.

use std::collections::BTreeMap;

use crate::analysis::Analysis;
use crate::ast::{Module, ModuleItem};

/// The kind of reach a signal is evidence of.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum Capability {
    /// Runs because the file was opened or closed, with nobody pressing
    /// anything. First because it changes what the reader is deciding.
    RunsOnOpen,
    /// Writes or runs VBA at run time, so what it does is not in the file
    /// being read.
    WritesItsOwnCode,
    /// Starts another program, or asks the shell to.
    StartsAProgram,
    /// Reaches the network.
    ReachesTheNetwork,
    /// Calls into a DLL directly.
    CallsIntoADll,
    /// Reads or writes the Windows registry.
    TouchesTheRegistry,
    /// Writes, moves or deletes files.
    TouchesFiles,
    /// Creates something outside Excel whose own behaviour is not read here.
    ReachesOutsideExcel,
    /// Builds what it reaches at run time rather than writing it down.
    HidesWhatItReaches,
    /// Turns off the prompts and screen updates that would show its work.
    QuietensExcel,
}

impl Capability {
    pub fn label(self) -> &'static str {
        match self {
            Capability::RunsOnOpen => "runs when the file is opened",
            Capability::WritesItsOwnCode => "writes or runs code at run time",
            Capability::StartsAProgram => "starts another program",
            Capability::ReachesTheNetwork => "reaches the network",
            Capability::CallsIntoADll => "calls into a DLL",
            Capability::TouchesTheRegistry => "reads or writes the registry",
            Capability::TouchesFiles => "reads or writes files",
            Capability::ReachesOutsideExcel => "creates an object outside Excel",
            Capability::HidesWhatItReaches => "builds what it reaches at run time",
            Capability::QuietensExcel => "turns off Excel's prompts or display",
        }
    }
}

/// One piece of evidence: what was found, where, and why it is worth seeing.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Signal {
    pub capability: Capability,
    /// The name or text as written, so a reader recognises their own code.
    pub what: String,
    pub reason: String,
    /// `0` when the evidence is a name counted across the module rather than
    /// one statement.
    pub line: u32,
}

/// What a module or project would be able to do.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SafetyReport {
    pub signals: Vec<Signal>,
    /// Procedures Office runs on its own, by name.
    pub runs_without_asking: Vec<String>,
    /// Lines holding a `CreateObject`/`GetObject` whose target is computed.
    pub unresolved_late_binding: Vec<u32>,
    /// Lines the parser could not read, carried through from the analyser.
    pub unread_lines: usize,
}

impl SafetyReport {
    /// Whether opening the file is enough to run code.
    pub fn runs_on_open(&self) -> bool {
        !self.runs_without_asking.is_empty()
    }

    /// The capabilities present, each with how much evidence there is.
    pub fn capabilities(&self) -> BTreeMap<Capability, usize> {
        let mut counts = BTreeMap::new();
        for signal in &self.signals {
            *counts.entry(signal.capability).or_default() += 1;
        }
        counts
    }

    /// Whether a person still has to read something. False means: no signal,
    /// nothing computed at run time, and nothing the parser could not read —
    /// the only state in which silence is honest.
    pub fn needs_a_reader(&self) -> bool {
        !self.signals.is_empty()
            || !self.unresolved_late_binding.is_empty()
            || self.unread_lines > 0
    }

    fn absorb(&mut self, other: SafetyReport) {
        self.signals.extend(other.signals);
        self.runs_without_asking.extend(other.runs_without_asking);
        self.unresolved_late_binding
            .extend(other.unresolved_late_binding);
        self.unread_lines += other.unread_lines;
    }

    fn settle(&mut self) {
        self.signals.sort_by(|a, b| {
            a.capability
                .cmp(&b.capability)
                .then(a.what.cmp(&b.what))
                .then(a.line.cmp(&b.line))
        });
        self.signals.dedup();
        self.runs_without_asking.sort();
        self.runs_without_asking.dedup();
        self.unresolved_late_binding.sort_unstable();
    }
}

/// Reads one module. `analysis` is the same module's [`Analysis`], passed in
/// rather than recomputed so the two reports cannot disagree about what the
/// file contains.
pub fn assess(module: &Module, analysis: &Analysis) -> SafetyReport {
    let mut report = SafetyReport {
        unread_lines: analysis.metrics.unparsed,
        ..SafetyReport::default()
    };

    for procedure in &analysis.procedures {
        if let Some(reason) = runs_by_itself(&procedure.name) {
            report.runs_without_asking.push(procedure.name.clone());
            report.signals.push(Signal {
                capability: Capability::RunsOnOpen,
                what: procedure.name.clone(),
                reason: reason.to_string(),
                line: procedure.line,
            });
        }
    }

    for item in &module.items {
        if let ModuleItem::ExternalProc(declared) = item {
            report.signals.push(Signal {
                capability: Capability::CallsIntoADll,
                what: format!("Declare {} Lib \"{}\"", declared.name, declared.lib),
                reason: format!(
                    "calls {} in {} directly; what it does is not VBA and is not read here",
                    declared.alias.as_deref().unwrap_or(&declared.name),
                    declared.lib
                ),
                line: declared.span.line,
            });
        }
    }

    // Dotted names keep their receiver, so `shell.Run` and a bare `Run` are
    // told apart. These are counted per module, so they carry no line.
    for name in analysis.api_names.keys() {
        if let Some((capability, reason)) = reached_by_name(name) {
            report.signals.push(Signal {
                capability,
                what: name.clone(),
                reason: reason.to_string(),
                line: 0,
            });
        }
    }

    for binding in &analysis.late_bindings {
        match &binding.target {
            Some(target) => {
                let (capability, reason) = reached_by_progid(target);
                report.signals.push(Signal {
                    capability,
                    what: format!("{}(\"{target}\")", binding.callee),
                    reason: reason.to_string(),
                    line: binding.line,
                });
            }
            None => {
                report.unresolved_late_binding.push(binding.line);
                report.signals.push(Signal {
                    capability: Capability::HidesWhatItReaches,
                    what: format!("{}(...)", binding.callee),
                    reason: "the object it creates is worked out while it runs, so reading the \
                             source cannot say what this reaches"
                        .to_string(),
                    line: binding.line,
                });
            }
        }
    }

    report.settle();
    report
}

/// Reads a whole project: one report over every module in it.
pub fn assess_project<'a>(
    modules: impl IntoIterator<Item = (&'a Module, &'a Analysis)>,
) -> SafetyReport {
    let mut whole = SafetyReport::default();
    for (module, analysis) in modules {
        whole.absorb(assess(module, analysis));
    }
    whole.settle();
    whole
}

/// The procedure names Office calls on its own. Opening the file is the
/// trigger, so for these the question is not whether to run the macro.
///
/// Matched on the leaf so a class module's `Workbook_Open` is caught however
/// the module qualifies it.
fn runs_by_itself(name: &str) -> Option<&'static str> {
    let lowered = name.to_ascii_lowercase();
    let plain = lowered.rsplit('.').next().unwrap_or(&lowered);
    Some(match plain {
        "auto_open" | "autoopen" | "auto_exec" | "autoexec" | "workbook_open"
        | "document_open" => "Office runs this as the file opens",
        "auto_close" | "autoclose" | "workbook_beforeclose" | "document_close" => {
            "Office runs this as the file closes"
        }
        "auto_new" | "autonew" | "document_new" | "workbook_activate"
        | "workbook_windowactivate" => "Office runs this without anyone asking",
        "workbook_beforesave" | "workbook_aftersave" => "Office runs this when the file is saved",
        "workbook_sheetchange" | "worksheet_change" | "worksheet_selectionchange"
        | "worksheet_activate" => "Office runs this when the sheet is touched",
        _ => return None,
    })
}

/// What a written-down ProgID reaches. Applied ONLY to the string inside a
/// `CreateObject` / `GetObject`, never to an identifier: a substring test is
/// right for a ProgID and wrong for code, where `web_WinHttpRequestOption` is
/// an enum member and reaches nothing.
fn reached_by_progid(progid: &str) -> (Capability, &'static str) {
    let lowered = progid.to_ascii_lowercase();
    if lowered.starts_with("msxml2.")
        || lowered.starts_with("winhttp.")
        || lowered.contains("xmlhttp")
    {
        return (
            Capability::ReachesTheNetwork,
            "fetches from a URL while it runs",
        );
    }
    if lowered.starts_with("wscript.shell") || lowered.starts_with("shell.application") {
        return (
            Capability::StartsAProgram,
            "creates the Windows shell object, which runs command lines",
        );
    }
    if lowered.contains("filesystemobject") {
        return (
            Capability::TouchesFiles,
            "opens the file system through the Scripting runtime",
        );
    }
    if lowered.starts_with("adodb.stream") {
        return (
            Capability::TouchesFiles,
            "writes raw bytes to disk through ADO, which is how fetched content is saved",
        );
    }
    if lowered.starts_with("scripting.") {
        return (
            Capability::ReachesOutsideExcel,
            "creates a Windows Scripting object",
        );
    }
    (
        Capability::ReachesOutsideExcel,
        "creates an object outside Excel; what that object does is not read here",
    )
}

/// What an identifier in the code reaches. Matched on whole dotted segments so
/// a name that merely CONTAINS a word is left alone.
fn reached_by_name(name: &str) -> Option<(Capability, &'static str)> {
    let lowered = name.to_ascii_lowercase();
    let segments: Vec<&str> = lowered.split('.').collect();
    let leaf = *segments.last().unwrap_or(&lowered.as_str());
    let dotted = segments.len() > 1;

    if segments
        .iter()
        .any(|segment| matches!(*segment, "vbproject" | "vbcomponents"))
    {
        return Some((
            Capability::WritesItsOwnCode,
            "reaches the VBA project itself, so it can add or change macros while it runs",
        ));
    }

    match leaf {
        "shell" if !dotted => Some((
            Capability::StartsAProgram,
            "VBA's own Shell runs a command line",
        )),
        "exec" if dotted => Some((Capability::StartsAProgram, "runs a command line")),
        "regwrite" | "regread" | "regdelete" => Some((
            Capability::TouchesTheRegistry,
            "reads or writes Windows settings outside this file",
        )),
        "kill" if !dotted => Some((Capability::TouchesFiles, "deletes a file")),
        "savecopyas" | "saveas" => Some((Capability::TouchesFiles, "writes a file to disk")),
        "displayalerts" => Some((
            Capability::QuietensExcel,
            "turns Excel's confirmation prompts off, so what follows happens without asking",
        )),
        "screenupdating" => Some((
            Capability::QuietensExcel,
            "stops the screen redrawing, so what follows is not shown",
        )),
        "environ" if !dotted => Some((
            Capability::HidesWhatItReaches,
            "reads an environment variable, commonly to build a path at run time",
        )),
        _ => None,
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    use crate::{analyse, parse_module};


    fn read(source: &str) -> SafetyReport {
        let module = parse_module(source).expect("should parse");
        let analysis = analyse(&module);
        assess(&module, &analysis)
    }

    #[test]
    fn a_macro_that_only_adds_up_says_nothing() {
        let report = read(
            "Option Explicit\n\
             Public Function Total(ByVal a As Long, ByVal b As Long) As Long\n\
               Total = a + b\n\
             End Function\n",
        );
        assert!(!report.needs_a_reader());
        assert!(!report.runs_on_open());
        assert!(report.signals.is_empty());
    }

    #[test]
    fn opening_the_file_is_the_trigger() {
        let report = read(
            "Private Sub Workbook_Open()\n\
               MsgBox \"hello\"\n\
             End Sub\n",
        );
        assert!(report.runs_on_open());
        assert_eq!(report.runs_without_asking, ["Workbook_Open"]);
        assert_eq!(report.signals[0].capability, Capability::RunsOnOpen);
        assert_eq!(report.signals[0].line, 1);
    }

    #[test]
    fn a_written_down_progid_is_named() {
        let report = read(
            "Sub Fetch()\n\
               Dim http As Object\n\
               Set http = CreateObject(\"MSXML2.XMLHTTP\")\n\
               Dim fso As Object\n\
               Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
             End Sub\n",
        );
        let capabilities = report.capabilities();
        assert_eq!(capabilities.get(&Capability::ReachesTheNetwork), Some(&1));
        assert_eq!(capabilities.get(&Capability::TouchesFiles), Some(&1));
        assert!(report.unresolved_late_binding.is_empty());
        assert!(report
            .signals
            .iter()
            .any(|s| s.what.contains("MSXML2.XMLHTTP") && s.line == 3));
    }

    /// The honest answer to a computed ProgID is "this cannot be read", not a
    /// guess and not silence.
    #[test]
    fn a_computed_progid_is_reported_as_unreadable() {
        let report = read(
            "Sub Go(ByVal which As String)\n\
               Dim thing As Object\n\
               Set thing = CreateObject(which)\n\
             End Sub\n",
        );
        assert_eq!(report.unresolved_late_binding, [3]);
        assert_eq!(
            report.signals[0].capability,
            Capability::HidesWhatItReaches
        );
        assert!(report.needs_a_reader());
    }

    #[test]
    fn a_declare_names_its_library() {
        let report = read(
            "Private Declare PtrSafe Function CreateThread Lib \"kernel32\" \
             (ByVal a As LongPtr) As LongPtr\n\
             Sub Go()\n\
             End Sub\n",
        );
        assert_eq!(report.signals[0].capability, Capability::CallsIntoADll);
        assert!(report.signals[0].what.contains("kernel32"));
    }

    #[test]
    fn writing_its_own_macros_is_its_own_capability() {
        let report = read(
            "Sub Grow()\n\
               ThisWorkbook.VBProject.VBComponents.Add 1\n\
             End Sub\n",
        );
        assert_eq!(
            report.capabilities().keys().next(),
            Some(&Capability::WritesItsOwnCode)
        );
    }

    /// A substring test is right for a ProgID and wrong for code. `VBA-Web`
    /// declares an enum whose members are named `web_WinHttpRequestOption_*`;
    /// they reach nothing, and calling them a network signal would be the kind
    /// of false alarm that gets the whole report ignored.
    #[test]
    fn a_name_that_merely_contains_a_word_reaches_nothing() {
        let report = read(
            "Public Enum web_WinHttpRequestOption\n\
               web_WinHttpRequestOption_EnableRedirects = 6\n\
             End Enum\n\
             Sub Go()\n\
               Dim n As Long\n\
               n = web_WinHttpRequestOption_EnableRedirects\n\
             End Sub\n",
        );
        assert_eq!(report.capabilities().get(&Capability::ReachesTheNetwork), None);
    }

    /// A ProgID that IS written down is known, whatever it is. Calling it
    /// "built at run time" would contradict the source in front of the reader.
    #[test]
    fn a_progid_written_down_is_never_called_hidden() {
        let report = read(
            "Sub Go()\n\
               Dim ie As Object\n\
               Set ie = CreateObject(\"InternetExplorer.Application\")\n\
             End Sub\n",
        );
        assert_eq!(
            report.signals[0].capability,
            Capability::ReachesOutsideExcel
        );
        assert_eq!(report.capabilities().get(&Capability::HidesWhatItReaches), None);
    }

    #[test]
    fn a_shell_call_is_told_apart_from_a_worksheet_named_shell() {
        let ran = read("Sub Go()\n  Shell \"cmd.exe /c dir\"\nEnd Sub\n");
        assert_eq!(
            ran.capabilities().get(&Capability::StartsAProgram),
            Some(&1)
        );
        let named = read("Sub Go()\n  Sheets(\"Shell\").Range(\"A1\").Value = 1\nEnd Sub\n");
        assert_eq!(named.capabilities().get(&Capability::StartsAProgram), None);
    }

    /// A line nobody could read is a line nobody has cleared, so it keeps the
    /// report from going quiet even when no signal fired.
    #[test]
    fn an_unread_line_keeps_the_report_awake() {
        let module = parse_module("Sub Go()\n  Next For 3 To\nEnd Sub\n").expect("should parse");
        let analysis = analyse(&module);
        let report = assess(&module, &analysis);
        assert!(report.unread_lines > 0);
        assert!(report.needs_a_reader());
    }

    #[test]
    fn a_project_is_read_as_one() {
        let opener = parse_module("Private Sub Workbook_Open()\n  Go\nEnd Sub\n").unwrap();
        let worker =
            parse_module("Sub Go()\n  Shell \"cmd.exe\"\nEnd Sub\n").unwrap();
        let (a, b) = (analyse(&opener), analyse(&worker));
        let whole = assess_project([(&opener, &a), (&worker, &b)]);
        assert!(whole.runs_on_open());
        assert_eq!(
            whole.capabilities().get(&Capability::StartsAProgram),
            Some(&1)
        );
    }
}
