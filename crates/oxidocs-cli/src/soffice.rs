// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `soffice --headless --convert-to pdf --outdir DIR FILE…`, answered by Oxi.
//!
//! Many programs convert Office files by running LibreOffice with exactly
//! these arguments, from a path the administrator can change: file servers
//! making thumbnails, report modules, the `libreoffice-convert` package. Run
//! as `soffice` (a link to this binary named so) or as `oxidocs soffice …`,
//! Oxi takes that command line as LibreOffice would for .docx, .pptx and
//! .xlsx to PDF or PNG, and writes `DIR/<name>.pdf` (or `.png`, the first
//! page) where those programs look for it.
//!
//! Everything else — another file type, another target, an option Oxi does
//! not have — goes to the real LibreOffice named by `OXIDOCS_SOFFICE_FALLBACK`,
//! with the same arguments; without one it is an error, never a silent skip.

use std::path::{Path, PathBuf};

/// The variable naming the real `soffice` that what Oxi does not take goes to.
pub(crate) const FALLBACK_VAR: &str = "OXIDOCS_SOFFICE_FALLBACK";

/// Names under which this binary answers as LibreOffice.
const NAMES: &[&str] = &["soffice", "soffice.bin", "libreoffice", "soffice.exe", "libreoffice.exe"];

/// The arguments after the program when this run is a LibreOffice command
/// line: the binary was started as `soffice` (or the like), or as
/// `oxidocs soffice …`. `None` for every other command.
pub(crate) fn invocation(args: &[String]) -> Option<&[String]> {
    let program = args.first().map(|p| {
        Path::new(p).file_name().and_then(|n| n.to_str()).unwrap_or("").to_ascii_lowercase()
    })?;
    if NAMES.contains(&program.as_str()) {
        return Some(&args[1..]);
    }
    if args.get(1).map(String::as_str) == Some("soffice") {
        return Some(&args[2..]);
    }
    None
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum Target {
    Pdf,
    Png,
}

#[derive(Debug, PartialEq)]
pub(crate) struct Plan {
    pub target: Target,
    /// The format as the caller wrote it, which names the output:
    /// `--convert-to .pdf` writes `name..pdf`, as LibreOffice does.
    pub extension: String,
    pub outdir: PathBuf,
    pub inputs: Vec<PathBuf>,
    /// Filter options after `pdf:filter:` — read by LibreOffice, not by Oxi.
    pub ignored_options: Option<String>,
}

#[derive(Debug, PartialEq)]
pub(crate) enum Parsed {
    Version,
    Help,
    Convert(Plan),
    /// A command line Oxi does not carry out, and why.
    NotOurs(String),
}

/// Options LibreOffice takes that change nothing for a conversion.
const HARMLESS: &[&str] = &[
    "headless", "invisible", "nologo", "norestore", "nofirststartwizard", "nodefault",
    "nolockcheck", "minimized", "nocrashreport", "quickstart=no", "writer", "calc",
    "impress", "draw", "web", "global",
];

/// Read a LibreOffice command line. Options take one dash or two, as
/// LibreOffice's do; `-env:…` settings (a throwaway profile, mostly) are
/// LibreOffice's own business and are passed over.
pub(crate) fn parse(args: &[String]) -> Parsed {
    let mut target = None;
    let mut extension = String::new();
    let mut options = None;
    let mut outdir = None;
    let mut inputs = Vec::new();
    let mut rest = args.iter();
    while let Some(arg) = rest.next() {
        if arg.starts_with("-env:") {
            continue;
        }
        let Some(option) = arg.strip_prefix("--").or_else(|| arg.strip_prefix('-').filter(|o| o.len() > 1)) else {
            inputs.push(PathBuf::from(arg));
            continue;
        };
        let (name, inline) = match option.split_once('=') {
            Some((name, value)) if name == "convert-to" || name == "outdir" => (name, Some(value.to_string())),
            _ => (option, None),
        };
        match name {
            "version" => return Parsed::Version,
            "help" | "h" | "?" => return Parsed::Help,
            "convert-to" => {
                let Some(spec) = inline.or_else(|| rest.next().cloned()) else {
                    return Parsed::NotOurs("--convert-to needs a format".into());
                };
                let mut parts = spec.splitn(3, ':');
                let given = parts.next().unwrap_or("");
                let format = given.trim_start_matches('.').to_ascii_lowercase();
                extension = given.to_string();
                target = Some(match format.as_str() {
                    "pdf" => Target::Pdf,
                    "png" => Target::Png,
                    _ => return Parsed::NotOurs(format!("--convert-to {spec}: Oxi writes pdf and png")),
                });
                // The filter name is LibreOffice's choice of exporter; its
                // options (page range, PDF/A, …) Oxi does not read.
                let _filter = parts.next();
                options = parts.next().map(str::to_string);
            }
            "outdir" => {
                let Some(dir) = inline.or_else(|| rest.next().cloned()) else {
                    return Parsed::NotOurs("--outdir needs a directory".into());
                };
                outdir = Some(PathBuf::from(dir));
            }
            name if HARMLESS.contains(&name) || name.starts_with("infilter=") => {}
            other => return Parsed::NotOurs(format!("-{}{other}: not an option Oxi carries out", if arg.starts_with("--") { "-" } else { "" })),
        }
    }
    let Some(target) = target else {
        return Parsed::NotOurs("no --convert-to: Oxi only converts".into());
    };
    if inputs.is_empty() {
        return Parsed::NotOurs("no file to convert".into());
    }
    Parsed::Convert(Plan {
        target,
        extension,
        outdir: outdir.unwrap_or_else(|| PathBuf::from(".")),
        inputs,
        ignored_options: options,
    })
}

/// Whether Oxi reads this file: an Office Open XML extension, or — for a
/// name that says nothing, as `libreoffice-convert`'s `source` does — a zip
/// package holding a Word, PowerPoint or Excel main part. A name that says
/// something else (.odt, .doc, .pdf) is never ours.
pub(crate) fn is_ours(input: &Path) -> bool {
    let ext = input.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
    match ext.as_str() {
        "docx" | "docm" | "dotx" | "dotm" | "pptx" | "pptm" | "ppsx" | "ppsm" | "potx" | "potm"
        | "xlsx" | "xlsm" | "xltx" | "xltm" => true,
        "" | "tmp" | "bin" | "dat" => std::fs::read(input).map(|data| is_ooxml_package(&data)).unwrap_or(false),
        _ => false,
    }
}

/// A zip package naming one of the three main parts. A zip's central
/// directory stores part names uncompressed, so the bytes are searched as
/// `is_presentation` searches them.
fn is_ooxml_package(data: &[u8]) -> bool {
    if !data.starts_with(b"PK\x03\x04") {
        return false;
    }
    [&b"word/document.xml"[..], b"ppt/presentation.xml", b"xl/workbook.xml"]
        .iter()
        .any(|part| data.windows(part.len()).any(|w| w == *part))
}

/// Where LibreOffice writes `input` converted: the output directory, the
/// file's name without its last extension, a dot, and the format exactly as
/// the caller wrote it.
pub(crate) fn output_path(outdir: &Path, input: &Path, extension: &str) -> PathBuf {
    let mut name = input.file_stem().map(|s| s.to_os_string()).unwrap_or_default();
    name.push(".");
    name.push(extension);
    outdir.join(name)
}

fn version_line() -> String {
    format!(
        "LibreOffice-compatible oxidocs {} (Oxi: .docx .pptx .xlsx to pdf and png; anything else to ${FALLBACK_VAR})",
        env!("CARGO_PKG_VERSION")
    )
}

fn help() -> String {
    format!(
        "{}\n\nUsage: soffice [--headless] --convert-to pdf|png [--outdir DIR] FILE...\n\
         \n\
         Writes DIR/<name>.pdf, or DIR/<name>.png of the first page, for each .docx,\n\
         .pptx or .xlsx (and their macro and template variants). Other files, other\n\
         targets and other options go to the LibreOffice named by ${FALLBACK_VAR},\n\
         with the same arguments; without it they are an error.",
        version_line()
    )
}

/// Hand the whole command line to the real LibreOffice, or say why not.
fn fall_back(args: &[String], why: &str) -> i32 {
    match std::env::var_os(FALLBACK_VAR).filter(|v| !v.is_empty()) {
        Some(soffice) => match std::process::Command::new(&soffice).args(args).status() {
            Ok(status) => status.code().unwrap_or(1),
            Err(e) => {
                eprintln!("Error: {why}; and the fallback {} did not run: {e}", Path::new(&soffice).display());
                1
            }
        },
        None => {
            eprintln!("Error: {why}. Set {FALLBACK_VAR} to a LibreOffice soffice to have it do this.");
            1
        }
    }
}

/// Carry out a LibreOffice command line; the process exit status.
///
/// Oxi's progress notes stay off: LibreOffice writes nothing of the kind,
/// and some callers take any stderr naming an error as a failed run.
pub(crate) fn run(args: &[String]) -> i32 {
    super::QUIET.store(true, std::sync::atomic::Ordering::Relaxed);
    let plan = match parse(args) {
        Parsed::Version => {
            println!("{}", version_line());
            return 0;
        }
        Parsed::Help => {
            println!("{}", help());
            return 0;
        }
        Parsed::NotOurs(why) => return fall_back(args, &why),
        Parsed::Convert(plan) => plan,
    };
    // One file Oxi does not read sends the whole run to LibreOffice, as the
    // caller asked for one conversion of all of them.
    if let Some(other) = plan.inputs.iter().find(|i| !is_ours(i)) {
        return fall_back(args, &format!("{} is not a .docx, .pptx or .xlsx", other.display()));
    }
    if let Some(options) = &plan.ignored_options {
        eprintln!("Warning: export options {options:?} are LibreOffice's; Oxi converts without them");
    }
    if let Err(e) = std::fs::create_dir_all(&plan.outdir) {
        eprintln!("Error: cannot create {}: {e}", plan.outdir.display());
        return 1;
    }
    let mut failed = 0;
    for input in &plan.inputs {
        let output = output_path(&plan.outdir, input, &plan.extension);
        let (inp, out) = (input.to_string_lossy(), output.to_string_lossy());
        let result = std::panic::catch_unwind(|| match plan.target {
            Target::Pdf => super::convert(&inp, &out, false),
            Target::Png => super::convert_png(&inp, &out, false, 1, super::raster::Size { width: None, height: None, trim: false }),
        });
        match result {
            Ok(Ok(_)) => println!("convert {inp} -> {out} using filter : oxidocs"),
            Ok(Err(super::ConvertError::Failed(why))) => {
                eprintln!("Error: {why}");
                failed += 1;
            }
            Ok(Err(super::ConvertError::Unsupported(_))) => unreachable!("not strict"),
            Err(_) => {
                eprintln!("Error: Oxi stopped on {inp}");
                failed += 1;
            }
        }
    }
    if failed > 0 { 1 } else { 0 }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(line: &str) -> Vec<String> {
        line.split_whitespace().map(str::to_string).collect()
    }

    #[test]
    fn the_command_lines_callers_send_are_read_as_libreoffice_reads_them() {
        // cloudreve's thumbnailer.
        let p = parse(&args("--headless --nologo --nofirststartwizard --invisible --norestore --convert-to png --outdir /tmp/t/soffice_1 /data/a.docx"));
        assert_eq!(p, Parsed::Convert(Plan {
            target: Target::Png,
            extension: "png".into(),
            outdir: "/tmp/t/soffice_1".into(),
            inputs: vec!["/data/a.docx".into()],
            ignored_options: None,
        }));
        // libreoffice-convert, with a throwaway profile and a filter.
        let p = parse(&args("-env:UserInstallation=file:///tmp/p --headless --convert-to pdf:writer_pdf_Export --outdir /o /i/b.xlsx /i/c.pptx"));
        let Parsed::Convert(plan) = p else { panic!("{p:?}") };
        assert_eq!((plan.target, plan.inputs.len()), (Target::Pdf, 2));
        // One dash, `=` forms, and no outdir (the working directory).
        let p = parse(&args("-headless -convert-to=pdf x.docx"));
        let Parsed::Convert(plan) = p else { panic!("{p:?}") };
        assert_eq!(plan.outdir, PathBuf::from("."));
    }

    #[test]
    fn what_oxi_does_not_do_is_not_taken() {
        for line in [
            "--headless --convert-to docx --outdir o a.odt",
            "--headless --accept=socket,host=localhost,port=2002;urp;",
            "--headless a.docx",
            "--headless --convert-to pdf",
            "-p a.docx",
        ] {
            assert!(matches!(parse(&args(line)), Parsed::NotOurs(_)), "{line}");
        }
        assert_eq!(parse(&args("--version")), Parsed::Version);
    }

    #[test]
    fn the_output_is_named_as_libreoffice_names_it() {
        let out = output_path(Path::new("/o"), Path::new("/in/Report 2026.final.docx"), "pdf");
        assert_eq!(out, PathBuf::from("/o/Report 2026.final.pdf"));
        // libreoffice-convert's README passes '.pdf' and reads `source..pdf`.
        let Parsed::Convert(plan) = parse(&args("--headless --convert-to .pdf --outdir /t /t/source")) else { panic!() };
        assert_eq!((plan.target, output_path(&plan.outdir, &plan.inputs[0], &plan.extension)), (Target::Pdf, PathBuf::from("/t/source..pdf")));
        assert!(is_ours(Path::new("a.XLSM")) && is_ours(Path::new("deck.ppsx")));
        assert!(!is_ours(Path::new("a.odt")) && !is_ours(Path::new("a.doc")) && !is_ours(Path::new("noext")));
        // A package named `source`, as libreoffice-convert writes it, by its parts.
        let dir = std::env::temp_dir().join(format!("oxi-soffice-test-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let docx = dir.join("source");
        std::fs::write(&docx, b"PK\x03\x04....word/document.xml....").unwrap();
        let odt = dir.join("other");
        std::fs::write(&odt, b"PK\x03\x04....content.xml....").unwrap();
        assert!(is_ours(&docx) && !is_ours(&odt));
        std::fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn it_answers_as_libreoffice_only_when_called_so() {
        assert!(invocation(&args("/usr/bin/soffice --version")).is_some());
        assert!(invocation(&args("/opt/oxi/libreoffice --version")).is_some());
        assert_eq!(invocation(&args("oxidocs soffice --version")).map(|a| a.len()), Some(1));
        assert!(invocation(&args("oxidocs docx-to-pdf a.docx a.pdf")).is_none());
        // A caller's version check looks for this word.
        assert!(version_line().contains("LibreOffice"));
    }
}
