// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Project storage invalidation when edited VBA is compiled again.

use crate::lexer::{LexError, Punct, Token, TokenKind};
use crate::parser::{module_tokens, project_blocks, project_sources};

#[derive(Default, PartialEq, Eq)]
struct ModuleDeclarations {
    declarations: Vec<String>,
    procedures: Vec<ProcedureDeclarations>,
}

#[derive(Default, PartialEq, Eq)]
struct ProcedureDeclarations {
    header: String,
    // Ordinary locals in a non-Static procedure do not own persistent storage.
    locals: Vec<String>,
    persistent: bool,
}

/// Whether the edited declarations require fresh project storage. Body and
/// standalone comment edits can resume existing slots and object identities.
/// Declaration spelling is intentionally retained: editing a declaration line
/// in the VBE can reset storage even when its type/layout is unchanged.
pub fn project_edit_requires_reset(previous: &str, edited: &str) -> Result<bool, LexError> {
    if previous == edited {
        return Ok(false);
    }
    let (old_standard, old_classes) = project_sources(previous);
    let (new_standard, new_classes) = project_sources(edited);
    if declarations(&old_standard)? != declarations(&new_standard)? {
        return Ok(true);
    }
    for (old_name, old_source) in &old_classes {
        let Some((_, new_source)) = new_classes.iter().find(|(name, _)| name == old_name) else {
            return Ok(true);
        };
        if declarations(old_source)? != declarations(new_source)? {
            return Ok(true);
        }
    }
    for (new_name, new_source) in &new_classes {
        if !old_classes.iter().any(|(name, _)| name == new_name) {
            let footprint = declarations(new_source)?;
            if !footprint.procedures.is_empty()
                || footprint.declarations.iter().any(|text| {
                    let lower = text.to_ascii_lowercase();
                    !lower.starts_with("attribute vb_name")
                        && !(lower.starts_with("version ") && lower.contains("class"))
                })
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn word(token: Option<&Token>) -> Option<&str> {
    match token.map(|token| &token.kind) {
        Some(TokenKind::Ident(text)) => Some(text),
        _ => None,
    }
}

fn is_word(token: Option<&Token>, expected: &str) -> bool {
    word(token).is_some_and(|text| text.eq_ignore_ascii_case(expected))
}

fn header(tokens: &[Token]) -> Option<bool> {
    let mut at = 0;
    let mut persistent = false;
    while word(tokens.get(at)).is_some_and(|text| {
        matches!(
            text.to_ascii_lowercase().as_str(),
            "public" | "private" | "friend" | "static"
        )
    }) {
        persistent |= is_word(tokens.get(at), "static");
        at += 1;
    }
    word(tokens.get(at))
        .filter(|text| {
            matches!(
                text.to_ascii_lowercase().as_str(),
                "sub" | "function" | "property"
            )
        })
        .map(|_| persistent)
}

fn declarations(source: &str) -> Result<ModuleDeclarations, LexError> {
    Ok(scan_declarations(source)?.0)
}

fn scan_declarations(source: &str) -> Result<(ModuleDeclarations, Vec<(usize, usize)>), LexError> {
    let tokens = module_tokens(source)?;
    let mut result = ModuleDeclarations::default();
    let mut protected = Vec::new();
    let mut procedure: Option<ProcedureDeclarations> = None;
    let mut start = 0;
    for (end, token) in tokens.iter().enumerate() {
        if !matches!(
            token.kind,
            TokenKind::Eol | TokenKind::Eof | TokenKind::Punct(Punct::Colon)
        ) {
            continue;
        }
        let statement = &tokens[start..end];
        start = end + 1;
        if statement.is_empty() || matches!(statement[0].kind, TokenKind::Comment(_)) {
            continue;
        }
        let raw = source[statement[0].span.start..token.span.start]
            .trim()
            .replace("\r\n", "\n");
        if let Some(current) = &mut procedure {
            if is_word(statement.first(), "end")
                && word(statement.get(1)).is_some_and(|text| {
                    matches!(
                        text.to_ascii_lowercase().as_str(),
                        "sub" | "function" | "property"
                    )
                })
            {
                result
                    .procedures
                    .push(procedure.take().expect("current procedure"));
            } else if is_word(statement.first(), "static")
                || (current.persistent
                    && word(statement.first()).is_some_and(|text| {
                        matches!(text.to_ascii_lowercase().as_str(), "dim" | "const")
                    }))
            {
                current.locals.push(raw);
                protected.push((statement[0].span.line as usize, token.span.line as usize));
            }
        } else if let Some(persistent) = header(statement) {
            procedure = Some(ProcedureDeclarations {
                header: raw,
                persistent,
                ..Default::default()
            });
            protected.push((statement[0].span.line as usize, token.span.line as usize));
        } else {
            result.declarations.push(raw);
            protected.push((statement[0].span.line as usize, token.span.line as usize));
        }
    }
    if let Some(unfinished) = procedure {
        result.procedures.push(unfinished);
    }
    Ok((result, protected))
}

/// A line operation in the exported project source. Line numbers are 1-based;
/// text contains zero or more complete lines, with an optional final newline.
pub struct SourceEdit {
    pub line: usize,
    pub delete_lines: usize,
    pub text: String,
}

pub struct SourceEditOutcome {
    pub source: String,
    pub reset: bool,
}

/// Apply actual editing operations as well as comparing their resulting code.
/// Replacing a declaration with identical text can reset the VBE project; a
/// final-source comparison alone cannot represent that operation.
pub fn apply_source_edits(
    previous: &str,
    edits: &[SourceEdit],
) -> Result<SourceEditOutcome, String> {
    let mut source = previous.replace("\r\n", "\n");
    if !source.is_empty() && !source.ends_with('\n') {
        source.push('\n');
    }
    let mut reset = false;
    let ranges = |text: &str| -> Result<Vec<(usize, usize)>, LexError> {
        let mut ranges = Vec::new();
        for block in project_blocks(text) {
            let (_, protected) = scan_declarations(&block.text)?;
            ranges.extend(
                protected
                    .into_iter()
                    .map(|(first, last)| (first + block.line - 1, last + block.line - 1)),
            );
        }
        Ok(ranges)
    };
    let mut protected = ranges(&source).map_err(|error| error.to_string())?;
    let mut last_valid = source.clone();
    for edit in edits {
        let mut starts = vec![0];
        starts.extend(source.match_indices('\n').map(|(at, _)| at + 1));
        let first = edit
            .line
            .checked_sub(1)
            .filter(|index| *index < starts.len())
            .ok_or_else(|| "invalid VBA source edit line".to_string())?;
        let end = first
            .checked_add(edit.delete_lines)
            .filter(|end| *end <= starts.len())
            .ok_or_else(|| "invalid VBA source edit deleteLines".to_string())?;
        if edit.delete_lines != 0 {
            let last_line = edit.line + edit.delete_lines - 1;
            reset |= protected
                .iter()
                .any(|(first, last)| edit.line <= *last && last_line >= *first);
        }
        let mut text = edit.text.replace("\r\n", "\n");
        if !text.is_empty() && !text.ends_with('\n') {
            text.push('\n');
        }
        let stop = starts.get(end).copied().unwrap_or(source.len());
        let next = format!("{}{}{}", &source[..starts[first]], text, &source[stop..]);
        let inserted = text.bytes().filter(|byte| *byte == b'\n').count();
        protected.retain_mut(|(first, last)| {
            if *last < edit.line {
                return true;
            }
            if *first >= edit.line + edit.delete_lines {
                if inserted >= edit.delete_lines {
                    *first += inserted - edit.delete_lines;
                    *last += inserted - edit.delete_lines;
                } else {
                    *first -= edit.delete_lines - inserted;
                    *last -= edit.delete_lines - inserted;
                }
                return true;
            }
            false
        });
        // Typing may pass through an unterminated string or other lexical
        // error. Keep applying the line operations; later valid code must not
        // make an earlier declaration edit disappear from the reset history.
        if let Ok(next_ranges) = ranges(&next) {
            reset |= project_edit_requires_reset(&last_valid, &next)
                .map_err(|error| error.to_string())?;
            protected = next_ranges;
            last_valid = next.clone();
        }
        source = next;
    }
    Ok(SourceEditOutcome { source, reset })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_snapshots_keep_body_comments_and_invocation_locals() {
        let source = "Private value As Long\nFunction NextValue() As Long\nStatic stored As Long\nstored = stored + 1\nNextValue = stored\nEnd Function\n";
        assert!(!project_edit_requires_reset(source, &format!("' comment\n{source}")).unwrap());
        assert!(
            !project_edit_requires_reset(source, &source.replace("stored + 1", "stored + 2"))
                .unwrap()
        );
        assert!(!project_edit_requires_reset(
            source,
            &source.replace(
                "Static stored As Long",
                "Static stored As Long\nDim temporary As Long"
            )
        )
        .unwrap());
        assert!(project_edit_requires_reset(
            source,
            &source.replace("Private value As Long", "Private value As Double")
        )
        .unwrap());
        assert!(project_edit_requires_reset(
            source,
            &source.replace("Static stored As Long", "Static stored As Double")
        )
        .unwrap());
    }

    #[test]
    fn edit_history_keeps_declaration_resets_after_undo_and_allows_incomplete_typing() {
        let source = "Function NextValue() As Long\nStatic stored As Long\nNextValue = stored\nEnd Function\n";
        let reverted = apply_source_edits(
            source,
            &[
                SourceEdit {
                    line: 2,
                    delete_lines: 1,
                    text: "Static stored As Double".into(),
                },
                SourceEdit {
                    line: 2,
                    delete_lines: 1,
                    text: "Static stored As Long".into(),
                },
            ],
        )
        .unwrap();
        assert_eq!(reverted.source, source);
        assert!(reverted.reset);
        let typing = apply_source_edits(
            source,
            &[
                SourceEdit {
                    line: 3,
                    delete_lines: 1,
                    text: "NextValue = Len(\"".into(),
                },
                SourceEdit {
                    line: 3,
                    delete_lines: 1,
                    text: "NextValue = stored".into(),
                },
            ],
        )
        .unwrap();
        assert_eq!(typing.source, source);
        assert!(!typing.reset);
        assert!(apply_source_edits(
            source,
            &[SourceEdit {
                line: 0,
                delete_lines: 0,
                text: String::new()
            }]
        )
        .is_err());
        assert!(apply_source_edits(
            source,
            &[SourceEdit {
                line: 1,
                delete_lines: usize::MAX,
                text: String::new()
            }]
        )
        .is_err());
    }

    #[test]
    fn declaration_edits_cover_continued_lines_and_class_offsets() {
        let source = "Sub Main()\nEnd Sub\nVERSION 1.0 CLASS\nAttribute VB_Name = \"StateClass\"\nPrivate values(1 To _\n2) As Long\n";
        let result = apply_source_edits(
            source,
            &[SourceEdit {
                line: 6,
                delete_lines: 1,
                text: "2) As Long".into(),
            }],
        )
        .unwrap();
        assert_eq!(result.source, source);
        assert!(result.reset);
    }

    #[test]
    fn japanese_designer_headers_keep_original_byte_offsets() {
        let source = "VERSION 5.00\nBegin VB.UserForm Form1\n Caption = \"日本語の文字列\"\nEnd\nAttribute VB_Name = \"Form1\"\nPrivate stored As Long\nPublic Function Value() As Long\nValue = stored + 1\nEnd Function\n";
        let edited = source.replace("stored + 1", "stored + 2");
        assert!(!project_edit_requires_reset(source, &edited).unwrap());
        let parsed = crate::parse_module(source).unwrap();
        let crate::ast::ModuleItem::Procedure(procedure) = parsed.items.last().unwrap() else { panic!("procedure after designer"); };
        assert_eq!(&source[procedure.span.start..procedure.span.end], "Public");
        let result = apply_source_edits(source, &[SourceEdit { line: 8, delete_lines: 1, text: "Value = stored + 2".into() }]).unwrap();
        assert_eq!(result.source, edited);
        assert!(!result.reset);
    }

    #[test]
    fn inactive_japanese_source_keeps_original_byte_offsets() {
        let source = "#If False Then\n日本語の文字列 = \"unterminated\n#End If\nPrivate stored As Long\nPublic Function Value() As Long\nValue = stored + 1\nEnd Function\n";
        assert_eq!(crate::precompile::live_source(source).len(), source.len());
        assert!(!project_edit_requires_reset(source, &source.replace("stored + 1", "stored + 2")).unwrap());
        let parsed = crate::parse_module(source).unwrap();
        let crate::ast::ModuleItem::Procedure(procedure) = parsed.items.last().unwrap() else { panic!("procedure after inactive code"); };
        assert_eq!(&source[procedure.span.start..procedure.span.end], "Public");
    }
}
