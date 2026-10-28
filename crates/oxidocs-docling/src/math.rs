// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! OMML expression tree → LaTeX, the form docling writes into a `formula`
//! item's `text`. Best effort: every construct Oxi parses gets a LaTeX
//! spelling, and anything unknown falls back to its characters.

use oxidocs_core::ir::math::{BarPos, LimLoc, LimitPos, MathBlock, MathExpr};

pub fn block_latex(b: &MathBlock) -> String {
    let content = match b {
        MathBlock::Inline(c) => c,
        MathBlock::Display { content, .. } => content,
    };
    tidy(&seq(content))
}

fn seq(v: &[MathExpr]) -> String {
    v.iter().map(latex).collect::<Vec<_>>().join("")
}

/// Script arguments always braced, as docling's OMML converter writes them (`r^{2}`).
fn arg(e: &MathExpr) -> String {
    format!("{{{}}}", latex(e))
}

fn latex(e: &MathExpr) -> String {
    use MathExpr::*;
    match e {
        Text(s) | Run { text: s, .. } => s.chars().map(ascii_math).collect(),
        Fraction { num, den, .. } => format!("\\frac{{{}}}{{{}}}", latex(num), latex(den)),
        Superscript { base, sup } => format!("{}^{}", latex(base), arg(sup)),
        Subscript { base, sub } => format!("{}_{}", latex(base), arg(sub)),
        SubSuperscript { base, sub, sup } => format!("{}_{}^{}", latex(base), arg(sub), arg(sup)),
        PreScript { base, sub, sup } => format!("{{}}_{}^{}{}", arg(sub), arg(sup), latex(base)),
        Radical { degree: Some(d), radicand } => format!("\\sqrt[{}]{{{}}}", latex(d), latex(radicand)),
        Radical { degree: None, radicand } => format!("\\sqrt{{{}}}", latex(radicand)),
        Nary { op, sub, sup, operand, lim_loc, .. } => {
            let mut s = nary_op(*op).to_string();
            if matches!(lim_loc, LimLoc::UndOvr) && (sub.is_some() || sup.is_some()) {
                s = format!("\\limits{s}");
            }
            if let Some(x) = sub { s += &format!("_{}", arg(x)); }
            if let Some(x) = sup { s += &format!("^{}", arg(x)); }
            format!("{s} {}", latex(operand))
        }
        Delimiter { beg, end, content, .. } => {
            format!("\\left{} {} \\right{}", delim(*beg), latex(content), delim(*end))
        }
        Function { name, arg: a } => {
            let n = latex(name);
            let n = if n.chars().all(|c| c.is_ascii_alphabetic()) && n.len() > 1 { format!("\\{n}") } else { n };
            format!("{n} {}", latex(a))
        }
        Matrix { rows, .. } => format!(
            "\\begin{{matrix}}{}\\end{{matrix}}",
            rows.iter().map(|r| r.iter().map(latex).collect::<Vec<_>>().join(" & ")).collect::<Vec<_>>().join(" \\\\ ")
        ),
        Accent { accent, base } => format!("{}{{{}}}", accent_cmd(*accent), latex(base)),
        Bar { pos: BarPos::Top, base } => format!("\\overline{{{}}}", latex(base)),
        Bar { pos: BarPos::Bot, base } => format!("\\underline{{{}}}", latex(base)),
        Limit { base, lim, pos: LimitPos::Lower } => format!("{}_{}", latex(base), arg(lim)),
        Limit { base, lim, pos: LimitPos::Upper } => format!("{}^{}", latex(base), arg(lim)),
        GroupChar { chr, pos, base } => {
            let cmd = match (chr, pos) {
                ('⏞', _) | (_, BarPos::Top) => "\\overbrace",
                _ => "\\underbrace",
            };
            format!("{cmd}{{{}}}", latex(base))
        }
        EqArray(rows) => format!(
            "\\begin{{aligned}}{}\\end{{aligned}}",
            rows.iter().map(latex).collect::<Vec<_>>().join(" \\\\ ")
        ),
        BoxExpr(x) | Phantom(x) => latex(x),
        BorderBox { base, .. } => format!("\\boxed{{{}}}", latex(base)),
        Seq(v) => seq(v),
    }
}

fn delim(c: char) -> String {
    match c {
        '{' => "\\{".into(),
        '}' => "\\}".into(),
        '\0' | ' ' => ".".into(),
        '⟨' => "\\langle".into(),
        '⟩' => "\\rangle".into(),
        '‖' => "\\|".into(),
        c => c.to_string(),
    }
}

fn nary_op(c: char) -> &'static str {
    match c {
        '∑' => "\\sum", '∏' => "\\prod", '∐' => "\\coprod",
        '∫' => "\\int", '∬' => "\\iint", '∭' => "\\iiint", '∮' => "\\oint",
        '⋃' => "\\bigcup", '⋂' => "\\bigcap", '⋁' => "\\bigvee", '⋀' => "\\bigwedge",
        _ => "\\int",
    }
}

fn accent_cmd(c: char) -> &'static str {
    match c {
        '\u{302}' | '^' => "\\hat", '\u{303}' | '~' => "\\tilde", '\u{304}' | '¯' => "\\bar",
        '\u{307}' | '˙' => "\\dot", '\u{308}' | '¨' => "\\ddot", '\u{20D7}' | '→' => "\\vec",
        '\u{306}' | '˘' => "\\breve", '\u{30C}' | 'ˇ' => "\\check", '\u{301}' | '´' => "\\acute",
        '\u{300}' | '`' => "\\grave",
        _ => "\\hat",
    }
}

/// Mathematical alphanumeric symbols (U+1D400–U+1D7FF) back to the plain
/// letter, digit or Greek letter they style; Greek letters to their LaTeX name.
fn ascii_math(c: char) -> String {
    let cp = c as u32;
    let plain = match cp {
        0x1D400..=0x1D6A3 => {
            let i = (cp - 0x1D400) % 52;
            Some(if i < 26 { (b'A' + i as u8) as char } else { (b'a' + (i - 26) as u8) as char })
        }
        0x1D6A8..=0x1D7CB => {
            let i = (cp - 0x1D6A8) % 58;
            Some(match i {
                0..=16 => char::from_u32(0x391 + i).unwrap(),
                17 => 'ϴ',
                18..=24 => char::from_u32(0x3A3 + i - 18).unwrap(),
                25 => '∇',
                26..=50 => char::from_u32(0x3B1 + i - 26).unwrap(),
                51 => '∂', 52 => 'ϵ', 53 => 'ϑ', 54 => 'ϰ', 55 => 'ϕ', 56 => 'ϱ', _ => 'ϖ',
            })
        }
        0x1D7CE..=0x1D7FF => Some((b'0' + ((cp - 0x1D7CE) % 10) as u8) as char),
        0x210E => Some('h'),
        _ => None,
    };
    let c = plain.unwrap_or(c);
    greek(c).map(|g| format!("\\{g} ")).unwrap_or_else(|| match c {
        '×' => "\\times ".into(), '·' | '⋅' => "\\cdot ".into(), '±' => "\\pm ".into(),
        '≤' => "\\leq ".into(), '≥' => "\\geq ".into(), '≠' => "\\neq ".into(), '≈' => "\\approx ".into(),
        '∞' => "\\infty ".into(), '→' => "\\rightarrow ".into(), '∈' => "\\in ".into(), '∂' => "\\partial ".into(),
        '∇' => "\\nabla ".into(), '−' => "-".into(), '…' => "\\ldots ".into(), '⋯' => "\\cdots ".into(),
        '\u{2061}'..='\u{2064}' => String::new(), // invisible operators
        c => c.to_string(),
    })
}

fn greek(c: char) -> Option<&'static str> {
    Some(match c {
        'α' => "alpha", 'β' => "beta", 'γ' => "gamma", 'δ' => "delta", 'ε' | 'ϵ' => "epsilon", 'ζ' => "zeta",
        'η' => "eta", 'θ' => "theta", 'ϑ' => "vartheta", 'ι' => "iota", 'κ' => "kappa", 'λ' => "lambda",
        'μ' => "mu", 'ν' => "nu", 'ξ' => "xi", 'π' => "pi", 'ρ' => "rho", 'σ' => "sigma", 'ς' => "varsigma",
        'τ' => "tau", 'υ' => "upsilon", 'φ' | 'ϕ' => "phi", 'χ' => "chi", 'ψ' => "psi", 'ω' => "omega",
        'Γ' => "Gamma", 'Δ' => "Delta", 'Θ' => "Theta", 'Λ' => "Lambda", 'Ξ' => "Xi", 'Π' => "Pi",
        'Σ' => "Sigma", 'Υ' => "Upsilon", 'Φ' => "Phi", 'Ψ' => "Psi", 'Ω' => "Omega",
        _ => return None,
    })
}

/// Collapse the spaces the Greek/operator spellings leave before a brace or at the end.
fn tidy(s: &str) -> String {
    let mut out = s.replace(" }", "}").replace(" ^", "^").replace(" _", "_");
    while out.contains("  ") {
        out = out.replace("  ", " ");
    }
    out.trim().to_string()
}
