// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Ported from docling's docling/backend/docx/latex/omml.py
// (Copyright The Docling Contributors, MIT License), which was itself adapted
// from dwml's dwml/omml.py (https://github.com/xiilei/dwml, Apache License 2.0).
// Changes: rewritten in Rust over a roxmltree tree. Both licenses' terms are in
// THIRD_PARTY_NOTICES.md and licenses/ next to this crate's Cargo.toml.

//! OMML (`m:oMath`) → LaTeX, spelled the way docling's own converter spells it
//! (`docling/backend/docx/latex/omml.py`, itself adapted from dwml), so a
//! `formula` item's `text` matches docling's output for the same equation.
//! Ported rule for rule, including the quirks (`\int_{}^{}`, `\log(\left(x\right))`,
//! `\text{ \textellipsis }`), since consumers compare against docling's text.

use std::collections::HashMap;

use roxmltree::Node;

const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const CHARS: &[char] = &['{', '}', '_', '^', '#', '&', '$', '%', '~'];
const MATH_CHARS: &[char] = &['\\', '<', '>', '=', '+', '*', '/', '^', '_', '{', '}'];
const FUNC_PLACE: &str = "{fe}";
const BRK: &str = "\\\\";

enum Out {
    S(String),
    Pr(Pr),
}

#[derive(Default)]
struct Pr {
    text: String,
    vals: HashMap<String, String>,
}

impl Pr {
    fn new(elm: Node) -> Pr {
        let mut pr = Pr::default();
        for c in elm.children().filter(|c| c.is_element() && c.tag_name().namespace() == Some(M)) {
            match c.tag_name().name() {
                "brk" => {
                    pr.vals.insert("brk".into(), BRK.into());
                    pr.text.push_str(BRK);
                }
                t @ ("chr" | "pos" | "begChr" | "endChr" | "type") => {
                    if let Some(v) = c.attribute((M, "val")) {
                        pr.vals.insert(t.into(), v.into());
                    }
                }
                _ => {}
            }
        }
        pr
    }
    fn get(&self, k: &str) -> Option<&str> {
        self.vals.get(k).map(String::as_str)
    }
}

fn get_val(key: Option<&str>, default: &str, store: &[(&str, &str)]) -> String {
    match key {
        Some(k) => store.iter().find(|(a, _)| *a == k).map(|(_, b)| b.to_string()).unwrap_or_else(|| k.to_string()),
        None => default.to_string(),
    }
}

fn escape_latex(s: &str) -> String {
    let s = s.replace("\\\\", "\\");
    let mut out = String::new();
    let mut last = None;
    for c in s.chars() {
        if CHARS.contains(&c) && last != Some('\\') {
            out.push('\\');
        }
        out.push(c);
        last = Some(c);
    }
    out
}

/// Convert one `m:oMath` element.
pub fn omath_to_latex(omath: Node) -> String {
    children_str(omath, None).replace("  ", " ")
}

fn children(elm: Node, include: Option<&[&str]>) -> Vec<(String, Out)> {
    let mut v = vec![];
    for c in elm.children().filter(|c| c.is_element() && c.tag_name().namespace() == Some(M)) {
        let tag = c.tag_name().name();
        if include.is_some_and(|inc| !inc.contains(&tag)) {
            continue;
        }
        if let Some(o) = call(c, tag) {
            v.push((tag.to_string(), o));
        }
    }
    v
}

fn children_str(elm: Node, include: Option<&[&str]>) -> String {
    children(elm, include).into_iter().map(|(_, o)| match o {
        Out::S(s) => s,
        Out::Pr(p) => p.text,
    }).collect()
}

fn children_dict(elm: Node, include: Option<&[&str]>) -> HashMap<String, Out> {
    children(elm, include).into_iter().collect()
}

fn s(d: &HashMap<String, Out>, k: &str) -> String {
    match d.get(k) {
        Some(Out::S(x)) => x.clone(),
        Some(Out::Pr(p)) => p.text.clone(),
        None => String::new(),
    }
}

fn pr<'a>(d: &'a HashMap<String, Out>, k: &str) -> Option<&'a Pr> {
    match d.get(k) {
        Some(Out::Pr(p)) => Some(p),
        _ => None,
    }
}

fn fmt1(template: &str, e: &str) -> String {
    if template.contains("%s") { template.replacen("%s", e, 1) } else { template.to_string() }
}

fn call(elm: Node, tag: &str) -> Option<Out> {
    let out = match tag {
        "acc" => {
            let d = children_dict(elm, None);
            let t = get_val(pr(&d, "accPr").and_then(|p| p.get("chr")), "\\hat{%s}", CHR);
            fmt1(&t, &s(&d, "e"))
        }
        "bar" => {
            let d = children_dict(elm, None);
            let p = pr(&d, "barPr");
            let t = get_val(p.and_then(|p| p.get("pos")), "\\overline{%s}", POS);
            format!("{}{}", p.map(|p| p.text.as_str()).unwrap_or(""), fmt1(&t, &s(&d, "e")))
        }
        "d" => {
            let d = children_dict(elm, None);
            let p = pr(&d, "dPr");
            let sv = get_val(p.and_then(|p| p.get("begChr")), "(", T);
            let ev = get_val(p.and_then(|p| p.get("endChr")), ")", T);
            let left = if sv.is_empty() { ".".to_string() } else { escape_latex(&sv) };
            let right = if ev.is_empty() { ".".to_string() } else { escape_latex(&ev) };
            format!("{}\\left{left}{}\\right{right}", p.map(|p| p.text.as_str()).unwrap_or(""), s(&d, "e"))
        }
        "sPre" => return None,
        "sSub" => {
            let d = children_dict(elm, Some(&["e", "sub", "sSubPr"]));
            let base = group(s(&d, "e").trim_end());
            format!("{base}_{{{}}}", unwrap(&s(&d, "sub"), "_"))
        }
        "sSup" => {
            let d = children_dict(elm, Some(&["e", "sup", "sSupPr"]));
            let base = group(s(&d, "e").trim_end());
            format!("{base}^{{{}}}", unwrap(&s(&d, "sup"), "^"))
        }
        "sSubSup" => {
            let d = children_dict(elm, Some(&["e", "sub", "sup", "sSubSupPr"]));
            let base = group(s(&d, "e").trim_end());
            format!("{base}_{{{}}}^{{{}}}", unwrap(&s(&d, "sub"), "_"), unwrap(&s(&d, "sup"), "^"))
        }
        "sub" => format!("_{{{}}}", children_str(elm, None)),
        "sup" => format!("^{{{}}}", children_str(elm, None)),
        "f" => {
            let d = children_dict(elm, None);
            let (num, den) = (s(&d, "num"), s(&d, "den"));
            match pr(&d, "fPr") {
                None => format!("\\frac{{{num}}}{{{den}}}"),
                Some(p) => {
                    let t = get_val(p.get("type"), "\\frac{%(num)s}{%(den)s}", F);
                    format!("{}{}", p.text, t.replace("%(num)s", &num).replace("%(den)s", &den))
                }
            }
        }
        "func" => {
            let d = children_dict(elm, None);
            s(&d, "fName").replace(FUNC_PLACE, &s(&d, "e"))
        }
        "fName" => {
            let mut parts = vec![];
            for (tag, o) in children(elm, None) {
                if let Out::S(t) = o {
                    if tag == "r" {
                        parts.push(FUNC.iter().find(|(k, _)| *k == t).map(|(_, v)| v.to_string()).unwrap_or(t));
                    } else {
                        parts.push(t);
                    }
                }
            }
            let t: String = parts.concat();
            if t.contains(FUNC_PLACE) { t } else { t + FUNC_PLACE }
        }
        "groupChr" => {
            let d = children_dict(elm, None);
            let p = pr(&d, "groupChrPr");
            let t = get_val(p.and_then(|p| p.get("chr")), "\\underbrace{%s}", CHR);
            format!("{}{}", p.map(|p| p.text.as_str()).unwrap_or(""), fmt1(&t, &s(&d, "e")))
        }
        "rad" => {
            let d = children_dict(elm, None);
            let (e, deg) = (s(&d, "e"), s(&d, "deg"));
            if deg.is_empty() { format!("\\sqrt{{{e}}}") } else { format!("\\sqrt[{deg}]{{{e}}}") }
        }
        "eqArr" => children(elm, Some(&["e"])).into_iter().filter_map(|(_, o)| match o { Out::S(x) => Some(x), _ => None })
            .collect::<Vec<_>>().join(BRK),
        "limLow" => {
            let d = children_dict(elm, Some(&["e", "lim"]));
            let (base, lim) = (s(&d, "e"), s(&d, "lim"));
            match LIM_FUNC.iter().find(|(k, _)| *k == base) {
                Some((_, t)) => t.replace("%(lim)s", &lim),
                None => format!("{base}_{{{lim}}}"),
            }
        }
        "limUpp" => {
            let d = children_dict(elm, Some(&["e", "lim"]));
            format!("\\overset{{{}}}{{{}}}", s(&d, "lim"), s(&d, "e"))
        }
        "lim" => {
            let mut r = children_str(elm, None).replace("\\rightarrow", "\\to");
            r = r.trim_end().to_string();
            if let Some(x) = r.strip_suffix(BRK) {
                r = x.trim_end().to_string();
            }
            if !r.is_empty() && !r.contains(MATH_CHARS) {
                r = r.replace(' ', "\\ ");
            }
            r
        }
        "m" => {
            let rows: Vec<String> = children(elm, None).into_iter()
                .filter_map(|(t, o)| match (t.as_str(), o) { ("mr", Out::S(x)) => Some(x), _ => None }).collect();
            format!("\\begin{{matrix}}{}\\end{{matrix}}", rows.join(BRK))
        }
        "mr" => children(elm, Some(&["e"])).into_iter().filter_map(|(_, o)| match o { Out::S(x) => Some(x), _ => None })
            .collect::<Vec<_>>().join("&"),
        "nary" => {
            let mut bo = String::new();
            let mut rest = String::new();
            for (t, o) in children(elm, None) {
                match (t.as_str(), o) {
                    ("naryPr", Out::Pr(p)) => bo = get_val(p.get("chr"), "\\int", CHR_BO),
                    (_, Out::S(x)) => rest.push_str(&x),
                    (_, Out::Pr(p)) => rest.push_str(&p.text),
                }
            }
            bo + &rest
        }
        "r" => run(elm),
        // unknown: direct containers, property bags, or nothing
        "box" | "num" | "den" | "deg" | "e" => children_str(elm, None),
        t if t.ends_with("Pr") => return Some(Out::Pr(Pr::new(elm))),
        _ => return None,
    };
    Some(Out::S(out))
}

fn group(base: &str) -> String {
    if base.contains("\\frac") || base.contains("\\sqrt") { format!("{{{base}}}") } else { base.to_string() }
}

fn unwrap(script: &str, marker: &str) -> String {
    let prefix = format!("{marker}{{");
    match script.strip_prefix(prefix.as_str()).and_then(|x| x.strip_suffix('}')) {
        Some(inner) => inner.to_string(),
        None => script.to_string(),
    }
}

fn run(elm: Node) -> String {
    let Some(text) = elm.children().find(|c| c.is_element() && c.tag_name() == (M, "t").into()).and_then(|t| t.text()) else {
        return String::new();
    };
    if text.is_empty() {
        return String::new();
    }
    let raw: String = text.chars().map(unicode).collect();
    let mut proc = escape_latex(&raw);
    if !text.contains('{') && proc.contains("\\{") {
        proc = proc.replace("\\{", "{");
    }
    if !text.contains('}') && proc.contains("\\}") {
        proc = proc.replace("\\}", "}");
    }
    // `^` is mapped to a math operator on purpose; escape_latex re-escaped it.
    if text.contains('^') && proc.contains("\\^") {
        proc = proc.replace("\\^", "^");
    }
    proc
}

/// docling's `process_unicode`: a few operators mapped directly, everything
/// else through pylatexenc (`braces-all`, unknown kept), then `{\cmd}` becomes
/// ` \cmd ` and a text-mode macro is wrapped in `\text{ … }`.
fn unicode(c: char) -> String {
    match c {
        '\u{2013}' | '\u{2014}' | '\u{2212}' => return "-".into(),
        '^' => return "^".into(),
        '\u{00D7}' => return "\\times ".into(),
        _ => {}
    }
    if c.is_ascii() {
        return match c {
            '{' => "\\{".into(),
            '}' => "\\}".into(),
            '#' | '$' | '%' | '&' | '_' => format!("\\{c}"),
            '\\' => " \\text{ \\textbackslash } ".into(),
            '~' => " \\text{ \\textasciitilde } ".into(),
            _ => c.to_string(),
        };
    }
    let Some(cmd) = unicode_cmd(c) else { return c.to_string() };
    if cmd.starts_with("\\text") {
        format!(" \\text{{ {cmd} }} ")
    } else {
        format!(" {cmd} ")
    }
}

fn unicode_cmd(c: char) -> Option<&'static str> {
    Some(match c {
        'α' => "\\alpha", 'β' => "\\beta", 'γ' => "\\gamma", 'δ' => "\\delta", 'ε' => "\\varepsilon", 'ϵ' => "\\epsilon",
        'ζ' => "\\zeta", 'η' => "\\eta", 'θ' => "\\theta", 'ϑ' => "\\vartheta", 'ι' => "\\iota", 'κ' => "\\kappa",
        'λ' => "\\lambda", 'μ' => "\\mu", 'ν' => "\\nu", 'ξ' => "\\xi", 'π' => "\\pi", 'ρ' => "\\rho", 'σ' => "\\sigma",
        'ς' => "\\varsigma", 'τ' => "\\tau", 'υ' => "\\upsilon", 'φ' => "\\varphi", 'ϕ' => "\\phi", 'χ' => "\\chi",
        'ψ' => "\\psi", 'ω' => "\\omega", 'Γ' => "\\Gamma", 'Δ' => "\\Delta", 'Θ' => "\\Theta", 'Λ' => "\\Lambda",
        'Ξ' => "\\Xi", 'Π' => "\\Pi", 'Σ' => "\\Sigma", 'Υ' => "\\Upsilon", 'Φ' => "\\Phi", 'Ψ' => "\\Psi", 'Ω' => "\\Omega",
        '∞' => "\\infty", '∂' => "\\partial", '∇' => "\\nabla", '∑' => "\\sum", '∏' => "\\prod", '∫' => "\\int",
        '√' => "\\surd", '∩' => "\\cap", '∪' => "\\cup", '∈' => "\\in", '∉' => "\\notin", '∋' => "\\ni", '∅' => "\\emptyset",
        '≤' => "\\leq", '≥' => "\\geq", '≠' => "\\neq", '≈' => "\\approx", '≡' => "\\equiv", '∼' => "\\sim", '≃' => "\\simeq",
        '≅' => "\\cong", '∝' => "\\propto", '±' => "\\pm", '∓' => "\\mp", '⋅' => "\\cdot", '∘' => "\\circ", '∗' => "\\ast",
        '÷' => "\\div", '⊂' => "\\subset", '⊆' => "\\subseteq", '⊃' => "\\supset", '⊇' => "\\supseteq", '⊕' => "\\oplus",
        '⊗' => "\\otimes", '∧' => "\\wedge", '∨' => "\\vee", '¬' => "\\neg", '∀' => "\\forall", '∃' => "\\exists",
        '→' => "\\rightarrow", '←' => "\\leftarrow", '↔' => "\\leftrightarrow", '⇒' => "\\Rightarrow", '⇐' => "\\Leftarrow",
        '⇔' => "\\Leftrightarrow", '↦' => "\\mapsto", '∴' => "\\therefore", '∵' => "\\because", '∠' => "\\angle",
        '⊥' => "\\perp", '∥' => "\\parallel", '′' => "\\prime", '…' => "\\textellipsis", '⋯' => "\\cdots", '⋮' => "\\vdots",
        '⋱' => "\\ddots", '·' => "\\textperiodcentered", '°' => "\\textdegree", 'ℝ' => "\\mathbb{R}", 'ℕ' => "\\mathbb{N}",
        'ℤ' => "\\mathbb{Z}", 'ℚ' => "\\mathbb{Q}", 'ℂ' => "\\mathbb{C}", 'ℓ' => "\\ell", 'ℏ' => "\\hbar", 'ℵ' => "\\aleph",
        '⌀' => "\\emptyset",
        _ => return None,
    })
}

const CHR: &[(&str, &str)] = &[
    ("\u{0300}", "\\grave{%s}"), ("\u{0301}", "\\acute{%s}"), ("\u{0302}", "\\hat{%s}"), ("\u{0303}", "\\tilde{%s}"),
    ("\u{0304}", "\\bar{%s}"), ("\u{0305}", "\\overbar{%s}"), ("\u{0306}", "\\breve{%s}"), ("\u{0307}", "\\dot{%s}"),
    ("\u{0308}", "\\ddot{%s}"), ("\u{0309}", "\\ovhook{%s}"), ("\u{030a}", "\\ocirc{%s}"), ("\u{030c}", "\\check{%s}"),
    ("\u{0310}", "\\candra{%s}"), ("\u{0312}", "\\oturnedcomma{%s}"), ("\u{0315}", "\\ocommatopright{%s}"),
    ("\u{031a}", "\\droang{%s}"), ("\u{0338}", "\\not{%s}"), ("\u{20d0}", "\\leftharpoonaccent{%s}"),
    ("\u{20d1}", "\\rightharpoonaccent{%s}"), ("\u{20d2}", "\\vertoverlay{%s}"), ("\u{20d6}", "\\overleftarrow{%s}"),
    ("\u{20d7}", "\\vec{%s}"), ("\u{20db}", "\\dddot{%s}"), ("\u{20dc}", "\\ddddot{%s}"), ("\u{20e1}", "\\overleftrightarrow{%s}"),
    ("\u{20e7}", "\\annuity{%s}"), ("\u{20e9}", "\\widebridgeabove{%s}"), ("\u{20f0}", "\\asteraccent{%s}"),
    ("\u{0330}", "\\wideutilde{%s}"), ("\u{0331}", "\\underbar{%s}"), ("\u{20e8}", "\\threeunderdot{%s}"),
    ("\u{20ec}", "\\underrightharpoondown{%s}"), ("\u{20ed}", "\\underleftharpoondown{%s}"), ("\u{20ee}", "\\underledtarrow{%s}"),
    ("\u{20ef}", "\\underrightarrow{%s}"), ("\u{23b4}", "\\overbracket{%s}"), ("\u{23dc}", "\\overparen{%s}"),
    ("\u{23de}", "\\overbrace{%s}"), ("\u{23b5}", "\\underbracket{%s}"), ("\u{23dd}", "\\underparen{%s}"), ("\u{23df}", "\\underbrace{%s}"),
];

const CHR_BO: &[(&str, &str)] = &[
    ("\u{2140}", "\\Bbbsum"), ("\u{220f}", "\\prod"), ("\u{2210}", "\\coprod"), ("\u{2211}", "\\sum"), ("\u{222b}", "\\int"),
    ("\u{222c}", "\\iint"), ("\u{222d}", "\\iiint"), ("\u{222e}", "\\oint"), ("\u{222f}", "\\oiint"), ("\u{2230}", "\\oiiint"),
    ("\u{22c0}", "\\bigwedge"), ("\u{22c1}", "\\bigvee"), ("\u{22c2}", "\\bigcap"), ("\u{22c3}", "\\bigcup"),
    ("\u{2a00}", "\\bigodot"), ("\u{2a01}", "\\bigoplus"), ("\u{2a02}", "\\bigotimes"),
];

/// docling's `T` table: the delimiter characters it maps (`get_val(..., store=T)`).
const T: &[(&str, &str)] = &[
    ("\u{2190}", "\\leftarrow "), ("\u{2191}", "\\uparrow "), ("\u{2192}", "\\rightarrow "), ("\u{2193}", "\\downright "),
    ("\u{2194}", "\\leftrightarrow "), ("\u{2195}", "\\updownarrow "), ("\u{22ee}", "\\vdots "), ("\u{22ef}", "\\cdots "),
    ("\u{2260}", "\\ne "), ("\u{2264}", "\\leq "), ("\u{2265}", "\\geq "), ("\u{221e}", "\\infty "), ("\u{00b1}", "\\pm "),
];

const POS: &[(&str, &str)] = &[("top", "\\overline{%s}"), ("bot", "\\underline{%s}")];

const F: &[(&str, &str)] = &[
    ("bar", "\\frac{%(num)s}{%(den)s}"), ("skw", "^{%(num)s}/_{%(den)s}"),
    ("noBar", "\\genfrac{}{}{0pt}{}{%(num)s}{%(den)s}"), ("lin", "{%(num)s}/{%(den)s}"),
];

const FUNC: &[(&str, &str)] = &[
    ("sin", "\\sin({fe})"), ("cos", "\\cos({fe})"), ("tan", "\\tan({fe})"), ("arcsin", "\\arcsin({fe})"),
    ("arccos", "\\arccos({fe})"), ("arctan", "\\arctan({fe})"), ("arccot", "\\arccot({fe})"), ("sinh", "\\sinh({fe})"),
    ("cosh", "\\cosh({fe})"), ("tanh", "\\tanh({fe})"), ("coth", "\\coth({fe})"), ("sec", "\\sec({fe})"), ("csc", "\\csc({fe})"),
    ("mod", "\\mod {fe}"), ("max", "\\max({fe})"), ("min", "\\min({fe})"), ("log", "\\log({fe})"), ("ln", "\\ln({fe})"),
    ("exp", "\\exp({fe})"), ("det", "\\det({fe})"), ("gcd", "\\gcd({fe})"), ("deg", "\\deg({fe})"), ("hom", "\\hom({fe})"),
    ("ker", "\\ker({fe})"), ("dim", "\\dim({fe})"), ("arg", "\\arg({fe})"), ("inf", "\\inf({fe})"), ("sup", "\\sup({fe})"),
    ("lim", "\\lim({fe})"), ("Pr", "\\Pr({fe})"),
];

const LIM_FUNC: &[(&str, &str)] = &[
    ("lim", "\\lim_{%(lim)s}"), ("max", "\\max_{%(lim)s}"), ("min", "\\min_{%(lim)s}"),
    ("argmax", "\\operatorname{argmax}_{%(lim)s}"), ("argmin", "\\operatorname{argmin}_{%(lim)s}"),
];

/// `m:oMath` elements of `word/document.xml`, keyed by the body child they sit
/// in: `[block]` for a top-level paragraph, `[block, row, cell, block]` inside a
/// table cell. Equations inside text boxes are left out (their paragraphs are
/// not body blocks). Order within a key is document order.
pub fn collect(document_xml: &str) -> HashMap<Vec<usize>, Vec<String>> {
    let mut out = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(document_xml) else { return out };
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let Some(body) = doc.descendants().find(|n| n.tag_name() == (w, "body").into()) else { return out };
    let is_block = |n: &Node| n.is_element() && matches!(n.tag_name().name(), "p" | "tbl") && n.tag_name().namespace() == Some(w);
    for (bi, blk) in body.children().filter(is_block).enumerate() {
        if blk.tag_name().name() == "p" {
            let v = omaths_of(blk);
            if !v.is_empty() {
                out.insert(vec![bi], v);
            }
            continue;
        }
        for (ri, row) in blk.children().filter(|n| n.tag_name() == (w, "tr").into()).enumerate() {
            for (ci, cell) in row.children().filter(|n| n.tag_name() == (w, "tc").into()).enumerate() {
                for (pi, p) in cell.children().filter(is_block).enumerate() {
                    if p.tag_name().name() == "p" {
                        let v = omaths_of(p);
                        if !v.is_empty() {
                            out.insert(vec![bi, ri, ci, pi], v);
                        }
                    }
                }
            }
        }
    }
    out
}

/// The paragraph's own equations, in order, not descending into text boxes.
fn omaths_of(p: Node) -> Vec<String> {
    let mut v = vec![];
    fn walk(n: Node, v: &mut Vec<String>) {
        for c in n.children().filter(|c| c.is_element()) {
            let name = c.tag_name().name();
            if c.tag_name().namespace() == Some(M) && name == "oMath" {
                v.push(omath_to_latex(c));
            } else if matches!(name, "txbxContent" | "pict" | "Fallback") {
                continue;
            } else {
                walk(c, v);
            }
        }
    }
    walk(p, &mut v);
    v
}
