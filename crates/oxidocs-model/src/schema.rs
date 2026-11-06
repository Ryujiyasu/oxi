// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Child order of every WordprocessingML property container (ECMA-376 Part 1,
//! transitional). A property bag accepts exactly these children and writes
//! them in this order, which is the order Word itself requires.

/// CT_PPrBase (§17.3.1.26), then CT_PPr's own tail.
pub const PPR: &[&str] = &[
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
];

/// The same base without the paragraph-only tail, as styles and
/// numbering levels and pPrChange use it.
pub const PPR_BASE: &[&str] = &[
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:pPrChange",
];

/// EG_RPrBase (§17.3.2.28) and CT_RPr's rPrChange.
pub const RPR: &[&str] = &[
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
    "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
    "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange",
];

/// CT_ParaRPr: the paragraph mark's run properties.
pub const PARA_RPR: &[&str] = &[
    "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rStyle", "w:rFonts", "w:b",
    "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
    "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color",
    "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs",
    "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
    "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
    "w:specVanish", "w:oMath", "w:rPrChange",
];

/// CT_TblPr (§17.4.60).
pub const TBLPR: &[&str] = &[
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange",
];

/// CT_TblPrEx (§17.4.61).
pub const TBLPREX: &[&str] = &[
    "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
    "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblPrExChange",
];

/// CT_TrPr (§17.4.82). The schema allows any order; this is Word's.
pub const TRPR: &[&str] = &[
    "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
    "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader",
    "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
];

/// CT_TcPr (§17.4.70).
pub const TCPR: &[&str] = &[
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
    "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
    "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns",
    "w:cellDel", "w:cellMerge", "w:tcPrChange",
];

/// CT_SectPr (§17.6.17).
pub const SECTPR: &[&str] = &[
    "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr",
    "w:type", "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders",
    "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign",
    "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
    "w:docGrid", "w:printerSettings", "w:sectPrChange",
];

const BORDERS_P: &[&str] = &["w:top", "w:left", "w:start", "w:bottom", "w:right", "w:end", "w:between", "w:bar"];
const BORDERS_TBL: &[&str] = &["w:top", "w:left", "w:start", "w:bottom", "w:right", "w:end", "w:insideH", "w:insideV"];
const BORDERS_TC: &[&str] = &["w:top", "w:left", "w:start", "w:bottom", "w:right", "w:end", "w:insideH", "w:insideV", "w:tl2br", "w:tr2bl"];
const MARGINS: &[&str] = &["w:top", "w:left", "w:start", "w:bottom", "w:right", "w:end"];
const PG_BORDERS: &[&str] = &["w:top", "w:left", "w:bottom", "w:right"];
const NOTE_PR: &[&str] = &["w:pos", "w:numFmt", "w:numStart", "w:numRestart"];

/// The children a nested property element may hold, by (container, element).
/// None means the element is a leaf that carries only attributes.
pub fn nested(container: &str, name: &str) -> Option<&'static [&'static str]> {
    Some(match (container, name) {
        ("w:pPr", "w:rPr") => PARA_RPR,
        ("w:pPr", "w:sectPr") => SECTPR,
        (_, "w:pBdr") => BORDERS_P,
        ("w:tblPr" | "w:tblPrEx", "w:tblBorders") => BORDERS_TBL,
        ("w:tcPr", "w:tcBorders") => BORDERS_TC,
        ("w:tblPr" | "w:tblPrEx", "w:tblCellMar") | ("w:tcPr", "w:tcMar") => MARGINS,
        (_, "w:tabs") => &["w:tab"],
        (_, "w:numPr") => &["w:ilvl", "w:numId", "w:numberingChange", "w:ins"],
        ("w:sectPr", "w:cols") => &["w:col"],
        ("w:sectPr", "w:pgBorders") => PG_BORDERS,
        ("w:sectPr", "w:footnotePr" | "w:endnotePr") => NOTE_PR,
        (_, "w:rPrChange") => &["w:rPr"],
        (_, "w:pPrChange") => &["w:pPr"],
        (_, "w:sectPrChange") => &["w:sectPr"],
        (_, "w:tblPrChange") => &["w:tblPr"],
        (_, "w:tblPrExChange") => &["w:tblPrEx"],
        (_, "w:trPrChange") => &["w:trPr"],
        (_, "w:tcPrChange") => &["w:tcPr"],
        _ => return None,
    })
}

/// The order list a container named `name` uses when found inside `parent`.
pub fn container_order(parent: &str, name: &str) -> Option<&'static [&'static str]> {
    Some(match (parent, name) {
        ("w:pPrChange", "w:pPr") => PPR_BASE,
        ("w:rPrChange", "w:rPr") => RPR,
        ("w:pPr", "w:rPr") => PARA_RPR,
        (_, "w:pPr") => PPR,
        (_, "w:rPr") => RPR,
        (_, "w:tblPr") => TBLPR,
        (_, "w:tblPrEx") => TBLPREX,
        (_, "w:trPr") => TRPR,
        (_, "w:tcPr") => TCPR,
        (_, "w:sectPr") => SECTPR,
        _ => return None,
    })
}

/// Extension namespaces Word appends after the schema children of a run or
/// paragraph property bag (text effects, ligatures, ...). Kept whole.
pub fn is_extension(name: &str) -> bool {
    ["w14:", "w15:", "w16:", "w16se:", "w16cid:", "w16du:", "w16sdtdh:", "w16cex:"]
        .iter()
        .any(|p| name.starts_with(p))
}
