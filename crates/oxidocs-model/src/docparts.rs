// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Document-wide parts: styles, numbering, settings, web settings, the font
//! table. Styles and numbering are typed because they are what a user
//! creates; settings and the font table are checked bags of named settings.

use serde::{Deserialize, Serialize};

use crate::content::{read_object, Object, PartSink, Resolver};
use crate::props::{is_session_attr, unsupported, Props, Unsupported};
use crate::xml::{Element, Node};

fn attrs(e: &Element) -> Vec<(String, String)> {
    e.attrs
        .iter()
        .filter(|(k, _)| !is_session_attr(k) && !k.starts_with("xmlns") && k != "mc:Ignorable")
        .cloned()
        .collect()
}

fn leaf(e: &Element, path: &str) -> Result<Element, Unsupported> {
    if let Some(c) = e.elements().next() {
        return unsupported(format!("{path}/{}/{}", e.name, c.name));
    }
    Ok(Element { name: e.name.clone(), attrs: attrs(e), children: e.children.clone() })
}

fn rank(order: &[&str], name: &str) -> usize {
    order.iter().position(|n| *n == name).unwrap_or(order.len())
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

pub const STYLE_META: &[&str] = &[
    "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine",
    "w:hidden", "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat",
    "w:locked", "w:personal", "w:personalCompose", "w:personalReply",
];

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Styles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_run: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_paragraph: Option<Props>,
    /// Whether docDefaults was present at all (it may be empty).
    #[serde(default)]
    pub has_defaults: bool,
    /// w:latentStyles: the built-in styles' visibility defaults.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latent: Option<Element>,
    pub styles: Vec<Style>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Style {
    /// w:type, w:styleId, w:default, w:customStyle.
    pub attrs: Vec<(String, String)>,
    /// Name, basedOn, next, link, uiPriority, qFormat ... in schema order.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub meta: Vec<Element>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell: Option<Props>,
    /// Conditional formatting of a table style (first row, banded rows ...).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub table_parts: Vec<Style>,
}

fn read_style_bags(s: &mut Style, c: &Element, here: &str) -> Result<bool, Unsupported> {
    match c.name.as_str() {
        "w:pPr" => s.paragraph = Some(Props::read("w:style", c)?),
        "w:rPr" => s.run = Some(Props::read("w:style", c)?),
        "w:tblPr" => s.table = Some(Props::read("w:style", c)?),
        "w:trPr" => s.row = Some(Props::read("w:style", c)?),
        "w:tcPr" => s.cell = Some(Props::read("w:style", c)?),
        _ => {
            let _ = here;
            return Ok(false);
        }
    }
    Ok(true)
}

impl Styles {
    pub fn read(root: &Element) -> Result<Styles, Unsupported> {
        if root.name != "w:styles" {
            return unsupported(format!("styles/{}", root.name));
        }
        let mut out = Styles::default();
        for c in root.elements() {
            let here = format!("w:styles/{}", c.name);
            match c.name.as_str() {
                "w:docDefaults" => {
                    out.has_defaults = true;
                    for d in c.elements() {
                        match (d.name.as_str(), d.elements().next()) {
                            ("w:rPrDefault", Some(r)) if r.name == "w:rPr" => out.default_run = Some(Props::read("w:rPrDefault", r)?),
                            ("w:pPrDefault", Some(p)) if p.name == "w:pPr" => out.default_paragraph = Some(Props::read("w:pPrDefault", p)?),
                            ("w:rPrDefault", None) => out.default_run = Some(Props::new("w:rPr")),
                            ("w:pPrDefault", None) => out.default_paragraph = Some(Props::new("w:pPr")),
                            _ => return unsupported(format!("{here}/{}", d.name)),
                        }
                    }
                }
                "w:latentStyles" => {
                    for x in c.elements() {
                        if x.name != "w:lsdException" {
                            return unsupported(format!("{here}/{}", x.name));
                        }
                        leaf(x, &here)?;
                    }
                    let mut l = c.clone();
                    l.attrs = attrs(c);
                    out.latent = Some(l);
                }
                "w:style" => {
                    let mut s = Style { attrs: attrs(c), ..Default::default() };
                    for x in c.elements() {
                        let hx = format!("{here}/{}", x.name);
                        if STYLE_META.contains(&x.name.as_str()) {
                            s.meta.push(leaf(x, &here)?);
                        } else if x.name == "w:rsid" {
                            // Editing-session id.
                        } else if read_style_bags(&mut s, x, &hx)? {
                        } else if x.name == "w:tblStylePr" {
                            let mut part = Style { attrs: attrs(x), ..Default::default() };
                            for y in x.elements() {
                                if !read_style_bags(&mut part, y, &hx)? {
                                    return unsupported(format!("{hx}/{}", y.name));
                                }
                            }
                            s.table_parts.push(part);
                        } else {
                            return unsupported(hx);
                        }
                    }
                    s.meta.sort_by_key(|m| rank(STYLE_META, &m.name));
                    out.styles.push(s);
                }
                _ => return unsupported(here),
            }
        }
        Ok(out)
    }

    pub fn write(&self) -> Element {
        let mut root = Element::new("w:styles");
        if self.has_defaults || self.default_run.is_some() || self.default_paragraph.is_some() {
            let mut dd = Element::new("w:docDefaults");
            if let Some(r) = &self.default_run {
                let mut e = Element::new("w:rPrDefault");
                if !r.is_empty() {
                    e.children.push(Node::Element(r.element().clone()));
                }
                dd.children.push(Node::Element(e));
            }
            if let Some(p) = &self.default_paragraph {
                let mut e = Element::new("w:pPrDefault");
                if !p.is_empty() {
                    e.children.push(Node::Element(p.element().clone()));
                }
                dd.children.push(Node::Element(e));
            }
            root.children.push(Node::Element(dd));
        }
        if let Some(l) = &self.latent {
            root.children.push(Node::Element(l.clone()));
        }
        for s in &self.styles {
            root.children.push(Node::Element(write_style(s, "w:style")));
        }
        root
    }
}

fn write_style(s: &Style, name: &str) -> Element {
    let mut e = Element { name: name.into(), attrs: s.attrs.clone(), children: Vec::new() };
    for m in &s.meta {
        e.children.push(Node::Element(m.clone()));
    }
    for p in [&s.paragraph, &s.run, &s.table, &s.row, &s.cell].into_iter().flatten() {
        e.children.push(Node::Element(p.element().clone()));
    }
    for part in &s.table_parts {
        e.children.push(Node::Element(write_style(part, "w:tblStylePr")));
    }
    e
}

// ---------------------------------------------------------------------------
// Numbering
// ---------------------------------------------------------------------------

pub const ABSTRACT_META: &[&str] = &["w:nsid", "w:multiLevelType", "w:tmpl", "w:name", "w:styleLink", "w:numStyleLink"];
pub const LEVEL_META: &[&str] = &[
    "w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff", "w:lvlText",
    "w:lvlPicBulletId", "w:legacy", "w:lvlJc",
];

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Numbering {
    /// Picture bullets: w:numPicBulletId and the picture.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub picture_bullets: Vec<(Vec<(String, String)>, Object)>,
    pub abstracts: Vec<AbstractNum>,
    pub nums: Vec<Num>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mac_cleanup: Option<Element>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct AbstractNum {
    pub attrs: Vec<(String, String)>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub meta: Vec<Element>,
    pub levels: Vec<Level>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Level {
    /// w:ilvl, w:tplc, w:tentative.
    pub attrs: Vec<(String, String)>,
    /// start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlJc ...
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub meta: Vec<Element>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph: Option<Props>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run: Option<Props>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Num {
    /// w:numId (and w16cid:durableId).
    pub attrs: Vec<(String, String)>,
    pub abstract_id: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub overrides: Vec<LevelOverride>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct LevelOverride {
    pub attrs: Vec<(String, String)>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Element>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub level: Option<Level>,
}

fn read_level(e: &Element, path: &str) -> Result<Level, Unsupported> {
    let mut l = Level { attrs: attrs(e), ..Default::default() };
    for c in e.elements() {
        if LEVEL_META.contains(&c.name.as_str()) {
            // A legacy numbering element carries only attributes too.
            l.meta.push(leaf(c, path)?);
        } else if c.name == "w:pPr" {
            l.paragraph = Some(Props::read("w:lvl", c)?);
        } else if c.name == "w:rPr" {
            l.run = Some(Props::read("w:lvl", c)?);
        } else {
            return unsupported(format!("{path}/{}", c.name));
        }
    }
    l.meta.sort_by_key(|m| rank(LEVEL_META, &m.name));
    Ok(l)
}

fn write_level(l: &Level) -> Element {
    let mut e = Element { name: "w:lvl".into(), attrs: l.attrs.clone(), children: Vec::new() };
    for m in &l.meta {
        e.children.push(Node::Element(m.clone()));
    }
    for p in [&l.paragraph, &l.run].into_iter().flatten() {
        e.children.push(Node::Element(p.element().clone()));
    }
    e
}

impl Numbering {
    pub fn read(root: &Element, res: &mut dyn Resolver) -> Result<Numbering, Unsupported> {
        if root.name != "w:numbering" {
            return unsupported(format!("numbering/{}", root.name));
        }
        let mut out = Numbering::default();
        for c in root.elements() {
            let here = format!("w:numbering/{}", c.name);
            match c.name.as_str() {
                "w:numPicBullet" => {
                    let pic = c.elements().next().ok_or_else(|| Unsupported { path: format!("{here}(empty)") })?;
                    if !matches!(pic.name.as_str(), "w:pict" | "w:drawing") {
                        return unsupported(format!("{here}/{}", pic.name));
                    }
                    out.picture_bullets.push((attrs(c), read_object(pic, res)?));
                }
                "w:abstractNum" => {
                    let mut a = AbstractNum { attrs: attrs(c), ..Default::default() };
                    for x in c.elements() {
                        if ABSTRACT_META.contains(&x.name.as_str()) {
                            a.meta.push(leaf(x, &here)?);
                        } else if x.name == "w:lvl" {
                            a.levels.push(read_level(x, &format!("{here}/w:lvl"))?);
                        } else {
                            return unsupported(format!("{here}/{}", x.name));
                        }
                    }
                    a.meta.sort_by_key(|m| rank(ABSTRACT_META, &m.name));
                    out.abstracts.push(a);
                }
                "w:num" => {
                    let mut n = Num { attrs: attrs(c), ..Default::default() };
                    for x in c.elements() {
                        match x.name.as_str() {
                            "w:abstractNumId" => n.abstract_id = x.attr("w:val").unwrap_or("").to_string(),
                            "w:lvlOverride" => {
                                let mut o = LevelOverride { attrs: attrs(x), ..Default::default() };
                                for y in x.elements() {
                                    match y.name.as_str() {
                                        "w:startOverride" => o.start = Some(leaf(y, &here)?),
                                        "w:lvl" => o.level = Some(read_level(y, &format!("{here}/w:lvlOverride/w:lvl"))?),
                                        _ => return unsupported(format!("{here}/w:lvlOverride/{}", y.name)),
                                    }
                                }
                                n.overrides.push(o);
                            }
                            _ => return unsupported(format!("{here}/{}", x.name)),
                        }
                    }
                    out.nums.push(n);
                }
                "w:numIdMacAtCleanup" => out.mac_cleanup = Some(leaf(c, "w:numbering")?),
                _ => return unsupported(here),
            }
        }
        Ok(out)
    }

    pub fn write(&self, sink: &mut dyn PartSink) -> Element {
        let mut root = Element::new("w:numbering");
        for (a, pic) in &self.picture_bullets {
            let mut e = Element { name: "w:numPicBullet".into(), attrs: a.clone(), children: Vec::new() };
            e.children.push(Node::Element(crate::content::write_object(pic, sink)));
            root.children.push(Node::Element(e));
        }
        for a in &self.abstracts {
            let mut e = Element { name: "w:abstractNum".into(), attrs: a.attrs.clone(), children: Vec::new() };
            for m in &a.meta {
                e.children.push(Node::Element(m.clone()));
            }
            for l in &a.levels {
                e.children.push(Node::Element(write_level(l)));
            }
            root.children.push(Node::Element(e));
        }
        for n in &self.nums {
            let mut e = Element { name: "w:num".into(), attrs: n.attrs.clone(), children: Vec::new() };
            e.children.push(Node::Element(Element::new("w:abstractNumId").with_attr("w:val", n.abstract_id.clone())));
            for o in &n.overrides {
                let mut oe = Element { name: "w:lvlOverride".into(), attrs: o.attrs.clone(), children: Vec::new() };
                if let Some(s) = &o.start {
                    oe.children.push(Node::Element(s.clone()));
                }
                if let Some(l) = &o.level {
                    oe.children.push(Node::Element(write_level(l)));
                }
                e.children.push(Node::Element(oe));
            }
            root.children.push(Node::Element(e));
        }
        if let Some(m) = &self.mac_cleanup {
            root.children.push(Node::Element(m.clone()));
        }
        root
    }
}

// ---------------------------------------------------------------------------
// Settings-like bags
// ---------------------------------------------------------------------------

/// CT_Settings (§17.15.1.78) in schema order.
pub const SETTINGS: &[&str] = &[
    "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation",
    "w:removeDateAndTime", "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape",
    "w:printPostScriptOverText", "w:printFractionalCharacterWidth", "w:printFormsData",
    "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts", "w:saveFormsData",
    "w:mirrorMargins", "w:alignBordersAndEdges", "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter", "w:gutterAtTop", "w:hideSpellingErrors",
    "w:hideGrammaticalErrors", "w:activeWritingStyle", "w:proofState", "w:formsDesign",
    "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter", "w:stylePaneSortMethod",
    "w:documentType", "w:mailMerge", "w:revisionView", "w:trackRevisions",
    "w:doNotTrackMoves", "w:doNotTrackFormatting", "w:documentProtection",
    "w:autoFormatOverride", "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop",
    "w:autoHyphenation", "w:consecutiveHyphenLimit", "w:hyphenationZone",
    "w:doNotHyphenateCaps", "w:showEnvelope", "w:summaryLength", "w:clickAndTypeStyle",
    "w:defaultTableStyle", "w:evenAndOddHeaders", "w:bookFoldRevPrinting",
    "w:bookFoldPrinting", "w:bookFoldPrintingSheets", "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing", "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery", "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin", "w:drawingGridVerticalOrigin", "w:doNotShadeFormData",
    "w:noPunctuationKerning", "w:characterSpacingControl", "w:printTwoOnOne",
    "w:strictFirstAndLastChars", "w:noLineBreaksAfter", "w:noLineBreaksBefore",
    "w:savePreviewPicture", "w:doNotValidateAgainstSchema", "w:saveInvalidXml",
    "w:ignoreMixedContent", "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly", "w:useXSLTWhenSaving", "w:saveThroughXslt", "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace", "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr",
    "w:endnotePr", "w:compat", "w:compatSetting", "w:docVars", "w:rsids", "m:mathPr",
    "w:uiCompat97To2003", "w:attachedSchema",
    "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions", "w:readModeInkLockDown",
    "w:smartTagType", "sl:schemaLibrary", "w:shapeDefaults", "w:doNotEmbedSmartTags",
    "w:decimalSymbol", "w:listSeparator",
];

/// The compatibility options of w:compat (§17.15.3). Some generators write
/// them straight under w:settings; they are kept where they were written.
pub const COMPAT_OPTIONS: &[&str] = &[
    "w:useSingleBorderforContiguousCells", "w:wpJustification", "w:noTabHangInd",
    "w:noLeading", "w:spaceForUL", "w:noColumnBalance", "w:balanceSingleByteDoubleByteWidth",
    "w:noExtraLineSpacing", "w:doNotLeaveBackslashAlone", "w:ulTrailSpace",
    "w:doNotExpandShiftReturn", "w:spacingInWholePoints", "w:lineWrapLikeWord6",
    "w:printBodyTextBeforeHeader", "w:printColBlack", "w:wpSpaceWidth", "w:showBreaksInFrames",
    "w:subFontBySize", "w:suppressBottomSpacing", "w:suppressTopSpacing",
    "w:suppressSpacingAtTopOfPage", "w:suppressTopSpacingWP", "w:suppressSpBfAfterPgBrk",
    "w:swapBordersFacingPages", "w:convMailMergeEsc", "w:truncateFontHeightsLikeWP6",
    "w:mwSmallCaps", "w:usePrinterMetrics", "w:doNotSuppressParagraphBorders",
    "w:wrapTrailSpaces", "w:footnoteLayoutLikeWW8", "w:shapeLayoutLikeWW8",
    "w:alignTablesRowByRow", "w:forgetLastTabAlignment", "w:adjustLineHeightInTable",
    "w:autoSpaceLikeWord95", "w:noSpaceRaiseLower", "w:doNotUseHTMLParagraphAutoSpacing",
    "w:layoutRawTableWidth", "w:layoutTableRowsApart", "w:useWord97LineBreakRules",
    "w:doNotBreakWrappedTables", "w:doNotSnapToGridInCell", "w:selectFldWithFirstOrLastChar",
    "w:applyBreakingRules", "w:doNotWrapTextWithPunct", "w:doNotUseEastAsianBreakRules",
    "w:useWord2002TableStyleRules", "w:growAutofit", "w:useFELayout", "w:useNormalStyleForList",
    "w:doNotUseIndentAsNumberingTabStop", "w:useAltKinsokuLineBreakRules",
    "w:allowSpaceOfSameStyleInTable", "w:doNotSuppressIndentation",
    "w:doNotAutofitConstrainedTables", "w:autofitToFirstFixedWidthCell",
    "w:underlineTabInNumList", "w:displayHangulFixedWidth", "w:splitPgBreakAndParaMark",
    "w:doNotVertAlignCellWithSp", "w:doNotBreakConstrainedForcedTable",
    "w:doNotVertAlignInTxbx", "w:useAnsiKerningPairs", "w:cachedColBalance",
];

/// CT_WebSettings (§17.15.2.44).
pub const WEB_SETTINGS: &[&str] = &[
    "w:frameset", "w:divs", "w:encoding", "w:optimizeForBrowser", "w:relyOnVML",
    "w:allowPNG", "w:doNotRelyOnCSS", "w:doNotSaveAsSingleFile", "w:doNotOrganizeInFolder",
    "w:doNotUseLongFileNames", "w:pixelsPerInch", "w:targetScreenSz", "w:saveSmartTagsAsXml",
];

pub const FONT: &[&str] = &[
    "w:altName", "w:panose1", "w:charset", "w:family", "w:notTrueType", "w:pitch", "w:sig",
    "w:embedRegular", "w:embedBold", "w:embedItalic", "w:embedBoldItalic",
];

/// A settings-like part: a root whose children are named settings in a fixed
/// order. Each setting's own subtree (compatibility options, document
/// variables, shape defaults) is kept whole.
pub fn read_settings_bag(
    root: &Element,
    root_name: &str,
    order: &[&str],
    res: &mut dyn Resolver,
) -> Result<Object, Unsupported> {
    if root.name != root_name {
        return unsupported(format!("{root_name}/{}", root.name));
    }
    let mut kept = Element { name: root.name.clone(), attrs: attrs(root), children: Vec::new() };
    for c in root.elements() {
        if c.name == "w:rsids" {
            continue; // editing-session ids
        }
        let misplaced_compat = root_name == "w:settings" && COMPAT_OPTIONS.contains(&c.name.as_str());
        if !order.contains(&c.name.as_str()) && !crate::schema::is_extension(&c.name) && !misplaced_compat {
            return unsupported(format!("{root_name}/{}", c.name));
        }
        kept.children.push(Node::Element(c.clone()));
    }
    kept.children.sort_by_key(|n| match n {
        Node::Element(c) => rank(order, &c.name),
        Node::Text(_) => usize::MAX,
    });
    read_object(&kept, res)
}

/// The font table: one w:font per face, each a list of named facts.
pub fn read_font_table(root: &Element, res: &mut dyn Resolver) -> Result<Object, Unsupported> {
    if root.name != "w:fonts" {
        return unsupported(format!("fontTable/{}", root.name));
    }
    let mut kept = Element { name: root.name.clone(), attrs: attrs(root), children: Vec::new() };
    for f in root.elements() {
        if f.name != "w:font" {
            return unsupported(format!("w:fonts/{}", f.name));
        }
        let mut font = Element { name: f.name.clone(), attrs: attrs(f), children: Vec::new() };
        for x in f.elements() {
            if !FONT.contains(&x.name.as_str()) {
                return unsupported(format!("w:fonts/w:font/{}", x.name));
            }
            font.children.push(Node::Element(leaf(x, "w:fonts/w:font")?));
        }
        font.children.sort_by_key(|n| match n {
            Node::Element(c) => rank(FONT, &c.name),
            Node::Text(_) => usize::MAX,
        });
        kept.children.push(Node::Element(font));
    }
    read_object(&kept, res)
}

/// A part that is a flat list of entries of known names, each kept whole
/// (comment thread state, durable comment ids).
pub fn read_list(root: &Element, root_name: &str, items: &[&str], res: &mut dyn Resolver) -> Result<Object, Unsupported> {
    if root.name != root_name {
        return unsupported(format!("{root_name}/{}", root.name));
    }
    for c in root.elements() {
        if !items.contains(&c.name.as_str()) {
            return unsupported(format!("{root_name}/{}", c.name));
        }
    }
    let mut kept = root.clone();
    kept.attrs = attrs(root);
    read_object(&kept, res)
}

/// The people part (comment authors): w15:person entries.
pub fn read_people(root: &Element, res: &mut dyn Resolver) -> Result<Object, Unsupported> {
    if root.name != "w15:people" {
        return unsupported(format!("people/{}", root.name));
    }
    for p in root.elements() {
        if p.name != "w15:person" {
            return unsupported(format!("w15:people/{}", p.name));
        }
        for x in p.elements() {
            if x.name != "w15:presenceInfo" {
                return unsupported(format!("w15:people/w15:person/{}", x.name));
            }
        }
    }
    let mut kept = root.clone();
    kept.attrs = attrs(root);
    read_object(&kept, res)
}
