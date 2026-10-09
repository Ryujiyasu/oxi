// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use crate::ir::StyleSheet;
use std::collections::HashMap;

/// Application language for built-in style names and field diagnostics.
/// Set this explicitly when the application's language differs from the host.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldLanguage {
    English,
    Japanese,
}

impl Default for FieldLanguage {
    fn default() -> Self {
        #[cfg(target_os = "windows")]
        {
            #[link(name = "kernel32")]
            extern "system" {
                fn GetUserDefaultUILanguage() -> u16;
            }
            // The Win32 call has no arguments and does not access Rust memory.
            if unsafe { GetUserDefaultUILanguage() } & 0x03ff == 0x11 {
                return Self::Japanese;
            }
        }
        Self::English
    }
}

#[derive(Default)]
struct Bindings {
    // None marks an ambiguous name; never depend on HashMap iteration order.
    names: HashMap<String, Option<String>>,
    language: FieldLanguage,
}

impl Bindings {
    fn add(&mut self, name: &str, id: &str) {
        let key = name.trim().to_lowercase();
        self.names.entry(key).and_modify(|value| {
            if value.as_deref() != Some(id) { *value = None; }
        }).or_insert_with(|| Some(id.to_owned()));
    }

    fn heading_name(&self, number: u8) -> String {
        match self.language {
            FieldLanguage::English => format!("Heading {number}"),
            FieldLanguage::Japanese => format!("\u{898b}\u{51fa}\u{3057} {number}"),
        }
    }

    fn lookup(&self, name: &str) -> Option<&str> {
        let key = name.trim().to_lowercase();
        if let Some(value) = self.names.get(&key) { return value.as_deref(); }
        // Word accepts a name followed by any subset of its aliases. All
        // components must identify the same style, rather than accepting an
        // arbitrary suffix after a valid primary name.
        let mut parts = key.split(',');
        let first = self.names.get(parts.next()?)?.as_deref()?;
        for part in parts {
            if self.names.get(part.trim())?.as_deref()? != first { return None; }
        }
        Some(first)
    }

    fn unknown(&self, name: &str) -> String {
        match self.language {
            FieldLanguage::English => format!("Error! Use the Home tab to apply {name} to the text that you want to appear here."),
            FieldLanguage::Japanese => format!("\u{30a8}\u{30e9}\u{30fc}! [\u{30db}\u{30fc}\u{30e0}] \u{30bf}\u{30d6}\u{3092}\u{4f7f}\u{7528}\u{3057}\u{3066}\u{3001}\u{3053}\u{3053}\u{306b}\u{8868}\u{793a}\u{3059}\u{308b}\u{6587}\u{5b57}\u{5217}\u{306b} {name} \u{3092}\u{9069}\u{7528}\u{3057}\u{3066}\u{304f}\u{3060}\u{3055}\u{3044}\u{3002}"),
        }
    }

    fn unused(&self) -> String {
        match self.language {
            FieldLanguage::English => "Error! No text of specified style in document.".to_owned(),
            FieldLanguage::Japanese => "\u{30a8}\u{30e9}\u{30fc}! \u{6307}\u{5b9a}\u{3057}\u{305f}\u{30b9}\u{30bf}\u{30a4}\u{30eb}\u{306f}\u{4f7f}\u{308f}\u{308c}\u{3066}\u{3044}\u{307e}\u{305b}\u{3093}\u{3002}".to_owned(),
        }
    }
}

thread_local! {
    static CURRENT: std::cell::RefCell<Bindings> = std::cell::RefCell::new(Bindings::default());
}

pub(super) fn initialize(styles: &StyleSheet, language: FieldLanguage) {
    let mut bindings = Bindings { language, ..Bindings::default() };
    let mut headings = HashMap::new();
    for style in styles.styles.values() {
        let name = style.display_name.as_deref().unwrap_or(&style.style_id);
        let lower = name.to_lowercase();
        let heading = if style.is_custom { None } else {
            lower.strip_prefix("heading ")
                .or_else(|| lower.strip_prefix("\u{898b}\u{51fa}\u{3057} "))
                .and_then(|n| n.parse::<u8>().ok()).filter(|n| (1..=9).contains(n))
        };
        if let Some(number) = heading {
            headings.insert(number, style.style_id.clone());
        } else if !style.is_custom && (lower == "normal" || lower == "\u{6a19}\u{6e96}") {
            let localized = match language {
                FieldLanguage::English => "Normal",
                FieldLanguage::Japanese => "\u{6a19}\u{6e96}",
            };
            bindings.add(localized, &style.style_id);
        } else {
            bindings.add(name, &style.style_id);
        }
        for alias in &style.aliases { bindings.add(alias, &style.style_id); }
    }
    // Built-in headings exist even when a package does not serialize their
    // unused definitions. Numeric references denote heading levels 1..9.
    for number in 1..=9 {
        let id = headings.get(&number).cloned().unwrap_or_else(|| format!("Heading{number}"));
        bindings.add(&number.to_string(), &id);
        bindings.add(&bindings.heading_name(number), &id);
    }
    CURRENT.with(|current| *current.borrow_mut() = bindings);
}

pub(super) fn style_id(name: &str) -> Option<String> {
    CURRENT.with(|current| current.borrow().lookup(name).map(str::to_owned))
}

pub(super) fn resolve(name: &str, occurrences: &HashMap<String, String>) -> String {
    CURRENT.with(|current| {
        let bindings = current.borrow();
        match bindings.lookup(name) {
            None => bindings.unknown(name),
            Some(id) => occurrences.get(id).map(|text| text.trim().to_owned())
                .unwrap_or_else(|| bindings.unused()),
        }
    })
}
