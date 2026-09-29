// SPDX-License-Identifier: MIT OR Apache-2.0

//! `CreateObject("VBScript.RegExp")`: the RegExp object, its
//! MatchCollection, Match and SubMatches, on the regular-expression engine
//! the worksheet's REGEX functions use. Measured (r79_regexp.vba): a fresh
//! RegExp has an empty Pattern and Global, IgnoreCase and MultiLine False;
//! Execute without Global finds one match; an unmatched group reads Empty;
//! Replace takes $1, $&, $`, $' and $$; a pattern that will not compile is
//! refused only when used, as error 5020.

use super::*;

/// One RegExp object's settings.
#[derive(Debug, Clone, Default)]
pub(super) struct RegExpState {
    pattern: String,
    global: bool,
    ignore_case: bool,
    multiline: bool,
}

/// One match an Execute found.
#[derive(Debug, Clone)]
pub(super) struct RegExpHit {
    value: String,
    first_index: i64,
    length: i64,
    groups: Vec<Option<String>>,
}

impl<'a> WorkbookHost<'a> {
    /// `CreateObject` for the classes this host makes itself.
    pub(super) fn create_host_object(&mut self, class: &str) -> Option<Value> {
        if !class.eq_ignore_ascii_case("vbscript.regexp") {
            return None;
        }
        let index = self.regexps.len();
        self.regexps.push(RegExpState::default());
        Some(self.object(HostObject::RegExp(index)))
    }

    fn compiled(&self, index: usize) -> Result<oxicells_calc::regex::Regex, String> {
        let state = &self.regexps[index];
        oxicells_calc::regex::Regex::new(&state.pattern, state.ignore_case, state.multiline).map_err(|_| {
            oxivba_core::host_error_from(5020, "VBAProject", "Application-defined or object-defined error".to_string())
        })
    }

    fn regexp_text(value: Option<&Value>) -> String {
        match value {
            Some(Value::String(text)) => text.clone(),
            Some(Value::Empty) | None => String::new(),
            Some(other) => format_debug_value(other),
        }
    }

    /// The members of a RegExp and of what it hands back.
    pub(super) fn regexp_member(&mut self, object: HostObject, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match object {
            HostObject::RegExp(index) => {
                let state = &self.regexps[index];
                Ok(Some(match lower.as_str() {
                    "pattern" => Value::String(state.pattern.clone()),
                    "global" => Value::Boolean(state.global),
                    "ignorecase" => Value::Boolean(state.ignore_case),
                    "multiline" => Value::Boolean(state.multiline),
                    "test" => {
                        let regex = self.compiled(index)?;
                        let text: Vec<char> = Self::regexp_text(args.first()).chars().collect();
                        Value::Boolean(regex.find_at(&text, 0).is_some())
                    }
                    "execute" => {
                        let regex = self.compiled(index)?;
                        let text: Vec<char> = Self::regexp_text(args.first()).chars().collect();
                        let found = if self.regexps[index].global {
                            regex.find_all(&text)
                        } else {
                            regex.find_at(&text, 0).into_iter().collect()
                        };
                        let piece = |span: Option<(usize, usize)>| span.map(|(s, e)| text[s..e].iter().collect::<String>());
                        let hits: Vec<RegExpHit> = found
                            .iter()
                            .map(|caps| {
                                let (s, e) = caps[0].unwrap_or((0, 0));
                                RegExpHit {
                                    value: text[s..e].iter().collect(),
                                    first_index: s as i64,
                                    length: (e - s) as i64,
                                    groups: caps[1..].iter().map(|c| piece(*c)).collect(),
                                }
                            })
                            .collect();
                        let set = self.regexp_hits.len();
                        self.regexp_hits.push(hits);
                        self.object(HostObject::RegExpMatches(set))
                    }
                    "replace" => {
                        let regex = self.compiled(index)?;
                        let text: Vec<char> = Self::regexp_text(args.first()).chars().collect();
                        let replacement = Self::regexp_text(args.get(1));
                        let found = if self.regexps[index].global {
                            regex.find_all(&text)
                        } else {
                            regex.find_at(&text, 0).into_iter().collect()
                        };
                        let mut out = String::new();
                        let mut last = 0;
                        for caps in &found {
                            let (s, e) = caps[0].unwrap_or((last, last));
                            out.extend(&text[last..s]);
                            out.push_str(&vbscript_expand(&replacement, &text, caps));
                            last = e;
                        }
                        out.extend(&text[last..]);
                        Value::String(out)
                    }
                    _ => return Err(host_error(438, &format!("RegExp has no member {name}"))),
                }))
            }
            HostObject::RegExpMatches(set) => {
                let count = self.regexp_hits[set].len();
                match lower.as_str() {
                    "count" => Ok(Some(Value::Integer(count as i64))),
                    "item" | "_default" => {
                        let at = args.first().and_then(any_whole_number).unwrap_or(-1);
                        if at < 0 || at as usize >= count {
                            return Err(host_error(5, "Invalid procedure call or argument"));
                        }
                        Ok(Some(self.object(HostObject::RegExpMatch(set, at as usize))))
                    }
                    _ => Err(host_error(438, &format!("MatchCollection has no member {name}"))),
                }
            }
            HostObject::RegExpMatch(set, at) => {
                let hit = &self.regexp_hits[set][at];
                match lower.as_str() {
                    "value" | "_default" => Ok(Some(Value::String(hit.value.clone()))),
                    "firstindex" => Ok(Some(Value::Integer(hit.first_index))),
                    "length" => Ok(Some(Value::Integer(hit.length))),
                    // `m.SubMatches(0)` names the group straight away.
                    "submatches" if !args.is_empty() => self.regexp_member(HostObject::RegExpSubMatches(set, at), "Item", args),
                    "submatches" => Ok(Some(self.object(HostObject::RegExpSubMatches(set, at)))),
                    _ => Err(host_error(438, &format!("Match has no member {name}"))),
                }
            }
            HostObject::RegExpSubMatches(set, at) => {
                let groups = &self.regexp_hits[set][at].groups;
                match lower.as_str() {
                    "count" => Ok(Some(Value::Integer(groups.len() as i64))),
                    "item" | "_default" | "value" => {
                        let Some(index) = args.first().and_then(any_whole_number) else {
                            return Err(host_error(450, "Wrong number of arguments or invalid property assignment"));
                        };
                        match groups.get(index.max(0) as usize).filter(|_| index >= 0) {
                            Some(Some(text)) => Ok(Some(Value::String(text.clone()))),
                            Some(None) => Ok(Some(Value::Empty)),
                            None => Err(host_error(5, "Invalid procedure call or argument")),
                        }
                    }
                    _ => Err(host_error(438, &format!("SubMatches has no member {name}"))),
                }
            }
            _ => Ok(None),
        }
    }

    pub(super) fn set_regexp_member(&mut self, index: usize, name: &str, value: &Value) -> Result<bool, String> {
        let truth = |value: &Value| match value {
            Value::Boolean(state) => *state,
            other => any_number(other).is_some_and(|n| n != 0.0),
        };
        let state = &mut self.regexps[index];
        match name.to_ascii_lowercase().as_str() {
            "pattern" => state.pattern = Self::regexp_text(Some(value)),
            "global" => state.global = truth(value),
            "ignorecase" => state.ignore_case = truth(value),
            "multiline" => state.multiline = truth(value),
            _ => return Err(host_error(438, &format!("RegExp has no member {name}"))),
        }
        Ok(true)
    }

    /// For Each over a MatchCollection or SubMatches.
    pub(super) fn regexp_items(&mut self, object: HostObject) -> Option<Vec<Value>> {
        match object {
            HostObject::RegExpMatches(set) => {
                let count = self.regexp_hits[set].len();
                Some((0..count).map(|at| self.object(HostObject::RegExpMatch(set, at))).collect())
            }
            HostObject::RegExpSubMatches(set, at) => Some(
                self.regexp_hits[set][at]
                    .groups
                    .iter()
                    .map(|group| group.clone().map(Value::String).unwrap_or(Value::Empty))
                    .collect(),
            ),
            _ => None,
        }
    }
}

/// VBScript's replacement text: $1..$99, $& the match, $` before it, $'
/// after it, $$ a dollar.
fn vbscript_expand(replacement: &str, text: &[char], caps: &oxicells_calc::regex::Captures) -> String {
    let r: Vec<char> = replacement.chars().collect();
    let (start, end) = caps[0].unwrap_or((0, 0));
    let mut out = String::new();
    let mut i = 0;
    while i < r.len() {
        if r[i] == '$' && i + 1 < r.len() {
            match r[i + 1] {
                '$' => {
                    out.push('$');
                    i += 2;
                    continue;
                }
                '&' => {
                    out.extend(&text[start..end]);
                    i += 2;
                    continue;
                }
                '`' => {
                    out.extend(&text[..start]);
                    i += 2;
                    continue;
                }
                '\'' => {
                    out.extend(&text[end..]);
                    i += 2;
                    continue;
                }
                d if d.is_ascii_digit() && d != '0' => {
                    let mut n = d as usize - 48;
                    let mut j = i + 2;
                    if j < r.len() && r[j].is_ascii_digit() && n * 10 + (r[j] as usize - 48) < caps.len() {
                        n = n * 10 + (r[j] as usize - 48);
                        j += 1;
                    }
                    if n < caps.len() {
                        if let Some((s, e)) = caps[n] {
                            out.extend(&text[s..e]);
                        }
                        i = j;
                        continue;
                    }
                }
                _ => {}
            }
        }
        out.push(r[i]);
        i += 1;
    }
    out
}
