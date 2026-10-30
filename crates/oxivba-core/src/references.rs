// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Host-supplied type information for static compilation, without COM or I/O.
use std::collections::{BTreeMap, BTreeSet};

#[derive(Debug, Clone)]
pub(crate) struct ReferenceSignature {
    pub result: String,
    pub parameters: Vec<(String, bool)>,
    pub variadic: bool,
    pub invocation: u8,
    /// FUNCDESC function kind: 4 is IDispatch, 3 is a static library function.
    pub function_kind: u8,
}

#[derive(Debug, Clone, Default)]
struct Interface {
    closed: bool,
    members: BTreeMap<String, Vec<ReferenceSignature>>,
    default: Option<String>,
}

/// A case-insensitive snapshot of reference-library members and signatures.
/// The tab-separated representation supports interfaces (I), type aliases (A),
/// members (M), default members (D), functions (F), and host globals (G).
#[derive(Debug, Clone, Default)]
pub struct CompileReferenceLibrary {
    interfaces: BTreeMap<String, Interface>,
    aliases: BTreeMap<String, String>,
    functions: BTreeMap<String, ReferenceSignature>,
    globals: BTreeMap<String, (String, bool)>,
}

impl CompileReferenceLibrary {
    /// Declared names of a reference interface, resolving its type aliases.
    pub fn member_names<'a>(&'a self, owner: &str) -> impl Iterator<Item = &'a str> {
        let interface = self.resolve_type(owner).and_then(|(name, _)| self.interfaces.get(&name));
        interface.into_iter().flat_map(|interface| interface.members.keys().map(String::as_str))
    }

    /// Invocation kinds declared by the referenced interface, after aliases.
    pub fn member_invocation_flags(&self, owner: &str, name: &str) -> Option<u8> {
        let (owner, _) = self.resolve_type(owner)?;
        let signatures = self.interfaces.get(&owner)?.members.get(&name.to_ascii_lowercase())?;
        Some(signatures.iter().fold(0, |flags, signature| flags | signature.invocation))
    }

    /// Parameter names of a referenced invocation, in declaration order.
    pub fn member_parameter_names(&self, owner: &str, name: &str, invocation: u8) -> Option<Vec<String>> {
        let (owner, _) = self.resolve_type(owner)?;
        let signature = self.interfaces.get(&owner)?.members.get(&name.to_ascii_lowercase())?
            .iter().find(|signature| signature.invocation & invocation != 0)?;
        Some(signature.parameters.iter().map(|(name, _)| name.clone()).collect())
    }

    /// The static interface of a host global, when it has one.
    pub fn global_type_name(&self, name: &str) -> Option<String> {
        self.global(name).map(|(name, _)| name)
    }

    /// The declared result of a referenced property getter or method.
    pub fn member_result_type(&self, owner: &str, name: &str) -> Option<String> {
        let (owner, _) = self.resolve_type(owner)?;
        self.member(&owner, name).map(|signature| signature.result.clone())
    }

    pub fn from_tsv(source: &str) -> Result<Self, String> {
        fn signature(
            result: &str,
            variadic: &str,
            params: &str,
            invocation: u8,
            function_kind: u8,
        ) -> ReferenceSignature {
            ReferenceSignature {
                result: result.to_ascii_lowercase(),
                parameters: params
                    .split(';')
                    .filter(|p| !p.is_empty())
                    .map(|p| {
                        (
                            p.trim_end_matches('?').to_ascii_lowercase(),
                            p.ends_with('?'),
                        )
                    })
                    .collect(),
                variadic: variadic == "1",
                invocation,
                function_kind,
            }
        }
        let mut library = Self::default();
        for (index, line) in source.lines().enumerate() {
            if line.is_empty() || line.starts_with('#') {
                continue;
            }
            let mut fields: Vec<_> = line.split('\t').collect();
            // Snapshots without function-kind columns retain their original
            // interface-member/static-function distinction.
            if fields.first() == Some(&"M") && fields.len() == 7 {
                fields.insert(4, "4");
            }
            if fields.first() == Some(&"F") && fields.len() == 6 {
                fields.insert(3, "3");
            }
            match fields.as_slice() {
                ["I", name, closed] => {
                    library
                        .interfaces
                        .entry(name.to_ascii_lowercase())
                        .or_default()
                        .closed = *closed == "1"
                }
                ["A", name, target] => {
                    library
                        .aliases
                        .insert(name.to_ascii_lowercase(), target.to_ascii_lowercase());
                }
                ["M", owner, name, invocation, function_kind, result, variadic, params] => {
                    let invocation = invocation.parse().map_err(|_| {
                        format!("Invalid reference invocation at line {}", index + 1)
                    })?;
                    let function_kind = function_kind.parse().map_err(|_| {
                        format!("Invalid reference function kind at line {}", index + 1)
                    })?;
                    library
                        .interfaces
                        .entry(owner.to_ascii_lowercase())
                        .or_default()
                        .members
                        .entry(name.to_ascii_lowercase())
                        .or_default()
                        .push(signature(
                            result,
                            variadic,
                            params,
                            invocation,
                            function_kind,
                        ));
                }
                ["D", owner, name] => {
                    library
                        .interfaces
                        .entry(owner.to_ascii_lowercase())
                        .or_default()
                        .default = Some(name.to_ascii_lowercase())
                }
                ["F", name, invocation, function_kind, result, variadic, params] => {
                    let invocation = invocation.parse().map_err(|_| {
                        format!("Invalid reference invocation at line {}", index + 1)
                    })?;
                    let function_kind = function_kind.parse().map_err(|_| {
                        format!("Invalid reference function kind at line {}", index + 1)
                    })?;
                    library.functions.insert(
                        name.to_ascii_lowercase(),
                        signature(result, variadic, params, invocation, function_kind),
                    );
                }
                ["G", name, target, closed] => {
                    library.globals.insert(
                        name.to_ascii_lowercase(),
                        (target.to_ascii_lowercase(), *closed == "1"),
                    );
                }
                _ => return Err(format!("Invalid reference record at line {}", index + 1)),
            }
        }
        Ok(library)
    }

    /// Scalar storage behind imported aliases, without treating every
    /// numeric declaration as Long. Object aliases retain their interfaces.
    pub(crate) fn resolve_scalar_type(&self, name: &str) -> Option<String> {
        let mut name = name.to_ascii_lowercase();
        let mut visited = BTreeSet::new();
        loop {
            if !visited.insert(name.clone()) {
                return None;
            }
            name = self.aliases.get(&name)?.clone();
            if matches!(name.as_str(),
                "byte" | "integer" | "long" | "longlong" | "longptr" |
                "single" | "double" | "currency" | "date" | "boolean" | "string")
            {
                return Some(name);
            }
            if self.interfaces.contains_key(&name) {
                return None;
            }
        }
    }

    pub(crate) fn resolve_type(&self, name: &str) -> Option<(String, bool)> {
        let mut name = name.to_ascii_lowercase();
        let mut visited = BTreeSet::new();
        loop {
            if let Some(interface) = self.interfaces.get(&name) {
                return Some((name, interface.closed));
            }
            if !visited.insert(name.clone()) {
                return None;
            }
            name = self.aliases.get(&name)?.clone();
        }
    }

    pub(crate) fn global(&self, name: &str) -> Option<(String, bool)> {
        let (name, closed) = self.globals.get(&name.to_ascii_lowercase())?;
        let (name, intrinsic_closed) = self.resolve_type(name)?;
        Some((name, *closed || intrinsic_closed))
    }

    pub(crate) fn has_member(&self, owner: &str, name: &str) -> bool {
        self.interfaces
            .get(owner)
            .is_some_and(|i| i.members.contains_key(&name.to_ascii_lowercase()))
    }

    pub(crate) fn member(&self, owner: &str, name: &str) -> Option<&ReferenceSignature> {
        self.interfaces
            .get(owner)?
            .members
            .get(&name.to_ascii_lowercase())?
            .iter()
            .find(|s| s.invocation & 3 != 0)
    }

    pub(crate) fn member_writable(&self, owner: &str, name: &str) -> bool {
        self.interfaces
            .get(owner)
            .and_then(|interface| interface.members.get(&name.to_ascii_lowercase()))
            .is_some_and(|signatures| {
                signatures
                    .iter()
                    .any(|signature| signature.invocation & 12 != 0)
            })
    }

    pub(crate) fn default_member(&self, owner: &str) -> Option<&ReferenceSignature> {
        let interface = self.interfaces.get(owner)?;
        self.member(owner, interface.default.as_ref()?)
    }

    pub(crate) fn function(&self, name: &str) -> Option<&ReferenceSignature> {
        self.functions.get(&name.to_ascii_lowercase())
    }
}
