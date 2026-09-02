// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A seal on a macro: someone read it, and vouched for this exact code.
//!
//! The point of a hanko here is to let a safety read go **quiet**. A report
//! that fires on every workbook stops being read, so once a person inside the
//! organisation has looked at a macro and sealed it, the next reader should
//! see nothing — until the macro changes, when the seal must break and the
//! report must come back.
//!
//! # What the seal is over, and why it matters
//!
//! The seal is over the **exact source text**, hashed with SHA-256.
//!
//! It is emphatically NOT over
//! [`fingerprint_module`](oxivba_core::fingerprint::fingerprint_module). That
//! fingerprint deliberately throws information away so that near-copies can be
//! matched: at `Strength::Standard` it ignores local variable names, and at
//! `Loose` it ignores literal values as well. Sealing a fingerprint would mean
//! a seal that survives someone editing the URL a macro fetches from, which is
//! the one edit a seal most needs to catch.
//!
//! The fingerprint still earns its place, on the other side of the question:
//! when a seal breaks, `compare` says *what* changed, so a reviewer reads the
//! three procedures that moved rather than the whole project again.
//!
//! # What a seal is worth without a signature
//!
//! A digest on its own proves nothing about who wrote it: anyone who can edit
//! the workbook can edit an unsigned attestation beside it. So an unsigned
//! match is its own answer, [`Clearance::MatchesButUnsigned`]; so is a signed
//! seal nobody checked, [`Clearance::MatchesButNobodyChecked`]. Reaching
//! [`Clearance::Cleared`] requires a [`SignatureVerifier`] that actually
//! checked the signature. The type is the guard rail — no caller can treat a
//! file someone dropped in a folder as a clearance by accident.

use std::fmt::Write as _;

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

/// The digest of one project's macros: each module, and the whole.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct MacroDigest {
    /// `(module name, hex digest)`, sorted by name so the order the container
    /// happens to store modules in cannot change the answer.
    pub modules: Vec<(String, String)>,
    /// Over every module name and digest in that order.
    pub project: String,
}

impl MacroDigest {
    fn module(&self, name: &str) -> Option<&str> {
        self.modules
            .iter()
            .find(|(module, _)| module == name)
            .map(|(_, digest)| digest.as_str())
    }
}

/// Hashes a project's macro source.
///
/// The module NAME is hashed alongside its text, so moving a body from one
/// module to another changes the answer: to a reader deciding whether to run
/// this file, that is a different file.
pub fn digest_project<'a>(modules: impl IntoIterator<Item = (&'a str, &'a str)>) -> MacroDigest {
    let mut digested: Vec<(String, String)> = modules
        .into_iter()
        .map(|(name, source)| {
            let mut hasher = Sha256::new();
            hasher.update(name.as_bytes());
            hasher.update([0u8]);
            hasher.update(source.as_bytes());
            (name.to_string(), hex(&hasher.finalize()))
        })
        .collect();
    digested.sort();

    let mut whole = Sha256::new();
    for (name, digest) in &digested {
        whole.update(name.as_bytes());
        whole.update([0u8]);
        whole.update(digest.as_bytes());
        whole.update([0u8]);
    }
    MacroDigest {
        modules: digested,
        project: hex(&whole.finalize()),
    }
}

/// Someone's seal on one exact project.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Attestation {
    /// The digest this seal was placed on.
    pub digest: MacroDigest,
    /// Who vouched for it, as they wish to be shown.
    pub signer: String,
    /// When, as written by whoever sealed it. Not checked here.
    pub sealed_at: String,
    /// What they are vouching for, in their words.
    pub note: String,
    /// Detached signature over [`Attestation::signed_bytes`]. `None` means
    /// nobody signed, and the seal cannot clear anything.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub signature: Option<Vec<u8>>,
    /// The signer's certificate, for whoever verifies.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub certificate: Option<Vec<u8>>,
}

impl Attestation {
    /// Exactly the bytes a signature covers. Everything a reader relies on is
    /// inside: the digest, who sealed it, when, and what they said.
    pub fn signed_bytes(&self) -> Vec<u8> {
        let mut out = String::new();
        let _ = writeln!(out, "oxihanko-macro-attestation-v1");
        let _ = writeln!(out, "project\t{}", self.digest.project);
        for (name, digest) in &self.digest.modules {
            let _ = writeln!(out, "module\t{name}\t{digest}");
        }
        let _ = writeln!(out, "signer\t{}", self.signer);
        let _ = writeln!(out, "sealed_at\t{}", self.sealed_at);
        let _ = writeln!(out, "note\t{}", self.note);
        out.into_bytes()
    }
}

/// Checks a detached signature. Supplied by the caller, because who counts as
/// an authority is a decision about an organisation and not about VBA.
pub trait SignatureVerifier {
    /// `Ok(name)` names the verified signer. `Err` explains what failed, in
    /// words a reader can act on.
    fn verify(
        &self,
        signed_bytes: &[u8],
        signature: &[u8],
        certificate: Option<&[u8]>,
    ) -> Result<String, String>;
}

/// What a seal says about the file in front of you.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Clearance {
    /// Signed, verified, and over exactly this code. The only answer that
    /// earns silence.
    Cleared { signer: String, note: String },
    /// The digest matches and nobody signed it. Anyone who can edit the
    /// workbook can write this file, so it clears nothing.
    MatchesButUnsigned,
    /// The digest matches and a signature is there, but no verifier was
    /// supplied, so nobody has said whose it is. Not the same as unsigned,
    /// and telling a reader otherwise would be untrue.
    MatchesButNobodyChecked,
    /// There is a seal, but the code has moved since. Carries what moved, so a
    /// reviewer reads the difference rather than the project.
    Stale(Changed),
    /// A signature was present and did not verify.
    SignatureRejected { why: String },
    /// No seal at all.
    NotSealed,
}

impl Clearance {
    /// Whether the safety report may stay quiet.
    pub fn is_cleared(&self) -> bool {
        matches!(self, Clearance::Cleared { .. })
    }
}

/// Which modules moved since the seal.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Changed {
    pub edited: Vec<String>,
    pub added: Vec<String>,
    pub removed: Vec<String>,
}

impl Changed {
    pub fn is_empty(&self) -> bool {
        self.edited.is_empty() && self.added.is_empty() && self.removed.is_empty()
    }
}

/// Compares what is in front of you with what was sealed.
///
/// `verifier` is what separates a seal from a note in a folder. Passing `None`
/// can never return [`Clearance::Cleared`].
pub fn clearance(
    now: &MacroDigest,
    attestation: Option<&Attestation>,
    verifier: Option<&dyn SignatureVerifier>,
) -> Clearance {
    let Some(attestation) = attestation else {
        return Clearance::NotSealed;
    };
    if attestation.digest.project != now.project {
        return Clearance::Stale(changed_since(&attestation.digest, now));
    }
    let Some(signature) = &attestation.signature else {
        return Clearance::MatchesButUnsigned;
    };
    let Some(verifier) = verifier else {
        return Clearance::MatchesButNobodyChecked;
    };
    match verifier.verify(
        &attestation.signed_bytes(),
        signature,
        attestation.certificate.as_deref(),
    ) {
        Ok(signer) => Clearance::Cleared {
            signer,
            note: attestation.note.clone(),
        },
        Err(why) => Clearance::SignatureRejected { why },
    }
}

/// Which modules differ between a sealed digest and the one in front of you.
pub fn changed_since(sealed: &MacroDigest, now: &MacroDigest) -> Changed {
    let mut changed = Changed::default();
    for (name, digest) in &now.modules {
        match sealed.module(name) {
            Some(sealed_digest) if sealed_digest == digest => {}
            Some(_) => changed.edited.push(name.clone()),
            None => changed.added.push(name.clone()),
        }
    }
    for (name, _) in &sealed.modules {
        if now.module(name).is_none() {
            changed.removed.push(name.clone());
        }
    }
    changed
}

fn hex(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(out, "{byte:02x}");
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    struct AlwaysTrusts;
    impl SignatureVerifier for AlwaysTrusts {
        fn verify(&self, _: &[u8], signature: &[u8], _: Option<&[u8]>) -> Result<String, String> {
            if signature == b"good" {
                Ok("Someone In Accounts".to_string())
            } else {
                Err("the signature does not match these bytes".to_string())
            }
        }
    }

    fn sealed(modules: &[(&str, &str)], signature: Option<&[u8]>) -> Attestation {
        Attestation {
            digest: digest_project(modules.iter().copied()),
            signer: "Someone In Accounts".to_string(),
            sealed_at: "2026-09-02".to_string(),
            note: "reviewed the fetch and the sheet writes".to_string(),
            signature: signature.map(|s| s.to_vec()),
            certificate: None,
        }
    }

    #[test]
    fn the_same_source_digests_the_same_whatever_order_it_arrives_in() {
        let one = digest_project([("A", "Sub Go()\nEnd Sub"), ("B", "Sub Stop2()\nEnd Sub")]);
        let other = digest_project([("B", "Sub Stop2()\nEnd Sub"), ("A", "Sub Go()\nEnd Sub")]);
        assert_eq!(one, other);
    }

    #[test]
    fn moving_a_body_to_another_module_is_a_different_project() {
        let one = digest_project([("A", "Sub Go()\nEnd Sub"), ("B", "")]);
        let other = digest_project([("A", ""), ("B", "Sub Go()\nEnd Sub")]);
        assert_ne!(one.project, other.project);
    }

    /// The whole point: a signed seal over exactly this code earns silence.
    #[test]
    fn a_verified_seal_clears_the_file() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = sealed(&modules, Some(b"good"));
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&AlwaysTrusts),
        );
        assert!(verdict.is_cleared());
        match verdict {
            Clearance::Cleared { signer, .. } => assert_eq!(signer, "Someone In Accounts"),
            other => panic!("{other:?}"),
        }
    }

    /// Signed but unchecked is not the same as unsigned, and a report that
    /// said "nothing signed it" about a signed seal would be lying.
    #[test]
    fn signed_but_unchecked_is_its_own_answer() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = sealed(&modules, Some(b"good"));
        assert_eq!(
            clearance(&digest_project(modules.iter().copied()), Some(&attestation), None),
            Clearance::MatchesButNobodyChecked
        );
    }

    /// A digest match with nobody's name behind it is not a clearance, because
    /// whoever can edit the workbook can write the file next to it.
    #[test]
    fn an_unsigned_seal_clears_nothing() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = sealed(&modules, None);
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&AlwaysTrusts),
        );
        assert_eq!(verdict, Clearance::MatchesButUnsigned);
        assert!(!verdict.is_cleared());
    }

    /// Even a real signature cannot clear a file when nobody is there to check
    /// it. Passing no verifier must not be a way to reach silence.
    #[test]
    fn without_a_verifier_nothing_is_cleared() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = sealed(&modules, Some(b"good"));
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            None,
        );
        assert!(!verdict.is_cleared());
    }

    /// The edit a seal most needs to catch: one character inside a literal.
    #[test]
    fn changing_one_literal_breaks_the_seal() {
        let before = [("A", "Sub Go()\n  Fetch \"https://ours.example\"\nEnd Sub")];
        let after = [("A", "Sub Go()\n  Fetch \"https://theirs.example\"\nEnd Sub")];
        let attestation = sealed(&before, Some(b"good"));
        let verdict = clearance(
            &digest_project(after.iter().copied()),
            Some(&attestation),
            Some(&AlwaysTrusts),
        );
        match verdict {
            Clearance::Stale(changed) => assert_eq!(changed.edited, ["A"]),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn a_broken_seal_says_which_modules_moved() {
        let before = [("Kept", "Sub A()\nEnd Sub"), ("Gone", "Sub B()\nEnd Sub")];
        let after = [("Kept", "Sub A()\nEnd Sub"), ("New", "Sub C()\nEnd Sub")];
        let attestation = sealed(&before, Some(b"good"));
        let verdict = clearance(
            &digest_project(after.iter().copied()),
            Some(&attestation),
            Some(&AlwaysTrusts),
        );
        match verdict {
            Clearance::Stale(changed) => {
                assert_eq!(changed.added, ["New"]);
                assert_eq!(changed.removed, ["Gone"]);
                assert!(changed.edited.is_empty());
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn a_signature_that_does_not_check_out_is_not_silence() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = sealed(&modules, Some(b"forged"));
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&AlwaysTrusts),
        );
        assert!(matches!(verdict, Clearance::SignatureRejected { .. }));
    }

    #[test]
    fn no_seal_at_all_is_its_own_answer() {
        let now = digest_project([("A", "Sub Go()\nEnd Sub")]);
        assert_eq!(clearance(&now, None, None), Clearance::NotSealed);
    }

    /// What a signature covers has to include who sealed it and what they
    /// said, or those could be rewritten under a valid signature.
    #[test]
    fn the_signed_bytes_cover_the_signer_and_the_note() {
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let mut one = sealed(&modules, Some(b"good"));
        let first = one.signed_bytes();
        one.signer = "Someone Else".to_string();
        assert_ne!(first, one.signed_bytes());
        one.signer = "Someone In Accounts".to_string();
        one.note = "did not actually read it".to_string();
        assert_ne!(first, one.signed_bytes());
    }
}
