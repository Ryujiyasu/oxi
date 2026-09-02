// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Checking the signature on a seal against a list of people trusted to place
//! one.
//!
//! Behind the `ring` feature, because who may clear a macro is a question
//! about an organisation, and because a browser build has no business carrying
//! a signature stack it will never use.
//!
//! # What this is, and what it is not
//!
//! This is the simplest arrangement that actually holds: a file listing the
//! public keys of the people allowed to seal macros, and Ed25519 signatures
//! over [`Attestation::signed_bytes`]. No certificate chain, no revocation, no
//! expiry — a key is trusted because it is in the list, and untrusted the
//! moment it is taken out.
//!
//! That is a real answer for one organisation with one list, and it is
//! deliberately not a PKI. An organisation that already has one implements
//! [`SignatureVerifier`] against it instead; nothing else here changes, which
//! is the whole reason verification is a trait.
//!
//! No cryptography is implemented here. The signature check is `ring`'s.

use std::collections::BTreeMap;

use ring::signature::{UnparsedPublicKey, ED25519};

use crate::attest::SignatureVerifier;

/// The people allowed to seal a macro, by name.
///
/// A name is what a reader sees in the report, so it should be the name they
/// would recognise on a colleague — not a key id.
#[derive(Debug, Clone, Default)]
pub struct TrustedSigners {
    keys: BTreeMap<String, Vec<u8>>,
}

impl TrustedSigners {
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds one trusted key. An Ed25519 public key is 32 bytes.
    pub fn trust(&mut self, name: impl Into<String>, public_key: Vec<u8>) -> Result<(), String> {
        if public_key.len() != 32 {
            return Err(format!(
                "an Ed25519 public key is 32 bytes; this one is {}",
                public_key.len()
            ));
        }
        self.keys.insert(name.into(), public_key);
        Ok(())
    }

    pub fn is_empty(&self) -> bool {
        self.keys.is_empty()
    }

    pub fn names(&self) -> impl Iterator<Item = &str> {
        self.keys.keys().map(String::as_str)
    }

    /// Reads a list written one signer per line as `name<TAB>hex-public-key`.
    /// Blank lines and lines starting with `#` are skipped, so the file can
    /// say who added whom and when.
    pub fn from_list(text: &str) -> Result<Self, String> {
        let mut signers = Self::new();
        for (index, line) in text.lines().enumerate() {
            let line = line.trim();
            if line.is_empty() || line.starts_with('#') {
                continue;
            }
            let (name, key) = line
                .split_once('\t')
                .or_else(|| line.rsplit_once(' '))
                .ok_or_else(|| {
                    format!(
                        "line {}: expected a name and a hex public key, separated by a tab",
                        index + 1
                    )
                })?;
            let key = decode_hex(key.trim())
                .map_err(|error| format!("line {}: {error}", index + 1))?;
            signers.trust(name.trim(), key)?;
        }
        Ok(signers)
    }
}

impl SignatureVerifier for TrustedSigners {
    /// Tries every trusted key and reports the name behind the one that
    /// verified.
    ///
    /// The seal does not say which key signed it, and it is not asked: a name
    /// written inside an unverified file is worth nothing, so the name that
    /// reaches the reader is the one attached to the key that actually checked
    /// out. `certificate` is ignored — this arrangement has no chain to walk.
    fn verify(
        &self,
        signed_bytes: &[u8],
        signature: &[u8],
        _certificate: Option<&[u8]>,
    ) -> Result<String, String> {
        if self.keys.is_empty() {
            return Err("nobody is trusted to seal macros yet, so no seal can clear".to_string());
        }
        for (name, key) in &self.keys {
            let public_key = UnparsedPublicKey::new(&ED25519, key);
            if public_key.verify(signed_bytes, signature).is_ok() {
                return Ok(name.clone());
            }
        }
        Err(format!(
            "the signature matches none of the {} trusted signer(s)",
            self.keys.len()
        ))
    }
}

fn decode_hex(text: &str) -> Result<Vec<u8>, String> {
    if !text.len().is_multiple_of(2) {
        return Err("a hex key has an even number of characters".to_string());
    }
    (0..text.len())
        .step_by(2)
        .map(|index| {
            u8::from_str_radix(&text[index..index + 2], 16)
                .map_err(|_| format!("{:?} is not hexadecimal", &text[index..index + 2]))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::attest::{clearance, digest_project, Attestation, Clearance};
    use ring::rand::SystemRandom;
    use ring::signature::{Ed25519KeyPair, KeyPair};

    fn a_key() -> Ed25519KeyPair {
        let random = SystemRandom::new();
        let document = Ed25519KeyPair::generate_pkcs8(&random).expect("a key");
        Ed25519KeyPair::from_pkcs8(document.as_ref()).expect("a key pair")
    }

    fn seal(modules: &[(&str, &str)], key: &Ed25519KeyPair, signer: &str) -> Attestation {
        let mut attestation = Attestation {
            digest: digest_project(modules.iter().copied()),
            signer: signer.to_string(),
            sealed_at: "2026-09-02".to_string(),
            note: "read it".to_string(),
            signature: None,
            certificate: None,
        };
        let signature = key.sign(&attestation.signed_bytes());
        attestation.signature = Some(signature.as_ref().to_vec());
        attestation
    }

    #[test]
    fn a_trusted_signer_clears_the_file() {
        let key = a_key();
        let mut trusted = TrustedSigners::new();
        trusted
            .trust("Someone In Accounts", key.public_key().as_ref().to_vec())
            .unwrap();
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = seal(&modules, &key, "whatever the file claims");
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&trusted),
        );
        // The name comes from the key that verified, not from the file.
        match verdict {
            Clearance::Cleared { signer, .. } => assert_eq!(signer, "Someone In Accounts"),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn a_signer_nobody_trusts_clears_nothing() {
        let theirs = a_key();
        let mut trusted = TrustedSigners::new();
        trusted
            .trust("Someone In Accounts", a_key().public_key().as_ref().to_vec())
            .unwrap();
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = seal(&modules, &theirs, "Someone In Accounts");
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&trusted),
        );
        assert!(matches!(verdict, Clearance::SignatureRejected { .. }));
    }

    /// Editing the note under a good signature must not survive, because the
    /// note is what a reader is being asked to rely on.
    #[test]
    fn rewriting_the_note_breaks_the_signature() {
        let key = a_key();
        let mut trusted = TrustedSigners::new();
        trusted.trust("Reader", key.public_key().as_ref().to_vec()).unwrap();
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let mut attestation = seal(&modules, &key, "Reader");
        attestation.note = "cleared everything, honest".to_string();
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&trusted),
        );
        assert!(matches!(verdict, Clearance::SignatureRejected { .. }));
    }

    #[test]
    fn an_empty_trust_list_clears_nothing() {
        let key = a_key();
        let modules = [("A", "Sub Go()\nEnd Sub")];
        let attestation = seal(&modules, &key, "Nobody");
        let verdict = clearance(
            &digest_project(modules.iter().copied()),
            Some(&attestation),
            Some(&TrustedSigners::new()),
        );
        assert!(matches!(verdict, Clearance::SignatureRejected { .. }));
    }

    #[test]
    fn a_signer_list_is_read_from_text() {
        let key = a_key();
        let hex: String = key
            .public_key()
            .as_ref()
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect();
        let text = format!("# added 2026-09-02\n\nSomeone In Accounts\t{hex}\n");
        let trusted = TrustedSigners::from_list(&text).expect("a list");
        assert_eq!(trusted.names().collect::<Vec<_>>(), ["Someone In Accounts"]);
    }

    #[test]
    fn a_key_of_the_wrong_length_is_refused_when_it_is_added() {
        let mut trusted = TrustedSigners::new();
        assert!(trusted.trust("Short", vec![0u8; 16]).is_err());
        assert!(trusted.is_empty());
    }
}
