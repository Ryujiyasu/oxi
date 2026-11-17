// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The host supplies local calendar time; the interpreter owns no timezone.

use std::{fmt, rc::Rc};

/// A shared host clock returning a finite, valid local OLE Automation date.
/// Read it when VBA asks for time, including during a running procedure.
/// Clones share the provider across projects and their virtual files.
#[derive(Clone)]
pub struct Clock(Rc<dyn Fn() -> f64>);

impl Clock {
    pub fn new(read: impl Fn() -> f64 + 'static) -> Self {
        Self(Rc::new(read))
    }

    pub fn now(&self) -> f64 {
        (self.0)()
    }
}

impl fmt::Debug for Clock {
    fn fmt(&self, out: &mut fmt::Formatter<'_>) -> fmt::Result {
        out.write_str("Clock { .. }")
    }
}
