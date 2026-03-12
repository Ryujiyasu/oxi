/// A raw PDF object as parsed from the file.
#[derive(Debug, Clone, PartialEq)]
pub enum PdfObject {
    Null,
    Boolean(bool),
    Integer(i64),
    Real(f64),
    String(Vec<u8>),
    HexString(Vec<u8>),
    Name(String),
    Array(Vec<PdfObject>),
    Dictionary(Vec<(String, PdfObject)>),
    Stream {
        dict: Vec<(String, PdfObject)>,
        data: Vec<u8>,
    },
    Reference(ObjRef),
}

/// An indirect object reference (object number + generation).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ObjRef {
    pub num: u32,
    pub gen: u16,
}

impl PdfObject {
    /// Try to get this object as an integer.
    pub fn as_i64(&self) -> Option<i64> {
        match self {
            PdfObject::Integer(n) => Some(*n),
            _ => None,
        }
    }

    /// Try to get this object as a float (accepts Integer too).
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            PdfObject::Real(n) => Some(*n),
            PdfObject::Integer(n) => Some(*n as f64),
            _ => None,
        }
    }

    /// Try to get this object as a name string.
    pub fn as_name(&self) -> Option<&str> {
        match self {
            PdfObject::Name(s) => Some(s.as_str()),
            _ => None,
        }
    }

    /// Try to get this object as a string (literal or hex).
    pub fn as_bytes(&self) -> Option<&[u8]> {
        match self {
            PdfObject::String(b) | PdfObject::HexString(b) => Some(b),
            _ => None,
        }
    }

    /// Try to get this object as a dictionary.
    pub fn as_dict(&self) -> Option<&[(String, PdfObject)]> {
        match self {
            PdfObject::Dictionary(entries) => Some(entries),
            PdfObject::Stream { dict, .. } => Some(dict),
            _ => None,
        }
    }

    /// Try to get this object as an array.
    pub fn as_array(&self) -> Option<&[PdfObject]> {
        match self {
            PdfObject::Array(items) => Some(items),
            _ => None,
        }
    }

    /// Try to get this object as a reference.
    pub fn as_ref(&self) -> Option<ObjRef> {
        match self {
            PdfObject::Reference(r) => Some(*r),
            _ => None,
        }
    }

    /// Look up a key in a dictionary object.
    pub fn dict_get(&self, key: &str) -> Option<&PdfObject> {
        self.as_dict()
            .and_then(|entries| entries.iter().find(|(k, _)| k == key).map(|(_, v)| v))
    }
}
