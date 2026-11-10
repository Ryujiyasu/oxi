// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Lossless VBA string values. UTF-16 code units, including lone surrogates,
//! are data; a Rust Unicode scalar string is used only when it is lossless.
use crate::runtime::Value;
use std::borrow::Cow;

#[derive(Debug)]
pub enum StringTextError {
    InvalidUtf16(std::string::FromUtf16Error),
    OddByteLength,
}
impl std::fmt::Display for StringTextError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidUtf16(error) => error.fmt(f),
            Self::OddByteLength => f.write_str("BSTR has an incomplete UTF-16 unit"),
        }
    }
}
impl std::error::Error for StringTextError {}

const ANSI_DATA: &[u8] = include_bytes!("data/vba_ansi_codepages.bin");

/// Portable ANSI conversion using a Windows code page or VBA conversion locale.
/// The built-in profile uses Western defaults; hosts can supply their own
/// system and user pages without depending on a platform API at execution time.
#[derive(Clone, Copy, Debug)]
pub struct AnsiCodec {
    descriptor: usize,
}

fn ansi_u16(offset: usize) -> u16 {
    u16::from_le_bytes([ANSI_DATA[offset], ANSI_DATA[offset + 1]])
}

fn ansi_u32(offset: usize) -> u32 {
    u32::from_le_bytes(ANSI_DATA[offset..offset + 4].try_into().expect("ANSI table word"))
}

impl AnsiCodec {
    pub fn from_code_page(page: u16) -> Option<Self> {
        let count = ansi_u32(8) as usize;
        let mut low = 0;
        let mut high = count;
        while low < high {
            let middle = (low + high) / 2;
            if ansi_u32(16 + middle * 28) < u32::from(page) { low = middle + 1; }
            else { high = middle; }
        }
        (low < count && ansi_u32(16 + low * 28) == u32::from(page))
            .then_some(Self { descriptor: 16 + low * 28 })
    }

    pub fn for_locale(locale: Option<i64>, system_page: u16, user_page: u16) -> Option<Self> {
        let page = match locale {
            None | Some(2048) => u32::from(system_page),
            Some(0 | 1 | 1024) => u32::from(user_page),
            Some(locale) => {
                let locale = u32::try_from(locale).ok()?;
                let start = 16 + ansi_u32(8) as usize * 28;
                let count = ansi_u32(12) as usize;
                let mut low = 0;
                let mut high = count;
                while low < high {
                    let middle = (low + high) / 2;
                    if ansi_u32(start + middle * 8) < locale { low = middle + 1; }
                    else { high = middle; }
                }
                if low == count || ansi_u32(start + low * 8) != locale { return None; }
                match ansi_u32(start + low * 8 + 4) {
                    0 => u32::from(system_page),
                    page => page,
                }
            }
        };
        Self::from_code_page(u16::try_from(page).ok()?)
    }

    pub fn encode(self, units: &[u16]) -> Vec<u8> {
        let table = ansi_u32(self.descriptor + 4) as usize;
        let count = ansi_u32(self.descriptor + 8) as usize;
        let mut output = Vec::with_capacity(units.len());
        // The fallback is per UTF-16 unit, including each half of a surrogate
        // pair. It is not a conversion through Unicode scalar characters.
        for &unit in units {
            if unit < 128 { output.push(unit as u8); continue; }
            let mut low = 0;
            let mut high = count;
            while low < high {
                let middle = (low + high) / 2;
                if ansi_u16(table + middle * 4) < unit { low = middle + 1; }
                else { high = middle; }
            }
            let code = if low < count && ansi_u16(table + low * 4) == unit {
                ansi_u16(table + low * 4 + 2)
            } else { u16::from(b'?') };
            if code > 255 { output.push((code >> 8) as u8); }
            output.push(code as u8);
        }
        output
    }

    pub fn decode(self, bytes: &[u8]) -> Vec<u16> {
        let singles = ansi_u32(self.descriptor + 12) as usize;
        let leads = ansi_u32(self.descriptor + 16) as usize;
        let pairs = ansi_u32(self.descriptor + 20) as usize;
        let mut output = Vec::with_capacity(bytes.len());
        let mut index = 0;
        while index < bytes.len() {
            let lead = ansi_u16(leads + bytes[index] as usize * 2);
            let offset = if lead != u16::MAX && index + 1 < bytes.len() {
                let offset = pairs + (lead as usize * 256 + bytes[index + 1] as usize) * 5;
                index += 2;
                offset
            } else {
                let offset = singles + bytes[index] as usize * 5;
                index += 1;
                offset
            };
            output.push(ansi_u16(offset + 1));
            // A lead followed by NUL can produce two units; retain both.
            if ANSI_DATA[offset] == 2 { output.push(ansi_u16(offset + 3)); }
        }
        output
    }
}

impl Value {
    pub fn from_bstr_bytes(bytes: Vec<u8>) -> Self {
        if bytes.len() % 2 == 1 { return Self::ByteString(bytes); }
        Self::from_utf16(bytes.chunks_exact(2).map(|b| u16::from_le_bytes([b[0], b[1]])).collect())
    }
    pub fn string_bytes(&self) -> Option<Cow<'_, [u8]>> {
        match self {
            Self::ByteString(bytes) => Some(Cow::Borrowed(bytes)),
            _ => self.string_units().map(|u| Cow::Owned(u.iter().flat_map(|u| u.to_le_bytes()).collect())),
        }
    }
    /// A scalar host read includes the BSTR's trailing zero byte when its
    /// byte count is odd. Character-count operations use string_units instead.
    pub fn padded_string_units(&self) -> Option<Cow<'_, [u16]>> {
        match self {
            Self::ByteString(bytes) => Some(Cow::Owned(bytes.chunks(2)
                .map(|b|u16::from_le_bytes([b[0],b.get(1).copied().unwrap_or(0)])).collect())),
            _ => self.string_units(),
        }
    }
    pub fn binary_string_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        let a = self.string_bytes()?;
        let b = other.string_bytes()?;
        let common = a.len().min(b.len());
        for i in (0..common - common % 2).step_by(2) {
            let order = u16::from_le_bytes([a[i], a[i+1]]).cmp(&u16::from_le_bytes([b[i], b[i+1]]));
            if !order.is_eq() { return Some(order); }
        }
        if common % 2 == 1 {
            let order = a[common-1].cmp(&b[common-1]);
            if !order.is_eq() { return Some(order); }
        }
        Some(a.len().cmp(&b.len()))
    }
    pub fn from_utf16(units: Vec<u16>) -> Self {
        match String::from_utf16(&units) {
            Ok(text) => Self::String(text),
            Err(_) => Self::Utf16(units),
        }
    }
    pub fn is_string(&self) -> bool {
        matches!(self, Self::String(_) | Self::Utf16(_) | Self::ByteString(_) | Self::NullString)
    }
    pub fn string_units(&self) -> Option<Cow<'_, [u16]>> {
        match self {
            Self::String(s) => Some(Cow::Owned(s.encode_utf16().collect())),
            Self::Utf16(s) => Some(Cow::Borrowed(s)),
            Self::ByteString(s) => Some(Cow::Owned(s.chunks_exact(2).map(|b|u16::from_le_bytes([b[0],b[1]])).collect())),
            Self::NullString => Some(Cow::Borrowed(&[])),
            _ => None,
        }
    }
    /// Conversion for interfaces requiring Unicode scalar text. Invalid units
    /// are reported, never replaced. VBA string operations use string_units.
    pub fn unicode_text(&self) -> Option<Result<Cow<'_, str>, StringTextError>> {
        match self {
            Self::String(s) => Some(Ok(Cow::Borrowed(s))),
            Self::NullString => Some(Ok(Cow::Borrowed(""))),
            Self::Utf16(s) => Some(String::from_utf16(s).map(Cow::Owned).map_err(StringTextError::InvalidUtf16)),
            Self::ByteString(_) => Some(Err(StringTextError::OddByteLength)),
            _ => None,
        }
    }
}
impl PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        if let (Self::String(a), Self::String(b)) = (self, other) {
            return a == b;
        }
        if let (Some(a), Some(b)) = (self.string_bytes(), other.string_bytes()) {
            return a == b;
        }
        use Value::*;
        match (self, other) {
            (Empty, Empty) | (Missing, Missing) | (Nothing, Nothing) | (Null, Null) => true,
            (Boolean(a), Boolean(b)) => a == b,
            (Int16(a), Int16(b)) => a == b,
            (Integer(a), Integer(b)) => a == b,
            (LongLong(a), LongLong(b)) => a == b,
            (Byte(a), Byte(b)) => a == b,
            (Single(a), Single(b)) => a == b,
            (Double(a), Double(b)) => a == b,
            (Currency(a), Currency(b)) => a == b,
            (Decimal(a), Decimal(b)) => a == b,
            (Date(a), Date(b)) => a == b,
            (Error(a), Error(b)) => a == b,
            (Array(a), Array(b)) => a == b,
            (Record(a), Record(b)) => a == b,
            (Object(a), Object(b)) => a == b,
            _ => false,
        }
    }
}

#[cfg(test)]
mod odd_byte_native_tests {
    #[test]
    fn ansi_locales_match_excel934_fingerprints() {
        // Exact input LCIDs and output fingerprint from the original Excel64
        // program, including locale1's user default and locale2's error5.
        let locales: &[i64] = &[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 124, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 140, 145, 146, 1024, 1025, 1026, 1027, 1028, 1029, 1030, 1031, 1032, 1033, 1034, 1035, 1036, 1037, 1038, 1039, 1040, 1041, 1042, 1043, 1044, 1045, 1046, 1047, 1048, 1049, 1050, 1051, 1052, 1053, 1054, 1055, 1056, 1057, 1058, 1059, 1060, 1061, 1062, 1063, 1064, 1065, 1066, 1067, 1068, 1069, 1070, 1071, 1072, 1073, 1074, 1075, 1076, 1077, 1078, 1079, 1080, 1081, 1082, 1083, 1085, 1086, 1087, 1088, 1089, 1090, 1091, 1092, 1093, 1094, 1095, 1096, 1097, 1098, 1099, 1100, 1101, 1102, 1103, 1104, 1105, 1106, 1107, 1108, 1109, 1110, 1111, 1112, 1113, 1114, 1115, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127, 1128, 1129, 1130, 1131, 1132, 1133, 1134, 1135, 1136, 1137, 1138, 1139, 1140, 1141, 1142, 1143, 1144, 1145, 1146, 1148, 1150, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1164, 1169, 1170, 1281, 1534, 2048, 2049, 2051, 2052, 2055, 2057, 2058, 2060, 2064, 2067, 2068, 2070, 2072, 2073, 2074, 2077, 2080, 2092, 2094, 2098, 2107, 2108, 2110, 2115, 2117, 2118, 2121, 2128, 2137, 2141, 2143, 2144, 2145, 2151, 2155, 2163, 2305, 2559, 3072, 3073, 3076, 3079, 3081, 3082, 3084, 3098, 3131, 3152, 3153, 3179, 4096, 4097, 4100, 4103, 4105, 4106, 4108, 4122, 4155, 4191, 5121, 5124, 5127, 5129, 5130, 5132, 5146, 5179, 6145, 6153, 6154, 6156, 6170, 6203, 7169, 7177, 7178, 7180, 7194, 7227, 8192, 8193, 8201, 8202, 8204, 8218, 8251, 9216, 9217, 9225, 9226, 9228, 9242, 9275, 10240, 10241, 10249, 10250, 10252, 10266, 11264, 11265, 11273, 11274, 11276, 11290, 12288, 12289, 12297, 12298, 12300, 12314, 13312, 13313, 13321, 13322, 13324, 14336, 14337, 14345, 14346, 14348, 15360, 15361, 15369, 15370, 15372, 16384, 16385, 16393, 16394, 17408, 17417, 17418, 18432, 18441, 18442, 19456, 19465, 19466, 20490, 21514, 22538, 23562, 25626, 26650, 27674, 28698, 28731, 29740, 29755, 30724, 30740, 30746, 30764, 30779, 30787, 30800, 30813, 30815, 31748, 31764, 31770, 31784, 31790, 31803, 31811, 31814, 31824, 31833, 31836, 31837, 31839, 31847, 31848, 31878, 31890, 65663, 66567, 66574, 66615, 133124, 135172, 136196, 197636, 263172, 263185, 265220, 267268, 329732, 331780];
        let signature: &[u16] = &[0x20ac, 0x410, 0x391, 0x5d0, 0x627, 0xe01, 0x11e, 0x104, 0x100, 0x1eae, 0x3042, 0xac00, 0x6c49, 0x6f22];
        let source: Vec<u8> = (128..=255).collect();
        let snapshot = |raw: &[u8]| {
            let units = raw.chunks_exact(2).map(|b| format!("{:04X}",u16::from_le_bytes([b[0],b[1]]))).collect::<String>();
            let bytes = raw.iter().map(|b|format!("{b:02X}")).collect::<String>();
            format!("S|L{}B{}U{units}R{bytes}",raw.len()/2,raw.len())
        };
        let mut rows = Vec::new();
        for &locale in locales {
            let codec = super::AnsiCodec::for_locale(Some(locale),932,1252);
            let encoded = codec.map_or_else(||"ERR|5".to_string(),|codec|snapshot(&codec.encode(signature)));
            let decoded = codec.map_or_else(||"ERR|5".to_string(),|codec|snapshot(&codec.decode(&source).into_iter().flat_map(u16::to_le_bytes).collect::<Vec<_>>()));
            rows.push(format!("{locale}.E:{encoded}"));
            rows.push(format!("{locale}.D:{decoded}"));
        }
        let actual = rows.join(";");
        let fingerprint = actual.bytes().fold(0xcbf29ce484222325u64,|value,b|(value^u64::from(b)).wrapping_mul(0x100000001b3));
        assert_eq!(fingerprint,15835503726520584194u64);
    }

    use crate::{parse_module, Runtime, Value};

    #[test]
    fn odd_byte_string_operations_match_excel() {
        let source = r####"Attribute VB_Name = "OddStringOperations"
Option Explicit
Public Function Probe() As String
Dim data() As Byte
Dim value As String
Dim sample As Long
Dim operation As Long
For sample = 1 To 3
If sample = 1 Then
ReDim data(0 To 0)
data(0) = 65
Else
ReDim data(0 To 2)
If sample = 2 Then
data(0) = 65
data(1) = 0
data(2) = 66
Else
data(0) = 0
data(1) = 216
data(2) = 65
End If
End If
value = data
For operation = 1 To 18
Probe = Probe & CStr(sample) & "." & CStr(operation) & ":" & Action(value, operation) & ";"
Next operation
Next sample
End Function
Private Function Action(ByVal value As String, ByVal operation As Long) As String
On Error GoTo Failed
Dim result As String
Select Case operation
Case 1: result = value
Case 2: result = CStr(value)
Case 3: result = CVar(value)
Case 4: result = value & "X"
Case 5: result = "X" & value
Case 6: result = Left$(value, 1)
Case 7: result = Right$(value, 1)
Case 8: result = Mid$(value, 1, 1)
Case 9: result = UCase$(value)
Case 10: result = LCase$(value)
Case 11: result = Trim$(value)
Case 12: result = StrReverse(value)
Case 13: result = Replace(value, "A", "Z")
Case 14: result = CStr(value = "")
Case 15: result = CStr(value = "A")
Case 16: result = CStr(StrComp(value, "", vbBinaryCompare))
Case 17: result = CStr(InStrB(1, value, "B", vbBinaryCompare))
Case 18: result = CStr(AscW(value))
End Select
Action = "L" & CStr(Len(result)) & "B" & CStr(LenB(result)) & "U" & Units(result) & "R" & RawBytes(result)
Exit Function
Failed:
Action = "ERR" & CStr(Err.Number)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####;
        let expected = r####"1.1:L0B1UR41;1.2:L0B1UR41;1.3:L0B1UR41;1.4:L1B3U5841R415800;1.5:L1B3U0058R580041;1.6:L0B1UR41;1.7:L0B1UR41;1.8:L0B1UR41;1.9:L0B0UR;1.10:L0B0UR;1.11:L0B0UR;1.12:L0B0UR;1.13:L0B0UR;1.14:L5B10U00460061006C00730065R460061006C0073006500;1.15:L5B10U00460061006C00730065R460061006C0073006500;1.16:L1B2U0031R3100;1.17:L1B2U0030R3000;1.18:ERR5;2.1:L1B3U0041R410042;2.2:L1B3U0041R410042;2.3:L1B3U0041R410042;2.4:L2B5U00415842R4100425800;2.5:L2B5U00580041R5800410042;2.6:L1B2U0041R4100;2.7:L1B2U4200R0042;2.8:L1B2U0041R4100;2.9:L1B2U0041R4100;2.10:L1B2U0061R6100;2.11:L1B2U0041R4100;2.12:L1B2U0041R4100;2.13:L1B2U005AR5A00;2.14:L5B10U00460061006C00730065R460061006C0073006500;2.15:L5B10U00460061006C00730065R460061006C0073006500;2.16:L1B2U0031R3100;2.17:L1B2U0030R3000;2.18:L2B4U00360035R36003500;3.1:L1B3UD800R00D841;3.2:L1B3UD800R00D841;3.3:L1B3UD800R00D841;3.4:L2B5UD8005841R00D8415800;3.5:L2B5U0058D800R580000D841;3.6:L1B2UD800R00D8;3.7:L1B2U41D8RD841;3.8:L1B2UD800R00D8;3.9:L1B2UD800R00D8;3.10:L1B2UD800R00D8;3.11:L1B2UD800R00D8;3.12:L1B2UD800R00D8;3.13:L1B2UD800R00D8;3.14:L5B10U00460061006C00730065R460061006C0073006500;3.15:L5B10U00460061006C00730065R460061006C0073006500;3.16:L1B2U0031R3100;3.17:L1B2U0030R3000;3.18:L6B12U002D00310030003200340030R2D0031003000320034003000;"####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe", vec![]);
        assert_eq!(actual, Ok(Value::String(expected.to_owned())));
    }
    #[test]
    fn odd_byte_binary_comparisons_match_excel() {
        let source = r####"Attribute VB_Name = "OddStringComparison"
Option Explicit
Public Function Probe() As String
Dim data() As Byte
Dim values(0 To 9) As String
Dim i As Long
Dim j As Long
For i = 0 To 2
If i = 0 Then
ReDim data(0 To 0)
data(0) = 65
Else
ReDim data(0 To 2)
If i = 1 Then
data(0) = 65
data(1) = 0
data(2) = 66
Else
data(0) = 0
data(1) = 216
data(2) = 65
End If
End If
values(i) = data
Next i
values(3) = ""
values(4) = "A"
values(5) = "B"
values(6) = ChrW(&H100)
values(7) = ChrW(&HD800)
values(8) = "A" & ChrW(0)
values(9) = ChrW(0)
For i = 0 To 2
For j = 0 To 9
Probe = Probe & CStr(i) & "." & CStr(j) & ":" & CStr(StrComp(values(i), values(j), vbBinaryCompare)) & "," & CStr(values(i) = values(j)) & "," & CStr(values(i) < values(j)) & "," & CStr(values(i) > values(j)) & ";"
Next j
Next i
End Function
"####;
        let expected = r####"0.0:0,True,False,False;0.1:-1,False,True,False;0.2:1,False,False,True;0.3:1,False,False,True;0.4:-1,False,True,False;0.5:-1,False,True,False;0.6:1,False,False,True;0.7:1,False,False,True;0.8:-1,False,True,False;0.9:1,False,False,True;1.0:1,False,False,True;1.1:0,True,False,False;1.2:-1,False,True,False;1.3:1,False,False,True;1.4:1,False,False,True;1.5:-1,False,True,False;1.6:-1,False,True,False;1.7:-1,False,True,False;1.8:1,False,False,True;1.9:1,False,False,True;2.0:-1,False,True,False;2.1:1,False,False,True;2.2:0,True,False,False;2.3:1,False,False,True;2.4:1,False,False,True;2.5:1,False,False,True;2.6:1,False,False,True;2.7:1,False,False,True;2.8:1,False,False,True;2.9:1,False,False,True;"####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe", vec![]);
        assert_eq!(actual, Ok(Value::String(expected.to_owned())));
    }

    #[test]
    fn odd_byte_numeric_coercions_match_excel() {
        let source = r####"Attribute VB_Name = "OddStringCoercion"
Option Explicit
Public Function Probe() As String
Dim data() As Byte
Dim value As String
Dim sample As Long
Dim operation As Long
For sample = 0 To 3
If sample = 0 Then
ReDim data(0 To 0)
data(0) = 49
ElseIf sample = 3 Then
ReDim data(0 To 3)
data(0) = 49
data(1) = 0
data(2) = 50
data(3) = 0
Else
ReDim data(0 To 2)
data(0) = 49
data(1) = 0
If sample = 1 Then
data(2) = 50
Else
data(2) = 0
End If
End If
value = data
For operation = 1 To 8
Probe = Probe & CStr(sample) & "." & CStr(operation) & ":" & Action(value, operation) & ";"
Next operation
Next sample
End Function
Private Function Action(ByVal value As String, ByVal operation As Long) As String
On Error GoTo Failed
Select Case operation
Case 1: Action = CStr(CDbl(value))
Case 2: Action = CStr(CLng(value))
Case 3: Action = CStr(CInt(value))
Case 4: Action = CStr(CBool(value))
Case 5: Action = CStr(CCur(value))
Case 6: Action = CStr(CDbl(CDate(value)))
Case 7: Action = CStr(Val(value))
Case 8: Action = CStr(IsNumeric(value))
End Select
Exit Function
Failed:
Action = "ERR" & CStr(Err.Number)
End Function
"####;
        let expected = r####"0.1:1;0.2:1;0.3:1;0.4:True;0.5:1;0.6:1;0.7:1;0.8:True;1.1:12;1.2:12;1.3:12;1.4:True;1.5:12;1.6:12;1.7:12;1.8:True;2.1:1;2.2:1;2.3:1;2.4:True;2.5:1;2.6:1;2.7:1;2.8:True;3.1:12;3.2:12;3.3:12;3.4:True;3.5:12;3.6:12;3.7:12;3.8:True;"####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe", vec![]);
        assert_eq!(actual, Ok(Value::String(expected.to_owned())));
    }

    #[test]
    fn byte_array_string_boundaries_match_excel() {
        let cases = [
            (r####"fixed_1"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 0) As Byte
Dim target As String
data(0) = 65
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:0|LenB:1|Units:|Bytes:41"####),
            (r####"fixed_2"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 1) As Byte
Dim target As String
data(0) = 65
data(1) = 0
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:1|LenB:2|Units:0041|Bytes:4100"####),
            (r####"fixed_3"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 2) As Byte
Dim target As String
data(0) = 65
data(1) = 0
data(2) = 66
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:1|LenB:3|Units:0041|Bytes:410042"####),
            (r####"fixed_4"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 3) As Byte
Dim target As String
data(0) = 65
data(1) = 0
data(2) = 66
data(3) = 0
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:2|LenB:4|Units:00410042|Bytes:41004200"####),
            (r####"fixed_5"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 4) As Byte
Dim target As String
data(0) = 65
data(1) = 0
data(2) = 66
data(3) = 0
data(4) = 67
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:2|LenB:5|Units:00410042|Bytes:4100420043"####),
            (r####"negative_bounds"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(-2 To 1) As Byte
Dim target As String
data(-2) = 65
data(-1) = 0
data(0) = 66
data(1) = 0
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:2|LenB:4|Units:00410042|Bytes:41004200"####),
            (r####"two_dimensions"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 1, 0 To 1) As Byte
Dim target As String
data(0,0) = 65
data(1,0) = 0
data(0,1) = 66
data(1,1) = 0
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"ERR|13|OxiNamedProject|型が一致しません。|C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm|1000013"####),
            (r####"three_dimensions"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 1, 0 To 1, 0 To 0) As Byte
Dim target As String
data(0,0,0) = 65
data(1,0,0) = 0
data(0,1,0) = 66
data(1,1,0) = 0
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"ERR|13|OxiNamedProject|型が一致しません。|C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm|1000013"####),
            (r####"dynamic_unallocated"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data() As Byte
Dim target As String

target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:0|LenB:0|Units:|Bytes:"####),
            (r####"dynamic_erased"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data() As Byte
Dim target As String
ReDim data(0 To 3)
data(0) = 65
Erase data
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:0|LenB:0|Units:|Bytes:"####),
            (r####"boxed_byte"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 3) As Byte
Dim boxed As Variant
Dim target As String
data(0) = 65
data(2) = 66
boxed = data
target = boxed
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:2|LenB:4|Units:00410042|Bytes:41004200"####),
            (r####"boxed_variant_byte_elements"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 3) As Variant
Dim boxed As Variant
Dim target As String
data(0) = CByte(65)
data(1) = CByte(0)
data(2) = CByte(66)
data(3) = CByte(0)
boxed = data
target = boxed
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"ERR|13|OxiNamedProject|型が一致しません。|C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm|1000013"####),
            (r####"unpaired_surrogate"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 1) As Byte
Dim target As String
data(0) = 0
data(1) = 216
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:1|LenB:2|Units:D800|Bytes:00D8"####),
            (r####"odd_unpaired"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 2) As Byte
Dim target As String
data(0) = 0
data(1) = 216
data(2) = 65
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:1|LenB:3|Units:D800|Bytes:00D841"####),
            (r####"fixed_string_target"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 3) As Byte
Dim target As String * 1
data(0) = 65
data(2) = 66
target = data
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:1|LenB:2|Units:0041|Bytes:4100"####),
            (r####"explicit_cstr"####, r####"Attribute VB_Name = "ByteStringBoundary"
Option Explicit
Public Function Probe() As String
On Error GoTo Failed
Dim data(0 To 3) As Byte
Dim target As String
data(0) = 65
data(2) = 66
target = CStr(data)
Probe = "OK|Len:" & CStr(Len(target)) & "|LenB:" & CStr(LenB(target)) & "|Units:" & Units(target) & "|Bytes:" & RawBytes(target)
Exit Function
Failed:
Probe = "ERR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"####, r####"OK|Len:2|LenB:4|Units:00410042|Bytes:41004200"####)
        ];
        let environment = crate::ErrorEnvironment { language:crate::ErrorLanguage::Japanese, help_file:r"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm".to_owned() };
        let mut failures = Vec::new();
        for (name,source,expected) in cases {
            let module = parse_module(source).unwrap();
            let actual = Runtime::new(&module).with_project_name("OxiNamedProject").with_error_environment(environment.clone()).call("Probe",vec![]);
            if actual != Ok(Value::String(expected.to_owned())) { failures.push(format!("{name}: {actual:?} expected={expected}")); }
        }
        assert!(failures.is_empty(), "{}", failures.join("\n"));
    }

    #[test]
    fn odd_and_normal_text_comparisons_match_excel() {
        let source = r####"Attribute VB_Name = "OddStringComparison"
Option Explicit
Option Compare Text
Public Function Probe() As String
Dim data() As Byte
Dim values(0 To 9) As String
Dim i As Long
Dim j As Long
For i = 0 To 2
If i = 0 Then
ReDim data(0 To 0)
data(0) = 65
Else
ReDim data(0 To 2)
If i = 1 Then
data(0) = 65
data(1) = 0
data(2) = 66
Else
data(0) = 0
data(1) = 216
data(2) = 65
End If
End If
values(i) = data
Next i
values(3) = ""
values(4) = "A"
values(5) = "B"
values(6) = ChrW(&H100)
values(7) = ChrW(&HD800)
values(8) = "A" & ChrW(0)
values(9) = ChrW(0)
For i = 0 To 9
For j = 0 To 9
Probe = Probe & CStr(i) & "." & CStr(j) & ":" & CStr(StrComp(values(i), values(j), vbTextCompare)) & "," & CStr(values(i) = values(j)) & "," & CStr(values(i) < values(j)) & "," & CStr(values(i) > values(j)) & ";"
Next j
Next i
End Function
"####;
        let expected = r####"0.0:0,True,False,False;0.1:-1,False,True,False;0.2:-1,False,True,False;0.3:0,True,False,False;0.4:-1,False,True,False;0.5:-1,False,True,False;0.6:-1,False,True,False;0.7:-1,False,True,False;0.8:-1,False,True,False;0.9:0,True,False,False;1.0:1,False,False,True;1.1:0,True,False,False;1.2:-1,False,True,False;1.3:1,False,False,True;1.4:0,True,False,False;1.5:-1,False,True,False;1.6:-1,False,True,False;1.7:-1,False,True,False;1.8:0,True,False,False;1.9:1,False,False,True;2.0:1,False,False,True;2.1:1,False,False,True;2.2:0,True,False,False;2.3:1,False,False,True;2.4:1,False,False,True;2.5:1,False,False,True;2.6:1,False,False,True;2.7:0,True,False,False;2.8:1,False,False,True;2.9:1,False,False,True;3.0:0,True,False,False;3.1:-1,False,True,False;3.2:-1,False,True,False;3.3:0,True,False,False;3.4:-1,False,True,False;3.5:-1,False,True,False;3.6:-1,False,True,False;3.7:-1,False,True,False;3.8:-1,False,True,False;3.9:-1,False,True,False;4.0:1,False,False,True;4.1:0,True,False,False;4.2:-1,False,True,False;4.3:1,False,False,True;4.4:0,True,False,False;4.5:-1,False,True,False;4.6:-1,False,True,False;4.7:-1,False,True,False;4.8:0,True,False,False;4.9:1,False,False,True;5.0:1,False,False,True;5.1:1,False,False,True;5.2:-1,False,True,False;5.3:1,False,False,True;5.4:1,False,False,True;5.5:0,True,False,False;5.6:1,False,False,True;5.7:-1,False,True,False;5.8:1,False,False,True;5.9:1,False,False,True;6.0:1,False,False,True;6.1:1,False,False,True;6.2:-1,False,True,False;6.3:1,False,False,True;6.4:1,False,False,True;6.5:-1,False,True,False;6.6:0,True,False,False;6.7:-1,False,True,False;6.8:1,False,False,True;6.9:1,False,False,True;7.0:1,False,False,True;7.1:1,False,False,True;7.2:0,True,False,False;7.3:1,False,False,True;7.4:1,False,False,True;7.5:1,False,False,True;7.6:1,False,False,True;7.7:0,True,False,False;7.8:1,False,False,True;7.9:1,False,False,True;8.0:1,False,False,True;8.1:0,True,False,False;8.2:-1,False,True,False;8.3:1,False,False,True;8.4:0,True,False,False;8.5:-1,False,True,False;8.6:-1,False,True,False;8.7:-1,False,True,False;8.8:0,True,False,False;8.9:1,False,False,True;9.0:0,True,False,False;9.1:-1,False,True,False;9.2:-1,False,True,False;9.3:1,False,False,True;9.4:-1,False,True,False;9.5:-1,False,True,False;9.6:-1,False,True,False;9.7:-1,False,True,False;9.8:-1,False,True,False;9.9:0,True,False,False;"####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe",vec![]);
        assert_eq!(actual,Ok(Value::String(expected.to_owned())));
    }

    #[test]
    fn empty_and_nul_text_comparisons_match_excel() {
        let source = r#####"Attribute VB_Name = "EmptyStringComparison"
Option Explicit
Option Compare Text
Public Function Probe() As String
Dim values(0 To 9) As String
Dim bytes() As Byte
Dim i As Long
Dim j As Long
values(0) = ""
values(1) = vbNullString
values(3) = Space$(0)
values(4) = String$(0, "A")
values(5) = Left$("A", 0)
values(6) = bytes
ReDim bytes(0 To 0)
bytes(0) = 65
values(7) = bytes
values(8) = ChrW(0)
values(9) = ChrW(0) & ChrW(0)
For i = 0 To 9
Probe = Probe & "H" & CStr(i) & ":" & CStr(Len(values(i))) & "," & CStr(LenB(values(i))) & "," & CStr(StrPtr(values(i)) = 0) & ";"
Next i
For i = 0 To 9
For j = 0 To 9
Probe = Probe & CStr(i) & "." & CStr(j) & ":" & CStr(StrComp(values(i), values(j), vbTextCompare)) & "," & CStr(values(i) = values(j)) & "," & CStr(values(i) < values(j)) & "," & CStr(values(i) > values(j)) & ";"
Next j
Next i
End Function
"#####;
        let expected = r#####"H0:0,0,False;H1:0,0,True;H2:0,0,True;H3:0,0,False;H4:0,0,False;H5:0,0,False;H6:0,0,True;H7:0,1,False;H8:1,2,False;H9:2,4,False;0.0:0,True,False,False;0.1:0,True,False,False;0.2:0,True,False,False;0.3:0,True,False,False;0.4:0,True,False,False;0.5:0,True,False,False;0.6:0,True,False,False;0.7:0,True,False,False;0.8:-1,False,True,False;0.9:-1,False,True,False;1.0:0,True,False,False;1.1:0,True,False,False;1.2:0,True,False,False;1.3:0,True,False,False;1.4:0,True,False,False;1.5:0,True,False,False;1.6:0,True,False,False;1.7:0,True,False,False;1.8:-1,False,True,False;1.9:-1,False,True,False;2.0:0,True,False,False;2.1:0,True,False,False;2.2:0,True,False,False;2.3:0,True,False,False;2.4:0,True,False,False;2.5:0,True,False,False;2.6:0,True,False,False;2.7:0,True,False,False;2.8:-1,False,True,False;2.9:-1,False,True,False;3.0:0,True,False,False;3.1:0,True,False,False;3.2:0,True,False,False;3.3:0,True,False,False;3.4:0,True,False,False;3.5:0,True,False,False;3.6:0,True,False,False;3.7:0,True,False,False;3.8:-1,False,True,False;3.9:-1,False,True,False;4.0:0,True,False,False;4.1:0,True,False,False;4.2:0,True,False,False;4.3:0,True,False,False;4.4:0,True,False,False;4.5:0,True,False,False;4.6:0,True,False,False;4.7:0,True,False,False;4.8:-1,False,True,False;4.9:-1,False,True,False;5.0:0,True,False,False;5.1:0,True,False,False;5.2:0,True,False,False;5.3:0,True,False,False;5.4:0,True,False,False;5.5:0,True,False,False;5.6:0,True,False,False;5.7:0,True,False,False;5.8:-1,False,True,False;5.9:-1,False,True,False;6.0:0,True,False,False;6.1:0,True,False,False;6.2:0,True,False,False;6.3:0,True,False,False;6.4:0,True,False,False;6.5:0,True,False,False;6.6:0,True,False,False;6.7:0,True,False,False;6.8:-1,False,True,False;6.9:-1,False,True,False;7.0:0,True,False,False;7.1:0,True,False,False;7.2:0,True,False,False;7.3:0,True,False,False;7.4:0,True,False,False;7.5:0,True,False,False;7.6:0,True,False,False;7.7:0,True,False,False;7.8:0,True,False,False;7.9:0,True,False,False;8.0:1,False,False,True;8.1:1,False,False,True;8.2:1,False,False,True;8.3:1,False,False,True;8.4:1,False,False,True;8.5:1,False,False,True;8.6:1,False,False,True;8.7:0,True,False,False;8.8:0,True,False,False;8.9:0,True,False,False;9.0:1,False,False,True;9.1:1,False,False,True;9.2:1,False,False,True;9.3:1,False,False,True;9.4:1,False,False,True;9.5:1,False,False,True;9.6:1,False,False,True;9.7:0,True,False,False;9.8:0,True,False,False;9.9:0,True,False,False;"#####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe",vec![]);
        assert_eq!(actual,Ok(Value::String(expected.to_owned())));
    }

    #[test]
    fn byte_builtins_variant_inputs_match_excel() {
        let source = r#####"Attribute VB_Name = "ByteBuiltinVariants"
Option Explicit
Public Function Probe() As String
Dim value As Variant
Dim bytes() As Byte
Dim odd As String
Dim sample As Long
Dim operation As Long
ReDim bytes(0 To 0)
bytes(0) = 65
odd = bytes
For sample = 0 To 16
Select Case sample
Case 0: value = 12
Case 1: value = CByte(12)
Case 2: value = CLng(12)
Case 3: value = CSng(12)
Case 4: value = CDbl(12)
Case 5: value = CCur(12)
Case 6: value = CDec(12)
Case 7: value = CDate(12)
Case 8: value = True
Case 9: value = Empty
Case 10: value = Null
Case 11: value = CVErr(2042)
Case 12: value = "AB"
Case 13: value = vbNullString
Case 14: value = odd
Case 15: value = bytes
Case 16: Set value = Nothing
End Select
For operation = 1 To 20
Probe = Probe & CStr(sample) & "." & CStr(operation) & ":" & Action(value, operation) & ";"
Next operation
Next sample
End Function
Private Function Action(ByVal value As Variant, ByVal operation As Long) As String
On Error GoTo Failed
Dim result As Variant
Select Case operation
Case 1: result = LeftB(value, 1)
Case 2: result = RightB(value, 1)
Case 3: result = MidB(value, 1, 1)
Case 4: result = AscB(value)
Case 5: result = LenB(value)
Case 6: result = InStrB(1, value, "2", vbBinaryCompare)
Case 7: result = InStrB(1, value, "2", vbTextCompare)
Case 8: result = LeftB(value, 0)
Case 9: result = MidB(value, 0, 1)
Case 10: result = MidB(value, 1, -1)
Case 11: result = RightB(value, -1)
Case 12: result = Len(value)
Case 13: result = InStrB(value, "2")
Case 14: result = InStrB("12", value)
Case 15: result = LeftB$(value, 1)
Case 16: result = RightB$(value, 1)
Case 17: result = MidB$(value, 1, 1)
Case 18: result = AscW(value)
Case 19: result = Left$(value, 1)
Case 20: result = Left(value, 1)
End Select
If IsNull(result) Then
Action = "NULL"
ElseIf VarType(result) = vbString Then
Action = "S|L" & CStr(Len(result)) & "B" & CStr(LenB(result)) & "U" & Units(CStr(result)) & "R" & RawBytes(CStr(result))
Else
Action = "N|" & TypeName(result) & "|" & CStr(VarType(result)) & "|" & CStr(result)
End If
Exit Function
Failed:
Action = "ERR|" & CStr(Err.Number)
End Function
Private Function Units(ByVal value As String) As String
Dim i As Long
For i = 1 To Len(value)
Units = Units & Right$("0000" & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&), 4)
Next i
End Function
Private Function RawBytes(ByVal value As String) As String
Dim i As Long
For i = 1 To LenB(value)
RawBytes = RawBytes & Right$("00" & Hex$(AscB(MidB$(value, i, 1))), 2)
Next i
End Function
"#####;
        let expected = r#####"0.1:S|L0B1UR31;0.2:S|L0B1UR00;0.3:S|L0B1UR31;0.4:N|Byte|17|49;0.5:N|Long|3|4;0.6:N|Long|3|3;0.7:N|Long|3|3;0.8:S|L0B0UR;0.9:ERR|5;0.10:ERR|5;0.11:ERR|5;0.12:N|Long|3|2;0.13:N|Long|3|3;0.14:N|Long|3|1;0.15:S|L0B1UR31;0.16:S|L0B1UR00;0.17:S|L0B1UR31;0.18:N|Integer|2|49;0.19:S|L1B2U0031R3100;0.20:S|L1B2U0031R3100;1.1:S|L0B1UR31;1.2:S|L0B1UR00;1.3:S|L0B1UR31;1.4:N|Byte|17|49;1.5:N|Long|3|4;1.6:N|Long|3|3;1.7:N|Long|3|3;1.8:S|L0B0UR;1.9:ERR|5;1.10:ERR|5;1.11:ERR|5;1.12:N|Long|3|2;1.13:N|Long|3|3;1.14:N|Long|3|1;1.15:S|L0B1UR31;1.16:S|L0B1UR00;1.17:S|L0B1UR31;1.18:N|Integer|2|49;1.19:S|L1B2U0031R3100;1.20:S|L1B2U0031R3100;2.1:S|L0B1UR31;2.2:S|L0B1UR00;2.3:S|L0B1UR31;2.4:N|Byte|17|49;2.5:N|Long|3|4;2.6:N|Long|3|3;2.7:N|Long|3|3;2.8:S|L0B0UR;2.9:ERR|5;2.10:ERR|5;2.11:ERR|5;2.12:N|Long|3|2;2.13:N|Long|3|3;2.14:N|Long|3|1;2.15:S|L0B1UR31;2.16:S|L0B1UR00;2.17:S|L0B1UR31;2.18:N|Integer|2|49;2.19:S|L1B2U0031R3100;2.20:S|L1B2U0031R3100;3.1:S|L0B1UR31;3.2:S|L0B1UR00;3.3:S|L0B1UR31;3.4:N|Byte|17|49;3.5:N|Long|3|4;3.6:N|Long|3|3;3.7:N|Long|3|3;3.8:S|L0B0UR;3.9:ERR|5;3.10:ERR|5;3.11:ERR|5;3.12:N|Long|3|2;3.13:N|Long|3|3;3.14:N|Long|3|1;3.15:S|L0B1UR31;3.16:S|L0B1UR00;3.17:S|L0B1UR31;3.18:N|Integer|2|49;3.19:S|L1B2U0031R3100;3.20:S|L1B2U0031R3100;4.1:S|L0B1UR31;4.2:S|L0B1UR00;4.3:S|L0B1UR31;4.4:N|Byte|17|49;4.5:N|Long|3|4;4.6:N|Long|3|3;4.7:N|Long|3|3;4.8:S|L0B0UR;4.9:ERR|5;4.10:ERR|5;4.11:ERR|5;4.12:N|Long|3|2;4.13:N|Long|3|3;4.14:N|Long|3|1;4.15:S|L0B1UR31;4.16:S|L0B1UR00;4.17:S|L0B1UR31;4.18:N|Integer|2|49;4.19:S|L1B2U0031R3100;4.20:S|L1B2U0031R3100;5.1:S|L0B1UR31;5.2:S|L0B1UR00;5.3:S|L0B1UR31;5.4:N|Byte|17|49;5.5:N|Long|3|4;5.6:N|Long|3|3;5.7:N|Long|3|3;5.8:S|L0B0UR;5.9:ERR|5;5.10:ERR|5;5.11:ERR|5;5.12:N|Long|3|2;5.13:N|Long|3|3;5.14:N|Long|3|1;5.15:S|L0B1UR31;5.16:S|L0B1UR00;5.17:S|L0B1UR31;5.18:N|Integer|2|49;5.19:S|L1B2U0031R3100;5.20:S|L1B2U0031R3100;6.1:S|L0B1UR31;6.2:S|L0B1UR00;6.3:S|L0B1UR31;6.4:N|Byte|17|49;6.5:N|Long|3|4;6.6:N|Long|3|3;6.7:N|Long|3|3;6.8:S|L0B0UR;6.9:ERR|5;6.10:ERR|5;6.11:ERR|5;6.12:N|Long|3|2;6.13:N|Long|3|3;6.14:N|Long|3|1;6.15:S|L0B1UR31;6.16:S|L0B1UR00;6.17:S|L0B1UR31;6.18:N|Integer|2|49;6.19:S|L1B2U0031R3100;6.20:S|L1B2U0031R3100;7.1:S|L0B1UR31;7.2:S|L0B1UR00;7.3:S|L0B1UR31;7.4:N|Byte|17|49;7.5:N|Long|3|18;7.6:N|Long|3|0;7.7:N|Long|3|0;7.8:S|L0B0UR;7.9:ERR|5;7.10:ERR|5;7.11:ERR|5;7.12:N|Long|3|9;7.13:N|Long|3|0;7.14:N|Long|3|0;7.15:S|L0B1UR31;7.16:S|L0B1UR00;7.17:S|L0B1UR31;7.18:N|Integer|2|49;7.19:S|L1B2U0031R3100;7.20:S|L1B2U0031R3100;8.1:S|L0B1UR54;8.2:S|L0B1UR00;8.3:S|L0B1UR54;8.4:N|Byte|17|84;8.5:N|Long|3|8;8.6:N|Long|3|0;8.7:N|Long|3|0;8.8:S|L0B0UR;8.9:ERR|5;8.10:ERR|5;8.11:ERR|5;8.12:N|Long|3|4;8.13:N|Long|3|0;8.14:N|Long|3|0;8.15:S|L0B1UR54;8.16:S|L0B1UR00;8.17:S|L0B1UR54;8.18:N|Integer|2|84;8.19:S|L1B2U0054R5400;8.20:S|L1B2U0054R5400;9.1:S|L0B0UR;9.2:S|L0B0UR;9.3:S|L0B0UR;9.4:ERR|5;9.5:N|Long|3|0;9.6:N|Long|3|0;9.7:N|Long|3|0;9.8:S|L0B0UR;9.9:ERR|5;9.10:ERR|5;9.11:ERR|5;9.12:N|Long|3|0;9.13:N|Long|3|0;9.14:N|Long|3|1;9.15:S|L0B0UR;9.16:S|L0B0UR;9.17:S|L0B0UR;9.18:ERR|5;9.19:S|L0B0UR;9.20:S|L0B0UR;10.1:NULL;10.2:NULL;10.3:NULL;10.4:ERR|94;10.5:NULL;10.6:NULL;10.7:NULL;10.8:NULL;10.9:NULL;10.10:NULL;10.11:NULL;10.12:NULL;10.13:NULL;10.14:NULL;10.15:ERR|94;10.16:ERR|94;10.17:ERR|94;10.18:ERR|94;10.19:ERR|94;10.20:NULL;11.1:ERR|13;11.2:ERR|13;11.3:ERR|13;11.4:ERR|13;11.5:ERR|13;11.6:ERR|13;11.7:ERR|13;11.8:ERR|13;11.9:ERR|13;11.10:ERR|13;11.11:ERR|13;11.12:ERR|13;11.13:ERR|13;11.14:ERR|13;11.15:ERR|13;11.16:ERR|13;11.17:ERR|13;11.18:ERR|13;11.19:ERR|13;11.20:ERR|13;12.1:S|L0B1UR41;12.2:S|L0B1UR00;12.3:S|L0B1UR41;12.4:N|Byte|17|65;12.5:N|Long|3|4;12.6:N|Long|3|0;12.7:N|Long|3|0;12.8:S|L0B0UR;12.9:ERR|5;12.10:ERR|5;12.11:ERR|5;12.12:N|Long|3|2;12.13:N|Long|3|0;12.14:N|Long|3|0;12.15:S|L0B1UR41;12.16:S|L0B1UR00;12.17:S|L0B1UR41;12.18:N|Integer|2|65;12.19:S|L1B2U0041R4100;12.20:S|L1B2U0041R4100;13.1:S|L0B0UR;13.2:S|L0B0UR;13.3:S|L0B0UR;13.4:ERR|5;13.5:N|Long|3|0;13.6:N|Long|3|0;13.7:N|Long|3|0;13.8:S|L0B0UR;13.9:ERR|5;13.10:ERR|5;13.11:ERR|5;13.12:N|Long|3|0;13.13:N|Long|3|0;13.14:N|Long|3|1;13.15:S|L0B0UR;13.16:S|L0B0UR;13.17:S|L0B0UR;13.18:ERR|5;13.19:S|L0B0UR;13.20:S|L0B0UR;14.1:S|L0B1UR41;14.2:S|L0B1UR41;14.3:S|L0B1UR41;14.4:N|Byte|17|65;14.5:N|Long|3|1;14.6:N|Long|3|0;14.7:N|Long|3|0;14.8:S|L0B0UR;14.9:ERR|5;14.10:ERR|5;14.11:ERR|5;14.12:N|Long|3|0;14.13:N|Long|3|0;14.14:N|Long|3|0;14.15:S|L0B1UR41;14.16:S|L0B1UR41;14.17:S|L0B1UR41;14.18:ERR|5;14.19:S|L0B1UR41;14.20:S|L0B1UR41;15.1:S|L0B1UR41;15.2:S|L0B1UR41;15.3:S|L0B1UR41;15.4:N|Byte|17|65;15.5:N|Long|3|1;15.6:N|Long|3|0;15.7:N|Long|3|0;15.8:S|L0B0UR;15.9:ERR|5;15.10:ERR|5;15.11:ERR|5;15.12:N|Long|3|0;15.13:N|Long|3|0;15.14:N|Long|3|0;15.15:S|L0B1UR41;15.16:S|L0B1UR41;15.17:S|L0B1UR41;15.18:ERR|5;15.19:S|L0B1UR41;15.20:S|L0B1UR41;16.1:ERR|91;16.2:ERR|91;16.3:ERR|91;16.4:ERR|91;16.5:ERR|91;16.6:ERR|91;16.7:ERR|91;16.8:ERR|91;16.9:ERR|91;16.10:ERR|91;16.11:ERR|91;16.12:ERR|91;16.13:ERR|91;16.14:ERR|91;16.15:ERR|91;16.16:ERR|91;16.17:ERR|91;16.18:ERR|91;16.19:ERR|91;16.20:ERR|91;"#####;
        let module = parse_module(source).unwrap();
        let actual = Runtime::new(&module).with_project_name("OxiNamedProject").call("Probe",vec![]);
        assert_eq!(actual,Ok(Value::String(expected.to_owned())));
    }

}
