//! PDF stream filter decompression.

use crate::error::PdfError;
use super::object::PdfObject;
use flate2::read::ZlibDecoder;
use std::io::Read;

/// Decode stream data according to the /Filter entry in the stream dictionary.
pub fn decode_stream(dict: &[(String, PdfObject)], data: &[u8]) -> Result<Vec<u8>, PdfError> {
    let filter = dict
        .iter()
        .find(|(k, _)| k == "Filter")
        .map(|(_, v)| v);

    match filter {
        None => Ok(data.to_vec()),
        Some(PdfObject::Name(name)) => apply_filter(name, data),
        Some(PdfObject::Array(filters)) => {
            let mut result = data.to_vec();
            for f in filters {
                if let PdfObject::Name(name) = f {
                    result = apply_filter(name, &result)?;
                }
            }
            Ok(result)
        }
        _ => Ok(data.to_vec()),
    }
}

fn apply_filter(name: &str, data: &[u8]) -> Result<Vec<u8>, PdfError> {
    match name {
        "FlateDecode" | "Fl" => flate_decode(data),
        "ASCIIHexDecode" | "AHx" => ascii_hex_decode(data),
        "ASCII85Decode" | "A85" => {
            Err(PdfError::Unsupported("ASCII85Decode not yet implemented".into()))
        }
        "LZWDecode" | "LZW" => {
            Err(PdfError::Unsupported("LZWDecode not yet implemented".into()))
        }
        "DCTDecode" | "DCT" => {
            // JPEG — return raw data, consumer handles decoding.
            Ok(data.to_vec())
        }
        other => Err(PdfError::Unsupported(format!("filter: {other}"))),
    }
}

/// Maximum decompressed stream size (256 MB).
const MAX_DECOMPRESSED_SIZE: usize = 256 * 1024 * 1024;

fn flate_decode(data: &[u8]) -> Result<Vec<u8>, PdfError> {
    let mut decoder = ZlibDecoder::new(data);
    let mut result = Vec::new();
    let mut buf = [0u8; 8192];
    loop {
        let n = decoder
            .read(&mut buf)
            .map_err(|e| PdfError::Parse(format!("FlateDecode failed: {e}")))?;
        if n == 0 {
            break;
        }
        result.extend_from_slice(&buf[..n]);
        if result.len() > MAX_DECOMPRESSED_SIZE {
            return Err(PdfError::Parse(format!(
                "decompressed stream exceeds {} MB limit (possible zip bomb)",
                MAX_DECOMPRESSED_SIZE / (1024 * 1024)
            )));
        }
    }
    Ok(result)
}

fn ascii_hex_decode(data: &[u8]) -> Result<Vec<u8>, PdfError> {
    let hex: Vec<u8> = data
        .iter()
        .filter(|b| !b.is_ascii_whitespace() && **b != b'>')
        .copied()
        .collect();
    let mut result = Vec::with_capacity(hex.len() / 2);
    for pair in hex.chunks(2) {
        let s = std::str::from_utf8(pair)
            .map_err(|_| PdfError::Parse("invalid hex in ASCIIHexDecode".into()))?;
        let byte = u8::from_str_radix(s, 16)
            .map_err(|_| PdfError::Parse(format!("invalid hex byte: {s}")))?;
        result.push(byte);
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_no_filter() {
        let dict = vec![];
        let data = b"hello";
        let result = decode_stream(&dict, data).unwrap();
        assert_eq!(result, b"hello");
    }

    #[test]
    fn test_flate_decode() {
        use flate2::write::ZlibEncoder;
        use flate2::Compression;
        use std::io::Write;

        let original = b"Hello, PDF world! This is a test of FlateDecode.";
        let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
        encoder.write_all(original).unwrap();
        let compressed = encoder.finish().unwrap();

        let dict = vec![("Filter".to_string(), PdfObject::Name("FlateDecode".to_string()))];
        let result = decode_stream(&dict, &compressed).unwrap();
        assert_eq!(result, original);
    }

    #[test]
    fn test_ascii_hex_decode() {
        let dict = vec![("Filter".to_string(), PdfObject::Name("ASCIIHexDecode".to_string()))];
        let data = b"48 65 6C 6C 6F>";
        let result = decode_stream(&dict, data).unwrap();
        assert_eq!(result, b"Hello");
    }
}
