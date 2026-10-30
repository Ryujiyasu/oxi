// SPDX-License-Identifier: MIT OR Apache-2.0

//! The members of the drawing objects: what `Shape`, `TextFrame`, `Chart`,
//! `Series` and the rest answer and take. Defaults and refusals as measured
//! (shapes.vba, shapes2.vba, 2026-09-05).

use super::shapes::*;
use super::*;
use oxicells_core::ir::{DrawingKind, Frame, GroupTransform};
use std::io::{Cursor, Read, Write};

const EMU_PER_POINT: f64 = 12_700.0;

/// Excel snaps both edges of a non-icon Word OLE Add extent to its 96-DPI
/// pixel grid. The pixel span then uses a 1.5pt base step and an alternating
/// one-pixel correction in the latter 11 pixels of each 24-pixel period.
fn word_ole_add_width(requested: f64, left: f64) -> Option<f64> {
    if !requested.is_finite() || !left.is_finite() {
        return None;
    }
    // Measuring the span between independently snapped edges explains both
    // negative widths (the far edge crosses left) and Left-sensitive results.
    let start_pixel = (left / 0.75).floor();
    let end_pixel = ((left + requested) / 0.75).floor();
    let pixel = (end_pixel - start_pixel).abs();
    if !pixel.is_finite() || pixel > i64::MAX as f64 {
        return None;
    }
    let pixel = pixel as i64;
    let phase = pixel.rem_euclid(24);
    let base = pixel.div_euclid(2) * 2;
    let extent = if phase < 13 {
        base
    } else if (phase - 13).rem_euclid(2) == 0 {
        base + 1
    } else {
        base - 1
    };
    Some(extent.max(1) as f64 * 0.75)
}

fn word_ole_add_dimension(value: &Value) -> Result<Option<f64>, ()> {
    let number = match value {
        Value::Missing => return Ok(None),
        Value::Empty | Value::Null => 0.0,
        // VBA's VARIANT_BOOL uses -1 for True when Excel coerces it to a number.
        Value::Boolean(value) => if *value { -1.0 } else { 0.0 },
        Value::Decimal(value) => value.to_f64(),
        Value::String(text) => return Ok(Some(text.trim().parse::<f64>().ok().filter(|number| number.is_finite()).ok_or(())?)),
        value => any_number(value).ok_or(())?,
    };
    number.is_finite().then_some(Some(number)).ok_or(())
}

fn word_ole_display_as_icon(value: &Value) -> Result<bool, ()> {
    match value {
        Value::Missing | Value::Empty | Value::Null => Ok(false),
        Value::Boolean(value) => Ok(*value),
        Value::String(value) if value.eq_ignore_ascii_case("true") => Ok(true),
        Value::String(value) if value.eq_ignore_ascii_case("false") => Ok(false),
        Value::String(_) => Err(()),
        value => Ok(word_ole_add_dimension(value)?.unwrap_or(0.0) != 0.0),
    }
}

fn word_document_with_ole_width(width: f64) -> Result<Vec<u8>, String> {
    let page_width_twips = if width <= 0.75 {
        3_405
    } else {
        let page_width_twips = 3_405.0 + width * 20.0;
        if !page_width_twips.is_finite() || page_width_twips > u32::MAX as f64 {
            return Err("Word document width is too large".into());
        }
        page_width_twips.round() as u32
    };
    let source = include_bytes!("data/word_document_empty.docx");
    let mut archive = zip::ZipArchive::new(Cursor::new(source.as_slice()))
        .map_err(|error| format!("cannot read blank Word document: {error}"))?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut replaced_page_size = false;
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(|error| error.to_string())?;
        let name = file.name().to_owned();
        if name == "word/document.xml" {
            let mut xml = String::new();
            file.read_to_string(&mut xml).map_err(|error| error.to_string())?;
            let original = "w:w=\"12240\" w:h=\"15840\"";
            if xml.matches(original).count() != 1 {
                return Err("blank Word document page size did not match its template".into());
            }
            xml = xml.replace(original, &format!("w:w=\"{page_width_twips}\" w:h=\"15840\""));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(file.compression())
                .last_modified_time(file.last_modified().unwrap_or_default());
            output.start_file(name, options).map_err(|error| error.to_string())?;
            output.write_all(xml.as_bytes()).map_err(|error| error.to_string())?;
            replaced_page_size = true;
        } else {
            output.raw_copy_file(file).map_err(|error| error.to_string())?;
        }
    }
    if !replaced_page_size {
        return Err("blank Word document has no document body part".into());
    }
    Ok(output.finish().map_err(|error| error.to_string())?.into_inner())
}

#[cfg(test)]
mod word_ole_width_tests {
    use super::{word_ole_add_dimension, word_ole_add_width, word_ole_display_as_icon, Value};

    #[test]
    fn word_ole_add_width_matches_excel_pixel_quantization_samples() {
        for (requested, left, expected) in [
            (0.25, 0.0, 0.75),
            (1.5, 0.0, 1.5),
            (9.5, 0.0, 9.0),
            (9.75, 0.0, 9.75),
            (10.5, 0.0, 9.75),
            (11.25, 0.0, 11.25),
            (17.75, 0.0, 17.25),
            (18.0, 0.0, 18.0),
            (50.0, 0.0, 48.75),
            (100.0, 0.0, 99.75),
            (125.0, 0.0, 123.75),
            (588.0, 0.0, 587.25),
            (-0.1, 0.0, 0.75),
            (-0.75, 0.0, 0.75),
            (-1.0, 0.0, 1.5),
            (-2.0, 0.0, 1.5),
            (-5.0, 0.0, 4.5),
            (-10.0, 0.0, 9.75),
            (-30.0, 0.0, 29.25),
            (50.0, 0.25, 50.25),
            (50.0, 0.75, 48.75),
            (25.0, 0.5, 25.5),
            (25.0, 0.75, 24.0),
        ] {
            assert_eq!(word_ole_add_width(requested, left), Some(expected), "Width={requested}, Left={left}");
        }
        assert_eq!(word_ole_add_width(f64::NAN, 0.0), None);
        assert_eq!(word_ole_add_dimension(&Value::String("50".into())), Ok(Some(50.0)));
        assert_eq!(word_ole_add_dimension(&Value::Boolean(true)), Ok(Some(-1.0)));
        assert_eq!(word_ole_add_dimension(&Value::Boolean(false)), Ok(Some(0.0)));
        assert_eq!(word_ole_add_dimension(&Value::Empty), Ok(Some(0.0)));
        assert_eq!(word_ole_add_dimension(&Value::Null), Ok(Some(0.0)));
        assert_eq!(word_ole_add_dimension(&Value::String("bad".into())), Err(()));
        assert_eq!(word_ole_display_as_icon(&Value::String("TRUE".into())), Ok(true));
        assert_eq!(word_ole_display_as_icon(&Value::String("False".into())), Ok(false));
        assert_eq!(word_ole_display_as_icon(&Value::Integer(1)), Ok(true));
        assert_eq!(word_ole_display_as_icon(&Value::String("1".into())), Err(()));
        assert_eq!(word_ole_display_as_icon(&Value::String(" True ".into())), Err(()));
    }
}

fn package_embedding(filename: &str, source_path: &str, data: &[u8]) -> Result<Vec<u8>, String> {
    let encoding = codepage::to_encoding(932).ok_or_else(|| "Windows-932 encoding is unavailable".to_string())?;
    let ansi = |text: &str| encoding.encode(text).0.into_owned();
    let temp_path = source_path;
    let basename = filename.rsplit(['\\', '/']).next().unwrap_or(filename);
    let mut native = Vec::new();
    native.extend_from_slice(&[0; 4]);
    native.extend_from_slice(&2u16.to_le_bytes());
    native.extend_from_slice(&ansi(basename)); native.push(0);
    native.extend_from_slice(&ansi(source_path)); native.push(0);
    native.extend_from_slice(&0u16.to_le_bytes());
    native.extend_from_slice(&3u16.to_le_bytes());
    let temp_ansi = ansi(temp_path);
    native.extend_from_slice(&u32::try_from(temp_ansi.len() + 1).map_err(|_| "OLE file path is too long")?.to_le_bytes());
    native.extend_from_slice(&temp_ansi); native.push(0);
    native.extend_from_slice(&u32::try_from(data.len()).map_err(|_| "OLE file is too large")?.to_le_bytes());
    native.extend_from_slice(data);
    for value in [source_path, basename, temp_path] {
        let mut utf16: Vec<u16> = value.encode_utf16().collect(); utf16.push(0);
        native.extend_from_slice(&u32::try_from(utf16.len() * 2).map_err(|_| "OLE file path is too long")?.to_le_bytes());
        for unit in utf16 { native.extend_from_slice(&unit.to_le_bytes()); }
    }
    let payload_len = u32::try_from(native.len() - 4).map_err(|_| "OLE package is too large")?;
    native[..4].copy_from_slice(&payload_len.to_le_bytes());

    // Excel's Package server uses this CompObj descriptor and the Ole10Native
    // stream; the latter contains the caller's file bytes and original path.
    const COMP_OBJ: &[u8] = &[
        0x01,0x00,0xfe,0xff,0x03,0x0a,0x00,0x00,0xff,0xff,0xff,0xff,0x0c,0x00,0x03,0x00,
        0x00,0x00,0x00,0x00,0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46,0x0c,0x00,0x00,0x00,
        0x4f,0x4c,0x45,0x20,0x50,0x61,0x63,0x6b,0x61,0x67,0x65,0x00,0x00,0x00,0x00,0x00,
        0x08,0x00,0x00,0x00,0x50,0x61,0x63,0x6b,0x61,0x67,0x65,0x00,0xf4,0x39,0xb2,0x71,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
    ];
    let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
        .map_err(|error| format!("could not create OLE package: {error}"))?;
    compound.create_stream(std::path::Path::new("/\u{1}CompObj"))
        .and_then(|mut stream| stream.write_all(COMP_OBJ))
        .map_err(|error| format!("could not write OLE package descriptor: {error}"))?;
    compound.create_stream(std::path::Path::new("/\u{1}Ole10Native"))
        .and_then(|mut stream| stream.write_all(&native))
        .map_err(|error| format!("could not write OLE package data: {error}"))?;
    compound.set_storage_clsid("/", uuid::uuid!("0003000c-0000-0000-c000-000000000046"))
        .map_err(|error| format!("could not set OLE Package class id: {error}"))?;
    compound.flush().map_err(|error| format!("could not finalize OLE package: {error}"))?;
    Ok(compound.into_inner().into_inner())
}

fn empty_package_embedding() -> Result<Vec<u8>, String> {
    let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
        .map_err(|error| format!("could not create OLE package: {error}"))?;
    compound.set_storage_clsid("/", uuid::uuid!("0003000c-0000-0000-c000-000000000046"))
        .map_err(|error| format!("could not set OLE Package class id: {error}"))?;
    compound.flush().map_err(|error| format!("could not finalize OLE package: {error}"))?;
    Ok(compound.into_inner().into_inner())
}

/// E_INVALIDARG, which is what `Shapes("nope")` and `Shapes(9)` raise.
const NO_SUCH_SHAPE: i64 = -2_147_024_809;

/// Excel's new AutoShape outline points at DrawingML `lnRef` slot 2, whose
/// Office format scheme shades the shape's accent by 15% in linear light.
fn default_shape_outline(accent: i64) -> i64 {
    let encode = |light: f64| {
        let srgb = if light <= 0.003_130_8 {
            12.92 * light
        } else {
            1.055 * light.powf(1.0 / 2.4) - 0.055
        };
        (srgb * 255.0).round().clamp(0.0, 255.0) as i64
    };
    let shade = |byte: i64| {
        let srgb = byte as f64 / 255.0;
        let light = if srgb <= 0.040_45 {
            srgb / 12.92
        } else {
            ((srgb + 0.055) / 1.055).powf(2.4)
        };
        encode(light * 0.15)
    };
    shade(accent & 0xff) | (shade((accent >> 8) & 0xff) << 8) | (shade((accent >> 16) & 0xff) << 16)
}

fn excel_shape_transparency(value: f64) -> Result<f64, String> {
    // The type library declares these setters as Single. Excel narrows the
    // incoming Variant before the FillFormat/LineFormat range check, so a
    // Double just above 1 can round to the accepted Single value 1.
    let value = f64::from(value as f32);
    if !(0.0..=1.0).contains(&value) {
        return Err(host_error(5, "Transparency must be between 0 and 1"));
    }
    Ok(1.0 - ((1.0 - value) * 65_536.0).round() / 65_536.0)
}

/// The Crop object describes the uncropped picture in worksheet points. Excel
/// can scale an image independently of its file's physical DPI, so recover the
/// current picture extent from the displayed shape and the DrawingML crop
/// fractions instead of using the raster's intrinsic pixel size.
fn crop_picture_geometry(shape: &ShapeRecord) -> Result<(f64, f64, f64, f64, f64, f64), String> {
    let Some(DrawingKind::Picture { bytes, format, .. }) = shape.original.as_ref().map(|drawing| &drawing.kind) else {
        return Err(host_error(438, "Object doesn't support this property or method"));
    };
    let (intrinsic_w, intrinsic_h) = picture_source_size_points(bytes, format).unwrap_or((shape.width, shape.height));
    let fraction_w = 1.0 - f64::from(format.crop.0 + format.crop.2) / 100_000.0;
    let fraction_h = 1.0 - f64::from(format.crop.1 + format.crop.3) / 100_000.0;
    let inferred_w = if fraction_w.abs() > f64::EPSILON { shape.width / fraction_w } else { intrinsic_w };
    let inferred_h = if fraction_h.abs() > f64::EPSILON { shape.height / fraction_h } else { intrinsic_h };
    let picture_w = inferred_w;
    let picture_h = inferred_h;
    if !picture_w.is_finite() || !picture_h.is_finite() || picture_w <= 0.0 || picture_h <= 0.0 {
        return Err(host_error(1004, "Invalid picture dimensions"));
    }
    let left = f64::from(format.crop.0) * picture_w / 100_000.0;
    let top = f64::from(format.crop.1) * picture_h / 100_000.0;
    let right = f64::from(format.crop.2) * picture_w / 100_000.0;
    let bottom = f64::from(format.crop.3) * picture_h / 100_000.0;
    Ok((picture_w, picture_h, left, top, right, bottom))
}

fn excel_line_weight(value: f64) -> Result<f64, String> {
    // LineFormat.Weight is a Single measured in points. Excel stores the
    // narrowed value as an integral EMU count (12,700 EMUs per point).
    let value = f64::from(value as f32);
    if !value.is_finite() || value < 0.0 {
        return Err(host_error(5, "Line weight must be finite and nonnegative"));
    }
    Ok((value * 12_700.0).round() / 12_700.0)
}

fn excel_line_dash_style(value: &Value) -> Result<i64, String> {
    excel_enum_variant(value, 1, 12, "Invalid line dash style")
}

fn excel_arrowhead_style(value: &Value) -> Result<i64, String> {
    excel_enum_variant(value, 1, 6, "Invalid arrowhead style")
}

fn excel_enum_variant(value: &Value, minimum: i64, maximum: i64, message: &str) -> Result<i64, String> {
    // Excel coerces these Variant-backed enum arguments to Long with
    // banker's rounding, then validates the enum range.
    let number = transparency_number(value)?;
    if !number.is_finite() {
        return Err(host_error(5, message));
    }
    let rounded = number.round_ties_even();
    if !(minimum as f64..=maximum as f64).contains(&rounded) {
        return Err(host_error(5, message));
    }
    Ok(rounded as i64)
}

fn transparency_number(value: &Value) -> Result<f64, String> {
    match value {
        Value::Empty => Ok(0.0),
        Value::Null => Err(host_error(94, "Invalid use of Null")),
        Value::Boolean(true) => Ok(-1.0),
        Value::Boolean(false) => Ok(0.0),
        // Preserve IEEE values until the property-specific range check can
        // return Excel's argument error instead of a conversion Type mismatch.
        Value::Single(value) => Ok(f64::from(*value)),
        Value::Double(value) => Ok(*value),
        Value::Decimal(value) => Ok(value.to_f64()),
        other => numeric_text(other).ok_or_else(|| host_error(13, "Type mismatch")),
    }
}

/// The original physical extent of a raster picture, in points. Excel bases
/// ScaleWidth/ScaleHeight with RelativeToOriginalSize on the embedded image's
/// pixel density, not on its currently displayed drawing extent.
fn picture_original_size_points(bytes: &[u8]) -> Option<(f64, f64)> {
    fn be_u32(bytes: &[u8], at: usize) -> Option<u32> {
        Some(u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
    }
    fn le_u32(bytes: &[u8], at: usize) -> Option<u32> {
        Some(u32::from_le_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
    }
    fn image_points(width: u32, height: u32, dpi_x: f64, dpi_y: f64) -> Option<(f64, f64)> {
        if width == 0 || height == 0 || !dpi_x.is_finite() || !dpi_y.is_finite() || dpi_x <= 0.0 || dpi_y <= 0.0 {
            return None;
        }
        Some((f64::from(width) * 72.0 / dpi_x, f64::from(height) * 72.0 / dpi_y))
    }
    fn jpeg_exif_dpi(app1: &[u8]) -> Option<(f64, f64)> {
        if !app1.starts_with(b"Exif\0\0") { return None; }
        let tiff = app1.get(6..)?;
        let little = match tiff.get(..2)? {
            b"II" => true,
            b"MM" => false,
            _ => return None,
        };
        fn read_u16(bytes: &[u8], at: usize, little: bool) -> Option<u16> {
            let pair: [u8; 2] = bytes.get(at..at + 2)?.try_into().ok()?;
            Some(if little { u16::from_le_bytes(pair) } else { u16::from_be_bytes(pair) })
        }
        fn read_u32(bytes: &[u8], at: usize, little: bool) -> Option<u32> {
            let quad: [u8; 4] = bytes.get(at..at + 4)?.try_into().ok()?;
            Some(if little { u32::from_le_bytes(quad) } else { u32::from_be_bytes(quad) })
        }
        if read_u16(tiff, 2, little)? != 42 { return None; }
        let ifd = usize::try_from(read_u32(tiff, 4, little)?).ok()?;
        let count = usize::from(read_u16(tiff, ifd, little)?);
        let entries = ifd.checked_add(2)?;
        let mut x_resolution = None;
        let mut y_resolution = None;
        let mut resolution_unit = None;
        for index in 0..count {
            let at = entries.checked_add(index.checked_mul(12)?)?;
            let tag = read_u16(tiff, at, little)?;
            let kind = read_u16(tiff, at + 2, little)?;
            let values = read_u32(tiff, at + 4, little)?;
            match tag {
                0x011a | 0x011b if kind == 5 && values >= 1 => {
                    let value_at = usize::try_from(read_u32(tiff, at + 8, little)?).ok()?;
                    let numerator = f64::from(read_u32(tiff, value_at, little)?);
                    let denominator = f64::from(read_u32(tiff, value_at + 4, little)?);
                    if denominator <= 0.0 { continue; }
                    let value = numerator / denominator;
                    if tag == 0x011a { x_resolution = Some(value); } else { y_resolution = Some(value); }
                }
                0x0128 if kind == 3 && values >= 1 => {
                    resolution_unit = read_u16(tiff, at + 8, little);
                }
                _ => {}
            }
        }
        let unit = resolution_unit?;
        let (x, y) = (x_resolution?, y_resolution?);
        match unit {
            2 => Some((x, y)),
            3 => Some((x * 2.54, y * 2.54)),
            _ => None,
        }
    }

    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        let width = be_u32(bytes, 16)?;
        let height = be_u32(bytes, 20)?;
        let mut dpi = (96.0, 96.0);
        let mut at = 8usize;
        while at.checked_add(12)? <= bytes.len() {
            let length = usize::try_from(be_u32(bytes, at)?).ok()?;
            let end = at.checked_add(12)?.checked_add(length)?;
            if end > bytes.len() { return None; }
            if &bytes[at + 4..at + 8] == b"pHYs" && length == 9 && bytes[at + 16] == 1 {
                dpi = (f64::from(be_u32(bytes, at + 8)?) * 0.0254, f64::from(be_u32(bytes, at + 12)?) * 0.0254);
                break;
            }
            at = end;
        }
        return image_points(width, height, dpi.0, dpi.1);
    }

    if bytes.starts_with(b"BM") && bytes.len() >= 54 {
        let header = le_u32(bytes, 14)?;
        if header < 40 { return None; }
        let width = i32::from_le_bytes(bytes.get(18..22)?.try_into().ok()?).unsigned_abs();
        let height = i32::from_le_bytes(bytes.get(22..26)?.try_into().ok()?).unsigned_abs();
        let xppm = i32::from_le_bytes(bytes.get(38..42)?.try_into().ok()?);
        let yppm = i32::from_le_bytes(bytes.get(42..46)?.try_into().ok()?);
        let dpi_x = if xppm > 0 { f64::from(xppm) * 0.0254 } else { 96.0 };
        let dpi_y = if yppm > 0 { f64::from(yppm) * 0.0254 } else { 96.0 };
        return image_points(width, height, dpi_x, dpi_y);
    }

    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        let width = u32::from(u16::from_le_bytes(bytes.get(6..8)?.try_into().ok()?));
        let height = u32::from(u16::from_le_bytes(bytes.get(8..10)?.try_into().ok()?));
        return image_points(width, height, 96.0, 96.0);
    }

    if bytes.starts_with(&[0xff, 0xd8]) {
        let mut jfif_dpi = None;
        let mut exif_dpi = None;
        let mut dimensions = None;
        let mut at = 2usize;
        while at + 4 <= bytes.len() {
            if bytes[at] != 0xff { at += 1; continue; }
            while bytes.get(at) == Some(&0xff) { at += 1; }
            let marker = *bytes.get(at)?;
            at += 1;
            if matches!(marker, 0xd8 | 0xd9 | 0x01 | 0xd0..=0xd7) { continue; }
            let length = usize::from(u16::from_be_bytes(bytes.get(at..at + 2)?.try_into().ok()?));
            if length < 2 || at.checked_add(length)? > bytes.len() { return None; }
            let data = &bytes[at + 2..at + length];
            if marker == 0xe0 && data.starts_with(b"JFIF\0") && data.len() >= 12 {
                let unit = data[7];
                let x = f64::from(u16::from_be_bytes(data[8..10].try_into().ok()?));
                let y = f64::from(u16::from_be_bytes(data[10..12].try_into().ok()?));
                if unit == 1 { jfif_dpi = Some((x, y)); }
                else if unit == 2 { jfif_dpi = Some((x * 2.54, y * 2.54)); }
            }
            if marker == 0xe1 {
                exif_dpi = jpeg_exif_dpi(data).or(exif_dpi);
            }
            if matches!(marker, 0xc0..=0xc3 | 0xc5..=0xc7 | 0xc9..=0xcb | 0xcd..=0xcf) && data.len() >= 5 {
                let height = u32::from(u16::from_be_bytes(data[1..3].try_into().ok()?));
                let width = u32::from(u16::from_be_bytes(data[3..5].try_into().ok()?));
                dimensions = Some((width, height));
            }
            if marker == 0xda {
                break;
            }
            at += length;
        }
        if let Some((width, height)) = dimensions {
            let (dpi_x, dpi_y) = jfif_dpi.or(exif_dpi).unwrap_or((96.0, 96.0));
            return image_points(width, height, dpi_x, dpi_y);
        }
    }
    None
}

fn picture_source_size_points(bytes: &[u8], format: &oxicells_core::ir::PictureStyle) -> Option<(f64, f64)> {
    format.original_size_points.or_else(|| picture_original_size_points(bytes))
}

impl<'a> WorkbookHost<'a> {
    pub(super) fn comment_shape(&mut self, at: CellAddress) -> Result<Value, String> {
        let id = self.shapes.iter().find(|shape| shape.comment_at == Some(at)).map(|shape| shape.id);
        let id = match id {
            Some(id) => id,
            None => {
                self.add_comment_shape(at);
                self.shapes.iter().find(|shape| shape.comment_at == Some(at))
                    .map(|shape| shape.id).ok_or_else(|| host_error(1004, "the comment shape is unavailable"))?
            }
        };
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// What `Selection` is: a ShapeRange when a macro has selected shapes,
    /// otherwise the cells.
    pub(super) fn selection_object(&mut self) -> Value {
        let live: Vec<u64> = self.shape_selection.iter().copied().filter(|id| self.shape(*id).is_ok()).collect();
        if live.is_empty() {
            return self.object(HostObject::Range(self.selection));
        }
        if let [id] = live.as_slice() {
            if self.shape(*id).ok().is_some_and(|shape| shape.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some())) {
                return self.part(DrawingPart::OleObject(*id));
            }
        }
        self.shape_range_object(live)
    }

    fn shape_range_object(&mut self, ids: Vec<u64>) -> Value {
        let index = self.shape_ranges.len();
        self.shape_ranges.push(ids);
        self.part(DrawingPart::ShapeRange(index))
    }

    fn shape_range_ids(&self, index: usize) -> Vec<u64> {
        self.shape_ranges.get(index).cloned().unwrap_or_default()
    }

    /// `ActiveChart`: the chart of the selected shape when it holds one, or
    /// the sheet's one chart when only one is there. Measured: after
    /// `AddChart2(...).Select`, ActiveChart is that chart. None -> 91.
    pub(super) fn active_chart_object(&mut self) -> Result<Value, String> {
        let selected: Option<u64> = self
            .shape_selection
            .iter()
            .copied()
            .find(|id| matches!(self.shape(*id).map(|s| s.kind.clone()), Ok(ShapeKind::Chart(_))));
        let chart = selected.or_else(|| {
            let charts = self.charts_on(self.active_sheet);
            (charts.len() == 1).then(|| charts[0])
        });
        match chart {
            Some(id) => Ok(self.part(DrawingPart::Chart(id))),
            None => Err(host_error(91, "there is no active chart")),
        }
    }

    pub(super) fn drawing_part(&self, object: &ObjectRef) -> Option<DrawingPart> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::Drawing(part)) => Some(*part),
            _ => None,
        }
    }

    pub(super) fn part(&mut self, part: DrawingPart) -> Value {
        if let DrawingPart::Chart(id)=part {
            if let Ok(shape)=self.shape(id) {
                let sheet=shape.sheet;
                if self.workbook.sheets[sheet].kind==oxicells_core::ir::SheetKind::Chart {
                    return self.object(HostObject::Worksheet(sheet));
                }
            }
        }
        self.object(HostObject::Drawing(part))
    }

    pub(super) fn chart_sheet_part(&self,sheet:usize)->Result<DrawingPart,String> {
        self.shapes.iter().find(|shape|shape.sheet==sheet && matches!(shape.kind,ShapeKind::Chart(_)))
            .map(|shape|DrawingPart::Chart(shape.id))
            .ok_or_else(||host_error(438,"the chart data is not available"))
    }

    /// The shapes on a sheet in drawing order; Excel changes collection order
    /// when a shape moves to the front or back of the z-order.
    fn shapes_on(&self, sheet: usize) -> Vec<u64> {
        let mut listed = Vec::new();
        let mut seen_groups = std::collections::HashSet::new();
        for shape in self.shapes.iter().filter(|shape| {
            shape.sheet == sheet && shape.comment_at.is_none_or(|at| self.note_at(at).is_some())
        }) {
            if !shape.group_children.is_empty() {
                continue;
            }
            if let Some(parent) = shape.group_parent_id {
                if seen_groups.insert(parent) {
                    listed.push(parent);
                }
            } else {
                listed.push(shape.id);
            }
        }
        listed
    }

    pub(super) fn ole_objects_on(&self, sheet: usize) -> Vec<u64> {
        self.shapes_on(sheet).into_iter().filter(|id| {
            self.shape(*id).is_ok_and(|shape| {
                shape.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some())
                    || !shape.group_children.is_empty()
                        && shape.group_children.iter().all(|child| {
                            self.shape(*child).ok().and_then(|child| child.original.as_ref())
                                .is_some_and(|drawing| drawing.ole_prog_id.is_some())
                        })
            })
        }).collect()
    }

    pub(super) fn ole_objects_for_part(&self, part: DrawingPart) -> Vec<u64> {
        match part {
            DrawingPart::OleObjects(sheet) => self.ole_objects_on(sheet),
            DrawingPart::OleObjectsSnapshot(_, index) => self.shape_range_ids(index),
            _ => Vec::new(),
        }
    }

    pub(super) fn ole_objects_snapshot(&mut self, sheet: usize, ids: Vec<u64>) -> Value {
        let index = self.shape_ranges.len();
        self.shape_ranges.push(ids);
        self.part(DrawingPart::OleObjectsSnapshot(sheet, index))
    }

    fn pick_ole_object(&self, listed: &[u64], index: &Value) -> Result<u64, String> {
        let missing = || oxivba_core::host_error_described(
            1004,
            "OLEObjects クラスの Item プロパティを取得できません。",
        );
        let id = match index {
            Value::String(name) => listed.iter().copied()
                .find(|id| self.shape(*id).is_ok_and(|shape| shape.name.eq_ignore_ascii_case(name)))
                .ok_or_else(missing),
            value => any_whole_number(value)
                .filter(|number| *number >= 1 && (*number as usize) <= listed.len())
                .map(|number| listed[number as usize - 1])
                .ok_or_else(missing),
        }?;
        if self.shape(id).is_ok_and(|shape| !shape.group_children.is_empty()) { Err(missing()) } else { Ok(id) }
    }

    fn group_shapes(&mut self, sheet: usize, children: Vec<u64>) -> Result<u64, String> {
        if children.len() < 2 {
            return Err(host_error(1004, "This cannot be done to a multiple selection."));
        }
        let mut bounds = None::<(f64, f64, f64, f64)>;
        for id in &children {
            let shape = self.shape(*id)?;
            bounds = Some(match bounds {
                None => (shape.left, shape.top, shape.width, shape.height),
                Some((left, top, width, height)) => {
                    let right = (left + width).max(shape.left + shape.width);
                    let bottom = (top + height).max(shape.top + shape.height);
                    let left = left.min(shape.left);
                    let top = top.min(shape.top);
                    (left, top, right - left, bottom - top)
                }
            });
        }
        let bounds = bounds.ok_or_else(|| host_error(1004, "There are no shapes to group."))?;
        let mut number = 1usize;
        while self.shapes.iter().any(|shape| shape.sheet == sheet && shape.name.eq_ignore_ascii_case(&format!("Group {number}"))) {
            number += 1;
        }
        let name = format!("Group {number}");
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let mut group = ShapeRecord::blank(id, sheet, name.clone(), ShapeKind::Other);
        group.left = bounds.0;
        group.top = bounds.1;
        group.width = bounds.2;
        group.height = bounds.3;
        group.visible = children.iter().all(|child| self.shape(*child).is_ok_and(|shape| shape.visible));
        group.group_children = children.clone();
        group.group_bounds = Some(bounds);
        let emu = |points: f64| (points / 0.75 * 9_525.0).round() as i64;
        let offset = (emu(bounds.0), emu(bounds.1));
        let extent = (emu(bounds.2), emu(bounds.3));
        group.group_transform = Some(oxicells_core::ir::GroupTransform {
            offset,
            extent,
            child_offset: offset,
            child_extent: extent,
        });
        for child in &children {
            let child_shape = self.shape_mut(*child)?;
            child_shape.group_parent_id = Some(id);
            child_shape.group_name = Some(name.clone());
        }
        self.shapes.push(group);
        self.shape_selection = vec![id];
        self.mirror_drawings(sheet);
        Ok(id)
    }

    fn move_shape_in_z_order(&mut self, id: u64, to_front: bool) -> Result<(), String> {
        let sheet = self.shape(id)?.sheet;
        let Some(at) = self.shapes.iter().position(|shape| shape.id == id) else {
            return Err(host_error(1004, "the shape is no longer on its worksheet"));
        };
        let shape = self.shapes.remove(at);
        let insert_at = if to_front {
            self.shapes.iter().rposition(|held| held.sheet == sheet).map_or(self.shapes.len(), |last| last + 1)
        } else {
            self.shapes.iter().position(|held| held.sheet == sheet).unwrap_or(self.shapes.len())
        };
        self.shapes.insert(insert_at, shape);
        self.mirror_drawings(sheet);
        self.redraw(id)
    }

    fn charts_on(&self, sheet: usize) -> Vec<u64> {
        self.shapes
            .iter()
            .filter(|shape| shape.sheet == sheet && matches!(shape.kind, ShapeKind::Chart(_)))
            .map(|shape| shape.id)
            .collect()
    }

    /// `Shapes(n)` / `Shapes("name")` / `ChartObjects(...)`, among `listed`.
    fn pick_shape(&self, listed: &[u64], index: &Value) -> Result<u64, String> {
        match index {
            Value::String(name) => {
                let alias = localized_alias(name);
                let sheet = listed.first().and_then(|id| self.shape(*id).ok()).map(|shape| shape.sheet);
                self.shapes
                    .iter()
                    .filter(|shape| sheet.is_none_or(|sheet| shape.sheet == sheet))
                    .map(|shape| shape.id)
                    .find(|id| {
                        self.shape(*id).is_ok_and(|shape| {
                            shape.name.eq_ignore_ascii_case(name)
                                || alias.as_deref().is_some_and(|alias| shape.name.eq_ignore_ascii_case(alias))
                        })
                    })
                    .ok_or_else(|| {
                        // Measured: `Shapes("none")` says "The item with the
                        // specified name wasn't found."
                        let _ = name;
                        oxivba_core::host_error_described(NO_SUCH_SHAPE, "The item with the specified name wasn't found.")
                    })
            }
            value => match any_whole_number(value) {
                Some(number) if number >= 1 && (number as usize) <= listed.len() => {
                    Ok(listed[number as usize - 1])
                }
                _ => Err(host_error(NO_SUCH_SHAPE, "there is no such shape")),
            },
        }
    }

    /// GroupItems.Item is scoped to the group's children; worksheet-wide name
    /// lookup would incorrectly accept the group itself or an unrelated shape.
    fn pick_group_item(&self, children: &[u64], index: &Value) -> Result<u64, String> {
        match index {
            Value::String(name) => {
                let alias = localized_alias(name);
                children
                    .iter()
                    .copied()
                    .find(|id| self.shape(*id).is_ok_and(|shape| {
                        shape.name.eq_ignore_ascii_case(name)
                            || alias.as_deref().is_some_and(|alias| shape.name.eq_ignore_ascii_case(alias))
                    }))
                    .ok_or_else(|| oxivba_core::host_error_described(
                        NO_SUCH_SHAPE,
                        "The item with the specified name wasn't found.",
                    ))
            }
            value => match any_whole_number(value) {
                Some(number) if number >= 1 && (number as usize) <= children.len() => {
                    Ok(children[number as usize - 1])
                }
                _ => Err(host_error(NO_SUCH_SHAPE, "there is no such shape")),
            },
        }
    }

    fn new_shape(&mut self, sheet: usize, kind: ShapeKind, base: &str) -> u64 {
        let number = self.next_shape_number(sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let mut record = ShapeRecord::blank(id, sheet, format!("{base} {number}"), kind);
        record.theme_palette = workbook_office_theme_palette(&self.workbook.theme_colors);
        let accent1_theme = 5usize;
        if record.fill_theme == Some(accent1_theme) {
            record.fill = office_theme_colour_from(&record.theme_palette, accent1_theme);
        }
        if record.line_theme == Some(accent1_theme) {
            record.line = default_shape_outline(office_theme_colour_from(&record.theme_palette, accent1_theme));
        }
        self.shapes.push(record);
        id
    }

    fn placed(args: &[Value], from: usize, what: &str) -> Result<(f64, f64, f64, f64), String> {
        let number = |index: usize| -> Result<f64, String> {
            args.get(index)
                .and_then(any_number)
                .ok_or_else(|| format!("{what} takes Left, Top, Width and Height"))
        };
        Ok((number(from)?, number(from + 1)?, number(from + 2)?, number(from + 3)?))
    }

    fn add_word_document_class(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let add_failure = || host_error_described(1004, "OLEObjects クラスの Add メソッドが失敗しました。");
        if matches!(args.get(1), Some(Value::String(filename)) if !filename.is_empty()) {
            return Err(add_failure());
        }
        let _link = args
            .get(2)
            .map(|value| word_ole_display_as_icon(value).map_err(|()| add_failure()))
            .transpose()?;
        let display_as_icon = args
            .get(3)
            .map(|value| word_ole_display_as_icon(value).map_err(|()| add_failure()))
            .transpose()?
            .unwrap_or(false);
        let has_geometry_argument = args.get(7..).is_some_and(|geometry| !geometry.is_empty());
        let dimension_argument = |index: usize| -> Result<Option<f64>, String> {
            args.get(index)
                .map(|value| word_ole_add_dimension(value).map_err(|()| add_failure()))
                .transpose()
                .map(Option::flatten)
        };
        let requested_width = dimension_argument(9)?;
        let requested_left = dimension_argument(7)?.unwrap_or(0.0);
        let requested_top = dimension_argument(8)?.unwrap_or(0.0);
        let requested_add_width = if !display_as_icon {
            match requested_width {
                Some(width) => Some(word_ole_add_width(width, requested_left).ok_or_else(add_failure)?),
                None => None,
            }
        } else {
            None
        };
        let left = requested_left.max(0.0);
        let top = requested_top.max(0.0);
        let left = (left / 0.75).floor() * 0.75;
        let top = (top / 0.75).floor() * 0.75;
        let (width, height) = if display_as_icon {
            (72.0, 60.75)
        } else if let Some(width) = requested_add_width {
            (width, 24.75)
        } else if has_geometry_argument {
            (0.75, 24.75)
        } else {
            (441.75, 24.75)
        };
        let package = if requested_add_width.is_some() {
            word_document_with_ole_width(width)?
        } else {
            include_bytes!("data/word_document_empty.docx").to_vec()
        };
        let number = self.next_ole_object_number(sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let name = format!("Object {number}");
        let from = self.anchor_at(sheet, left, top);
        let to = self.anchor_at(sheet, left + width, top + height);
        let mut drawing = ShapeRecord::blank(id, sheet, name, ShapeKind::Other);
        drawing.left = left;
        drawing.top = top;
        drawing.width = width;
        drawing.height = height;
        drawing.visible = true;
        drawing.placement = 2;
        drawing.fill_visible = false;
        drawing.original = Some(oxicells_core::ir::Drawing {
            name: Some(drawing.name.clone()),
            ole_prog_id: Some("Word.Document.12".into()),
            ole_source_name: None,
            ole_clone_position: Some((left, top)),
            ole_embedded_package: Some(package),
            ole_package_is_empty: None,
            ole_clone_visible: None,
            ole_placement: Some(2),
            ole_type: Some(1),
            ole_display_as_icon: Some(display_as_icon),
            ole_auto_load: Some(false),
            ole_enabled: Some(true),
            ole_locked: Some(true),
            ole_print_object: Some(true),
            ole_shadow: Some(false),
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from,
            to: Some(to),
            extent: Some(((width / 0.75 * 9_525.0) as i64, (height / 0.75 * 9_525.0) as i64)),
            kind: DrawingKind::Other,
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        });
        self.shapes.push(drawing);
        self.redraw(id)?;
        Ok(self.part(DrawingPart::OleObject(id)))
    }

    fn add_empty_package_class(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        if matches!(args.get(1), Some(Value::String(filename)) if !filename.is_empty()) {
            return Err(host_error_described(1004, "OLEObjects クラスの Add メソッドが失敗しました。"));
        }
        let package = empty_package_embedding()?;
        let display_as_icon = matches!(args.get(3), Some(Value::Boolean(true)))
            || args.get(3).and_then(any_number).is_some_and(|value| value != 0.0);
        let left = args.get(7).and_then(any_number).unwrap_or(0.0).max(0.0);
        let top = args.get(8).and_then(any_number).unwrap_or(0.0).max(0.0);
        let left = (left / 0.75).floor() * 0.75;
        let top = (top / 0.75).floor() * 0.75;
        let number = self.next_ole_object_number(sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let name = format!("Object {number}");
        let width = 30.0;
        let height = 27.75;
        let from = self.anchor_at(sheet, left, top);
        let to = self.anchor_at(sheet, left + width, top + height);
        let mut drawing = ShapeRecord::blank(id, sheet, name, ShapeKind::Other);
        drawing.left = left;
        drawing.top = top;
        drawing.width = width;
        drawing.height = height;
        drawing.visible = true;
        drawing.placement = 2;
        drawing.fill_visible = false;
        drawing.original = Some(oxicells_core::ir::Drawing {
            name: Some(drawing.name.clone()),
            ole_prog_id: Some("Package".into()),
            ole_source_name: None,
            ole_clone_position: Some((left, top)),
            ole_embedded_package: Some(package),
            ole_package_is_empty: Some(true),
            ole_clone_visible: None,
            ole_placement: Some(2),
            ole_type: Some(1),
            ole_display_as_icon: Some(display_as_icon),
            ole_auto_load: Some(false),
            ole_enabled: Some(true),
            ole_locked: Some(true),
            ole_print_object: Some(true),
            ole_shadow: None,
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from,
            to: Some(to),
            extent: Some(((width / 0.75 * 9_525.0) as i64, (height / 0.75 * 9_525.0) as i64)),
            kind: DrawingKind::Other,
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        });
        self.shapes.push(drawing);
        self.redraw(id)?;
        Ok(self.part(DrawingPart::OleObject(id)))
    }

    pub(super) fn add_ole_object(&mut self, sheet: usize, args: &[Value], files: &mut oxivba_core::files::Files) -> Result<Value, String> {
        let cannot_insert = || host_error_described(1004, "オブジェクトを挿入できません。");
        let class_type = args.first().filter(|value| !matches!(value, Value::Missing | Value::Empty));
        if class_type.is_some_and(|value| text_of_value(value).eq_ignore_ascii_case("Package")) {
            return self.add_empty_package_class(sheet, args);
        }
        if class_type.is_some_and(|value| {
            matches!(text_of_value(value).to_ascii_lowercase().as_str(), "word.document" | "word.document.12")
        }) {
            return self.add_word_document_class(sheet, args);
        }
        if class_type.is_some_and(|value| !text_of_value(value).is_empty()) { return Err(cannot_insert()); }
        if class_type.is_some_and(|value| text_of_value(value).is_empty())
            && !matches!(args.get(1), Some(Value::String(_)))
        {
            return Err(cannot_insert());
        }
        let Some(Value::String(filename)) = args.get(1) else {
            return Err(host_error_described(1004, "OLEObjects クラスの Add メソッドが失敗しました。"));
        };
        if filename.is_empty() { return Err(cannot_insert()); }
        if matches!(args.get(2), Some(Value::Boolean(true)) | Some(Value::Integer(-1)) | Some(Value::Int16(-1))) {
            return Err(cannot_insert());
        }
        let display_as_icon = matches!(args.get(3), Some(Value::Boolean(true)))
            || args.get(3).and_then(any_number).is_some_and(|value| value != 0.0);
        let path = files.absolute_path(filename);
        let contents = files.bytes(&path).map_err(|_| cannot_insert())?.to_vec();
        let package = package_embedding(filename, &path, &contents)?;
        let left = args.get(7).and_then(any_number).unwrap_or(0.0).max(0.0);
        let top = args.get(8).and_then(any_number).unwrap_or(0.0).max(0.0);
        let left = (left / 0.75).floor() * 0.75;
        let top = (top / 0.75).floor() * 0.75;
        let number = self.next_ole_object_number(sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let name = format!("Object {number}");
        let from = self.anchor_at(sheet, left, top);
        let to = self.anchor_at(sheet, left + 51.0, top + 40.5);
        let mut drawing = ShapeRecord::blank(id, sheet, name, ShapeKind::Other);
        drawing.left = left;
        drawing.top = top;
        drawing.width = 51.0;
        drawing.height = 40.5;
        drawing.visible = true;
        drawing.placement = 2;
        drawing.fill_visible = false;
        drawing.original = Some(oxicells_core::ir::Drawing {
            name: Some(drawing.name.clone()),
            ole_prog_id: Some("Package".into()),
            ole_source_name: None,
            ole_clone_position: Some((left, top)),
            ole_embedded_package: Some(package),
            ole_package_is_empty: Some(false),
            ole_clone_visible: None,
            ole_placement: Some(2),
            ole_type: Some(1),
            ole_display_as_icon: Some(display_as_icon),
            ole_auto_load: Some(false),
            ole_enabled: Some(true),
            ole_locked: Some(true),
            ole_print_object: Some(true),
            ole_shadow: None,
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from,
            to: Some(to),
            extent: Some(((51.0 / 0.75 * 9_525.0) as i64, (40.5 / 0.75 * 9_525.0) as i64)),
            kind: DrawingKind::Other,
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        });
        self.shapes.push(drawing);
        self.redraw(id)?;
        Ok(self.part(DrawingPart::OleObject(id)))
    }

    /// `Shapes.AddShape Type, Left, Top, Width, Height`.
    fn add_shape(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let kind = args
            .first()
            .and_then(any_whole_number)
            .ok_or_else(|| "Shapes.AddShape takes a type by number".to_string())?;
        // Measured: 99 is taken (Curved Up Ribbon), -1 is 1004.
        if !(1..=183).contains(&kind) {
            return Err(host_error(1004, "that is not a shape type"));
        }
        let (left, top, width, height) = Self::placed(args, 1, "Shapes.AddShape")?;
        let base = auto_shape_name(kind);
        let id = self.new_shape(sheet, ShapeKind::Auto(kind), base);
        self.place_shape(id, left, top, width, height)?;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `Shapes.AddTextbox Orientation, Left, Top, Width, Height`.
    fn add_textbox(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let (left, top, width, height) = Self::placed(args, 1, "Shapes.AddTextbox")?;
        let id = self.new_shape(sheet, ShapeKind::TextBox, "TextBox");
        self.place_shape(id, left, top, width, height)?;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `Shapes.AddLine BeginX, BeginY, EndX, EndY`: the box is the two
    /// points' bounds, and a line drawn back towards the top-left is flipped
    /// both ways -- measured on (110, 300) to (10, 250).
    fn add_line(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let (x1, y1, x2, y2) = Self::placed(args, 0, "Shapes.AddLine")?;
        let id = self.new_shape(sheet, ShapeKind::Line, "Straight Connector");
        self.place_shape(id, x1.min(x2), y1.min(y2), (x2 - x1).abs(), (y2 - y1).abs())?;
        let shape = self.shape_mut(id)?;
        shape.flip_h = x2 < x1;
        shape.flip_v = y2 < y1;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `ChartObjects.Add Left, Top, Width, Height` and
    /// `Shapes.AddChart2 Style, XlChartType, Left, Top, Width, Height`: an
    /// empty clustered column chart, without a title, with a legend.
    fn add_chart(&mut self, sheet: usize, left: f64, top: f64, width: f64, height: f64, chart_type: i64) -> Result<u64, String> {
        let chart = ChartRecord {
            chart_type,
            series: Vec::new(),
            has_title: false,
            title: String::new(),
            has_legend: true,
            legend_position: -4152,
            axes: [AxisRecord::default(), AxisRecord::default()],
            style: 201,
            title_auto: true,
        };
        let id = self.new_shape(sheet, ShapeKind::Chart(Box::new(chart)), "Chart");
        self.place_shape(id, left, top, width, height)?;
        Ok(id)
    }

    /// Put a shape where it was asked, and draw it. Measured: a negative
    /// Left or Width becomes 0 without a word.
    fn place_shape(&mut self, id: u64, left: f64, top: f64, width: f64, height: f64) -> Result<(), String> {
        let sheet = {
            let shape = self.shape_mut(id)?;
            shape.left = left.max(0.0);
            shape.top = top.max(0.0);
            shape.width = width.max(0.0);
            shape.height = height.max(0.0);
            shape.sheet
        };
        self.mirror_drawings(sheet);
        Ok(())
    }

    fn redraw(&mut self, id: u64) -> Result<(), String> {
        let sheet = self.shape(id)?.sheet;
        self.mirror_drawings(sheet);
        Ok(())
    }

    fn replace_shape_characters(&mut self, id: u64, start: u32, length: Option<u32>, value: &Value) -> Result<(), String> {
        let shape = self.shape(id)?.clone();
        let added = value.string_units().map(|units| units.into_owned())
            .unwrap_or_else(|| text_of_value(value).encode_utf16().collect());
        let existing = shape.comment_at.and_then(|at| self.notes.iter().find(|note| note.at == at).map(Note::units))
            .unwrap_or_else(|| shape.text().encode_utf16().collect());
        let begin = (start.max(1) as usize - 1).min(existing.len());
        let end = length.filter(|n| *n != 0).map_or(existing.len(), |n| begin.saturating_add(n as usize).min(existing.len()));
        if let Some(at) = shape.comment_at {
            let mut changed = existing[..begin].to_vec();
            changed.extend_from_slice(&added);
            changed.extend_from_slice(&existing[end..]);
            let original = String::from_utf16(&existing);
            let prefix = String::from_utf16(&existing[..begin]);
            let selected = String::from_utf16(&existing[begin..end]);
            let inserted = String::from_utf16(&added);
            if let (Ok(original), Ok(prefix), Ok(selected), Ok(inserted)) = (original, prefix, selected, inserted) {
                if shape.text() == original {
                    self.shape_mut(id)?.replace_characters((prefix.chars().count() + 1) as u32, Some(selected.chars().count() as u32), &inserted);
                } else {
                    self.shape_mut(id)?.set_text(&String::from_utf16_lossy(&changed));
                }
            } else {
                self.shape_mut(id)?.set_text(&String::from_utf16_lossy(&changed));
            }
            self.redraw(id)?;
            if let Some(index) = self.note_at(at) { self.notes[index].set_units(changed); }
        } else {
            let text = shape.text();
            let mut offset = 0usize;
            let mut first = None;
            let mut last = 0usize;
            let mut chars = 0usize;
            for ch in text.chars() {
                let next = offset + ch.len_utf16();
                if offset < end && next > begin { first.get_or_insert(chars); last = chars + 1; }
                offset = next;
                chars += 1;
            }
            let first = first.unwrap_or(chars);
            let selected = last.saturating_sub(first);
            self.shape_mut(id)?.replace_characters((first + 1) as u32, Some(selected as u32), &String::from_utf16_lossy(&added));
            self.redraw(id)?;
        }
        Ok(())
    }

    fn delete_shape(&mut self, id: u64) -> Result<(), String> {
        let deleted = self.shape(id)?.clone();
        let sheet = deleted.sheet;
        if !deleted.group_children.is_empty() {
            let members: std::collections::HashSet<u64> = deleted.group_children.iter().copied().collect();
            self.shapes.retain(|shape| shape.id != id && !members.contains(&shape.id));
        } else {
            self.shapes.retain(|shape| shape.id != id);
            if let Some(parent) = deleted.group_parent_id {
                let remaining = if let Ok(group) = self.shape_mut(parent) {
                    group.group_children.retain(|child| *child != id);
                    group.group_children.clone()
                } else {
                    Vec::new()
                };
                if remaining.len() <= 1 {
                    let group_rotation = self.shape(parent).map_or(0.0, |group| group.rotation);
                    self.shapes.retain(|shape| shape.id != parent);
                    if let Some(child) = remaining.first().copied() {
                        if let Ok(shape) = self.shape_mut(child) {
                            shape.rotation += group_rotation;
                            shape.group_parent_id = None;
                            shape.group_name = None;
                        }
                    }
                } else if deleted.original.as_ref().is_some_and(|drawing| drawing.group_nested)
                    && self.shape(parent).is_ok_and(|group| group.group_nested)
                {
                    self.refit_group_after_child_delete(parent, &remaining)?;
                }
            }
        }
        self.mirror_drawings(sheet);
        Ok(())
    }

    fn refit_group_after_child_delete(&mut self, parent: u64, children: &[u64]) -> Result<(), String> {
        let group = self.shape(parent)?.clone();
        let pivot = group.group_bounds.unwrap_or((group.left, group.top, group.width, group.height));
        let pivot_x = pivot.0 + pivot.2 / 2.0;
        let pivot_y = pivot.1 + pivot.3 / 2.0;
        let angle = group.rotation.to_radians();
        let (sin, cos) = angle.sin_cos();
        let mut local = Vec::with_capacity(children.len());
        for child_id in children {
            let child = self.shape(*child_id)?.clone();
            let dx = child.left + child.width / 2.0 - pivot_x;
            let dy = child.top + child.height / 2.0 - pivot_y;
            let center_x = pivot_x + dx * cos + dy * sin;
            let center_y = pivot_y - dx * sin + dy * cos;
            local.push((
                *child_id,
                center_x - child.width / 2.0,
                center_y - child.height / 2.0,
                child.width,
                child.height,
            ));
        }
        let min_x = local.iter().map(|item| item.1).fold(f64::INFINITY, f64::min);
        let min_y = local.iter().map(|item| item.2).fold(f64::INFINITY, f64::min);
        let max_x = local.iter().map(|item| item.1 + item.3).fold(f64::NEG_INFINITY, f64::max);
        let max_y = local.iter().map(|item| item.2 + item.4).fold(f64::NEG_INFINITY, f64::max);
        let width = max_x - min_x;
        let height = max_y - min_y;
        if width <= 0.0 || height <= 0.0 { return Ok(()); }
        let local_center_x = (min_x + max_x) / 2.0;
        let local_center_y = (min_y + max_y) / 2.0;
        let center_x = pivot_x + (local_center_x - pivot_x) * cos - (local_center_y - pivot_y) * sin;
        let center_y = pivot_y + (local_center_x - pivot_x) * sin + (local_center_y - pivot_y) * cos;
        let left = center_x - width / 2.0;
        let top = center_y - height / 2.0;
        let from = self.anchor_at(group.sheet, left, top);
        let to = self.anchor_at(group.sheet, left + width, top + height);
        let width_emu = (width * EMU_PER_POINT).round() as i64;
        let height_emu = (height * EMU_PER_POINT).round() as i64;
        let group_transform = GroupTransform {
            offset: ((left * EMU_PER_POINT).round() as i64, (top * EMU_PER_POINT).round() as i64),
            extent: (width_emu, height_emu),
            child_offset: (0, 0),
            child_extent: (width_emu, height_emu),
        };
        for (child_id, child_left, child_top, child_width, child_height) in local {
            if let Ok(child) = self.shape_mut(child_id) {
                if let Some(drawing) = child.original.as_mut() {
                    drawing.from = from;
                    drawing.to = Some(to);
                    drawing.frame = Some(Frame {
                        x: (child_left - min_x) / width,
                        y: (child_top - min_y) / height,
                        w: child_width / width,
                        h: child_height / height,
                    });
                    drawing.group_transform = Some(group_transform);
                    drawing.group_nested = false;
                }
            }
        }
        if let Ok(group) = self.shape_mut(parent) {
            group.left = left;
            group.top = top;
            group.width = width;
            group.height = height;
            group.group_bounds = Some((left, top, width, height));
            group.group_transform = Some(group_transform);
            group.group_nested = false;
        }
        Ok(())
    }

    /// A copy of a shape, twelve points down and across, numbered on:
    /// measured, `Duplicate` of Rectangle 6 at (10, 300) is Rectangle 8 at
    /// (22, 312).
    fn duplicate_shape(&mut self, id: u64, at: Option<(f64, f64)>) -> Result<u64, String> {
        let source = self.shape(id)?.clone();
        let is_ole = source.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some());
        let base = if is_ole {
            "Object".to_string()
        } else {
            match &source.kind {
                ShapeKind::Auto(kind) => auto_shape_name(*kind).to_string(),
                ShapeKind::TextBox => "TextBox".to_string(),
                ShapeKind::Line => "Straight Connector".to_string(),
                ShapeKind::Chart(_) => "Chart".to_string(),
                ShapeKind::Picture => "Picture".to_string(),
                ShapeKind::Other => "Object".to_string(),
            }
        };
        let number = if is_ole {
            self.next_ole_object_number(source.sheet)
        } else {
            self.next_shape_number(source.sheet)
        };
        let new_id = self.next_shape_id;
        self.next_shape_id += 1;
        let (left, top) = at.unwrap_or((source.left + 12.0, source.top + 12.0));
        let ole_clone_of = is_ole.then(|| source.ole_clone_of.clone().unwrap_or_else(|| source.name.clone()));
        let copy = ShapeRecord { id: new_id, name: format!("{base} {number}"), left, top, ole_clone_of, ..source };
        self.shapes.push(copy);
        self.redraw(new_id)?;
        Ok(new_id)
    }

    /// `Worksheet.Paste` with a shape set aside: a copy lands with its
    /// top-left on the selection. Measured: pasted with D1 selected, the
    /// copy sits at Left 162, Top 0.
    pub(super) fn paste_shape(&mut self, sheet: usize) -> Result<Option<Value>, String> {
        let Some(id) = self.shape_clipboard else {
            return Ok(None);
        };
        if self.shape(id).is_err() {
            return Ok(None);
        }
        let at = self.selection.first();
        let left = (0..at.column).map(|c| f64::from(self.column_px(sheet, c))).sum::<f64>() * 0.75;
        let top = (1..at.row).map(|r| f64::from(self.row_px(sheet, r))).sum::<f64>() * 0.75;
        let source_sheet = self.shape(id)?.sheet;
        let new_id = self.duplicate_shape(id, Some((left, top)))?;
        if source_sheet != sheet {
            let old = source_sheet;
            self.shape_mut(new_id)?.sheet = sheet;
            self.mirror_drawings(old);
            self.mirror_drawings(sheet);
        }
        Ok(Some(Value::Boolean(true)))
    }

    /// `Chart.SetSourceData Source, PlotBy`: the block is read the way
    /// Excel reads it. By columns (the default), a first row of text is the
    /// series' names and a first column of text the categories; by rows the
    /// other way about. Measured: A1:B3 (k v / a 1 / b 2) gives one series
    /// `v` with categories a, b and formula
    /// `=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)`; by rows it
    /// gives `a` and `b` with `=SERIES(Sheet1!$A$2,Sheet1!$B$1,Sheet1!$B$2,1)`;
    /// B2:B3 alone gives `系列1` with `=SERIES(,,Sheet1!$B$2:$B$3,1)`.
    fn set_source_data(&mut self, id: u64, range: CellRange, by_rows: bool) -> Result<(), String> {
        self.settle(range);
        let sheet = range.sheet;
        let is_text = |host: &Self, row: u32, column: u32| {
            host.cell_here(sheet, row, column)
                .is_some_and(|cell| !matches!(cell.value, CellValue::Number(_) | CellValue::Empty))
        };
        let number_at = |host: &Self, row: u32, column: u32| -> f64 {
            match host.cell_here(sheet, row, column).map(|cell| cell.value) {
                Some(CellValue::Number(number)) => number,
                _ => 0.0,
            }
        };
        let text_at = |host: &Self, row: u32, column: u32| -> String {
            host.cell_here(sheet, row, column)
                .map(|cell| shown_text(&from_cell_value(&cell.value), cell.style.number_format.as_deref()))
                .unwrap_or_default()
        };
        let cell_ref = |host: &Self, row: u32, column: u32| -> String {
            host.address_of(CellRange::single(CellAddress { sheet, row, column }))
        };
        let block_ref = |host: &Self, r1: u32, c1: u32, r2: u32, c2: u32| -> String {
            host.address_of(CellRange { sheet, start_row: r1, start_column: c1, end_row: r2, end_column: c2 })
        };
        let (rows, columns) = (
            (range.start_row..=range.end_row).collect::<Vec<u32>>(),
            (range.start_column..=range.end_column).collect::<Vec<u32>>(),
        );
        let mut series = Vec::new();
        if !by_rows {
            let header = rows.len() > 1 && columns.iter().skip(1).any(|c| is_text(self, rows[0], *c))
                || (rows.len() > 1 && columns.len() == 1 && is_text(self, rows[0], columns[0]));
            let data_rows: Vec<u32> = if header { rows[1..].to_vec() } else { rows.clone() };
            let categories = columns.len() > 1 && data_rows.iter().any(|r| is_text(self, *r, columns[0]));
            let value_columns: Vec<u32> = if categories { columns[1..].to_vec() } else { columns.clone() };
            let (first, last) = (data_rows[0], data_rows[data_rows.len() - 1]);
            for column in value_columns {
                let name_ref = header.then(|| cell_ref(self, rows[0], column));
                series.push(SeriesRecord {
                    name: match &name_ref {
                        Some(_) => text_at(self, rows[0], column),
                        None => format!("系列{}", series.len() + 1),
                    },
                    name_ref,
                    values_ref: Some(block_ref(self, first, column, last, column)),
                    x_ref: categories.then(|| block_ref(self, first, columns[0], last, columns[0])),
                    values: data_rows.iter().map(|r| number_at(self, *r, column)).collect(),
                    xs: data_rows
                        .iter()
                        .enumerate()
                        .map(|(at, r)| if categories { text_at(self, *r, columns[0]) } else { (at + 1).to_string() })
                        .collect(),
                    color: None,
                    chart_type: None,
                    has_labels: false,
                });
            }
        } else {
            let header = columns.len() > 1 && rows.iter().skip(1).any(|r| is_text(self, *r, columns[0]));
            let data_columns: Vec<u32> = if header { columns[1..].to_vec() } else { columns.clone() };
            let categories = rows.len() > 1 && data_columns.iter().any(|c| is_text(self, rows[0], *c));
            let value_rows: Vec<u32> = if categories { rows[1..].to_vec() } else { rows.clone() };
            let (first, last) = (data_columns[0], data_columns[data_columns.len() - 1]);
            for row in value_rows {
                let name_ref = header.then(|| cell_ref(self, row, columns[0]));
                series.push(SeriesRecord {
                    name: match &name_ref {
                        Some(_) => text_at(self, row, columns[0]),
                        None => format!("系列{}", series.len() + 1),
                    },
                    name_ref,
                    values_ref: Some(block_ref(self, row, first, row, last)),
                    x_ref: categories.then(|| block_ref(self, rows[0], first, rows[0], last)),
                    values: data_columns.iter().map(|c| number_at(self, row, *c)).collect(),
                    xs: data_columns
                        .iter()
                        .enumerate()
                        .map(|(at, c)| if categories { text_at(self, rows[0], *c) } else { (at + 1).to_string() })
                        .collect(),
                    color: None,
                    chart_type: None,
                    has_labels: false,
                });
            }
        }
        let chart = self.chart_mut(id)?;
        chart.series = series;
        self.redraw(id)
    }

    /// `=SERIES(name, categories, values, order)`, the way Excel writes a
    /// series' formula: a name set as text is quoted, a part not given is
    /// left empty.
    fn series_formula(&self, id: u64, number: usize) -> Result<String, String> {
        let series = self.series(id, number)?;
        let name = match (&series.name_ref, series.name.is_empty()) {
            (Some(reference), _) => reference.clone(),
            (None, false) if !series.name.starts_with("系列") => format!("\"{}\"", series.name),
            _ => String::new(),
        };
        Ok(format!(
            "=SERIES({name},{},{},{number})",
            series.x_ref.clone().unwrap_or_default(),
            series.values_ref.clone().unwrap_or_default()
        ))
    }

    fn cells_of_range_object(&self, value: &Value) -> Option<CellRange> {
        match value {
            Value::Object(object) => self.range(object),
            _ => None,
        }
    }

    fn range_numbers(&mut self, range: CellRange) -> Vec<f64> {
        self.settle(range);
        range
            .addresses()
            .map(|at| match self.cell_here(at.sheet, at.row, at.column).map(|cell| cell.value) {
                Some(CellValue::Number(number)) => number,
                _ => 0.0,
            })
            .collect()
    }

    fn range_texts(&mut self, range: CellRange) -> Vec<String> {
        self.settle(range);
        range
            .addresses()
            .map(|at| {
                self.cell_here(at.sheet, at.row, at.column)
                    .map(|cell| shown_text(&from_cell_value(&cell.value), cell.style.number_format.as_deref()))
                    .unwrap_or_default()
            })
            .collect()
    }

    fn variant_row(values: Vec<Value>) -> Value {
        Value::Array(ArrayValue {
            dimensions: vec![ArrayDimension { lower_bound: 1, length: values.len() }],
            values,
            element_default: Box::new(Value::Empty),
            resizable: true,
        })
    }

    // ---- reading ------------------------------------------------------------

    pub(super) fn drawing_get(&mut self, part: DrawingPart, name: &str) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Shapes(sheet) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shapes_on(sheet).len() as i64)),
                "parent" => Some(self.object(HostObject::Worksheet(sheet))),
                _ => None,
            }),
            DrawingPart::OleObjects(sheet) | DrawingPart::OleObjectsSnapshot(sheet, _) => {
                let ids = self.ole_objects_for_part(part);
                Ok(match lower.as_str() {
                "count" => Some(Value::Integer(ids.len() as i64)),
                "parent" => Some(self.object(HostObject::Worksheet(sheet))),
                "application" => Some(self.object(HostObject::Application)),
                "creator" => Some(Value::Integer(1_480_803_660)),
                "visible" => {
                    if ids.is_empty() {
                        Some(Value::Empty)
                    } else {
                        Some(Value::Boolean(ids.iter().all(|id| self.shape(*id).is_ok_and(|shape| shape.visible))))
                    }
                }
                "shaperange" => Some(self.shape_range_object(ids)),
                "left" | "top" | "width" | "height" | "locked" | "enabled" | "placement"
                | "printobject" | "onaction" | "zorder" | "shadow"
                | "autoload" => {
                    return Err(host_error_described(
                        1004,
                        format!("OLEObjects \u{30af}\u{30e9}\u{30b9}\u{306e} {name} \u{30d7}\u{30ed}\u{30d1}\u{30c6}\u{30a3}\u{3092}\u{53d6}\u{5f97}\u{3067}\u{304d}\u{307e}\u{305b}\u{3093}\u{3002}"),
                    ));
                }
                "border" | "interior" => {
                    return Err(host_error(
                        438,
                        "\u{30aa}\u{30d6}\u{30b8}\u{30a7}\u{30af}\u{30c8}\u{306f}\u{3001}\u{3053}\u{306e}\u{30d7}\u{30ed}\u{30d1}\u{30c6}\u{30a3}\u{307e}\u{305f}\u{306f}\u{30e1}\u{30bd}\u{30c3}\u{30c9}\u{3092}\u{30b5}\u{30dd}\u{30fc}\u{30c8}\u{3057}\u{3066}\u{3044}\u{307e}\u{305b}\u{3093}\u{3002}",
                    ));
                }
                "sourcename" => {
                    return Err(host_error_described(
                        1004,
                        "\u{30a2}\u{30d7}\u{30ea}\u{30b1}\u{30fc}\u{30b7}\u{30e7}\u{30f3}\u{5b9a}\u{7fa9}\u{307e}\u{305f}\u{306f}\u{30aa}\u{30d6}\u{30b8}\u{30a7}\u{30af}\u{30c8}\u{5b9a}\u{7fa9}\u{306e}\u{30a8}\u{30e9}\u{30fc}\u{3067}\u{3059}\u{3002}",
                    ));
                }
                _ => None,
                })
            }
            DrawingPart::GroupObject(id) => self.shape_get(part, id, &lower),
            DrawingPart::GroupItems(id) => {
                let group = self.shape(id)?;
                if group.group_children.is_empty() {
                    return Err(host_error(1004, "the shape is in no group"));
                }
                Ok(match lower.as_str() {
                    "count" => Some(Value::Integer(group.group_children.len() as i64)),
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    _ => None,
                })
            }
            DrawingPart::OleFormat(id) => {
                let shape = self.shape(id)?;
                let prog_id = shape.original.as_ref().and_then(|drawing| drawing.ole_prog_id.clone())
                    .ok_or_else(|| host_error(438, "Object doesn't support this property or method"))?;
                Ok(match lower.as_str() {
                    "progid" => Some(Value::String(prog_id)),
                    "creator" => Some(Value::Integer(1_480_803_660)),
                    "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
                    "application" => Some(self.object(HostObject::Application)),
                    "autoload" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_auto_load).unwrap_or(false))),
                    "enabled" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_enabled).unwrap_or(true))),
                    "oletype" => Some(Value::Integer(shape.original.as_ref().and_then(|drawing| drawing.ole_type).unwrap_or(1))),
                    "locked" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_locked).unwrap_or(true))),
                    "printobject" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_print_object).unwrap_or(true))),
                    "shadow" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_shadow).unwrap_or(false))),
                    "object" => Some(self.part(DrawingPart::OleObject(id))),
                    _ => None,
                })
            }
            DrawingPart::OleObject(id) => {
                let shape = self.shape(id)?.clone();
                let prog_id = shape.original.as_ref().and_then(|drawing| drawing.ole_prog_id.clone())
                    .ok_or_else(|| host_error(438, "Object doesn't support this property or method"))?;
                Ok(match lower.as_str() {
                    "name" => Some(Value::String(shape.name.clone())),
                    "index" => Some(Value::Integer(self.shapes_on(shape.sheet).iter().position(|candidate| *candidate == id).map_or(0, |at| at as i64 + 1))),
                    "progid" => Some(Value::String(prog_id)),
                    "creator" => Some(Value::Integer(1_480_803_660)),
                    "application" => Some(self.object(HostObject::Application)),
                    "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
                    "left" => Some(Value::Double(shape.left)),
                    "top" => Some(Value::Double(shape.top)),
                    "width" => Some(Value::Double(shape.width)),
                    "height" => Some(Value::Double(shape.height)),
                    "placement" => Some(Value::Integer(shape.placement)),
                    "visible" => Some(Value::Boolean(shape.visible)),
                    "autoload" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_auto_load).unwrap_or(false))),
                    "enabled" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_enabled).unwrap_or(true))),
                    "oletype" => Some(Value::Integer(shape.original.as_ref().and_then(|drawing| drawing.ole_type).unwrap_or(1))),
                    "locked" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_locked).unwrap_or(true))),
                    "printobject" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_print_object).unwrap_or(true))),
                    "shadow" => Some(Value::Boolean(shape.original.as_ref().and_then(|drawing| drawing.ole_shadow).unwrap_or(false))),
                    "onaction" => Some(Value::String(shape.on_action.clone())),
                    "object" => Some(if prog_id.eq_ignore_ascii_case("Word.Document") || prog_id.eq_ignore_ascii_case("Word.Document.12") {
                        self.object(HostObject::WordDocument(id))
                    } else {
                        Value::Nothing
                    }),
                    "althtml" | "autoupdate" | "linkedcell" | "listfillrange" | "sourcename" => Some(Value::Nothing),
                    "shaperange" => Some(self.shape_range_object(vec![id])),
                    "border" => Some(self.part(DrawingPart::OleBorder(id))),
                    "interior" => Some(self.part(DrawingPart::OleInterior(id))),
                    "topleftcell" => {
                        let cell = self.cell_under(shape.sheet, shape.left, shape.top);
                        Some(self.object(HostObject::Range(CellRange::single(cell))))
                    }
                    "bottomrightcell" => {
                        let cell = self.cell_under(shape.sheet, shape.left + shape.width, shape.top + shape.height);
                        Some(self.object(HostObject::Range(CellRange::single(cell))))
                    }
                    _ => None,
                })
            }
            DrawingPart::OleBorder(id) => Ok(match lower.as_str() {
                "application" => Some(self.object(HostObject::Application)),
                "creator" => Some(Value::Integer(1_480_803_660)),
                "parent" => Some(self.part(DrawingPart::OleObject(id))),
                "color" => {
                    let color = self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_border_color.as_deref()).and_then(|color| colour_to_packed(Some(color)));
                    let value = color.map_or(0, |color| {
                        let index = self.nearest_palette_index(color);
                        if (1..=56).contains(&index) { self.indexed_colour_palette()[index as usize - 1] } else { color }
                    });
                    Some(Value::Integer(value))
                }
                "colorindex" => {
                    let color = self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_border_color.as_deref()).and_then(|color| colour_to_packed(Some(color)));
                    Some(Value::Integer(color.map(|color| self.nearest_palette_index(color)).unwrap_or(-4105)))
                }
                "linestyle" => Some(Value::Integer(if self.shape(id)?.original.as_ref().is_some_and(|drawing| drawing.ole_border_color.is_some()) { 1 } else { -4105 })),
                "weight" => Some(Value::Integer(self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_border_weight).unwrap_or(2))),
                "themecolor" => Some(Value::Nothing),
                "tintandshade" => Some(Value::Double(0.0)),
                _ => None,
            }),
            DrawingPart::OleInterior(id) => Ok(match lower.as_str() {
                "application" => Some(self.object(HostObject::Application)),
                "creator" => Some(Value::Integer(1_480_803_660)),
                "parent" => Some(self.part(DrawingPart::OleObject(id))),
                "color" => Some(Value::Integer(self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_interior_color.as_deref()).and_then(|color| colour_to_packed(Some(color))).unwrap_or(16_777_215))),
                "colorindex" => {
                    let color = self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_interior_color.as_deref()).and_then(|color| colour_to_packed(Some(color)));
                    Some(Value::Integer(color.map(|color| self.nearest_palette_index(color)).unwrap_or(-4105)))
                }
                "patterncolorindex" => Some(Value::Integer(-4105)),
                "pattern" => Some(Value::Integer(1)),
                "patterncolor" => Some(Value::Integer(0)),
                "themecolor" | "patternthemecolor" | "patterntintandshade" | "invertifnegative" => Some(Value::Nothing),
                "tintandshade" => Some(Value::Double(0.0)),
                _ => None,
            }),
            DrawingPart::PictureFormat(id) => {
                let shape = self.shape(id)?.clone();
                let format = match shape.original.as_ref().map(|drawing| &drawing.kind) {
                    Some(DrawingKind::Picture { format, .. }) => Some(format),
                    _ => None,
                };
                let image_size = match shape.original.as_ref().map(|drawing| &drawing.kind) {
                    Some(DrawingKind::Picture { bytes, format }) => picture_source_size_points(bytes, format),
                    _ => None,
                }.unwrap_or((shape.width, shape.height));
                let value = match (format, lower.as_str()) {
                    (Some(format), "brightness") => Some(Value::Single((0.5 + f64::from(format.brightness) / 200_000.0) as f32)),
                    (Some(format), "contrast") => Some(Value::Single((0.5 + f64::from(format.contrast) / 200_000.0) as f32)),
                    (Some(format), "colortype") => Some(Value::Integer(format.color_type)),
                    (Some(format), "cropleft") => Some(Value::Single((f64::from(format.crop.0) * image_size.0 / 100_000.0) as f32)),
                    (Some(format), "croptop") => Some(Value::Single((f64::from(format.crop.1) * image_size.1 / 100_000.0) as f32)),
                    (Some(format), "cropright") => Some(Value::Single((f64::from(format.crop.2) * image_size.0 / 100_000.0) as f32)),
                    (Some(format), "cropbottom") => Some(Value::Single((f64::from(format.crop.3) * image_size.1 / 100_000.0) as f32)),
                    (Some(format), "transparencycolor") => Some(Value::Integer(format.transparency_color.as_deref().and_then(|color| colour_to_packed(Some(color))).unwrap_or(i32::MIN as i64))),
                    (Some(format), "transparentbackground") => Some(mso(format.transparent_background)),
                    (_, "creator") => Some(Value::Integer(1_480_803_660)),
                    (Some(_), "crop") => return Ok(Some(self.part(DrawingPart::Crop(id)))),
                    (None, "crop") => return Err(host_error(438, "Object doesn't support this property or method")),
                    (None, "brightness" | "contrast" | "colortype" | "cropleft" | "croptop" | "cropright" | "cropbottom" | "transparencycolor" | "transparentbackground") => Some(Value::Empty),
                    (_, "parent") => Some(self.part(DrawingPart::Shape(id))),
                    (_, "application") => Some(Value::String("Microsoft Excel".into())),
                    _ => None,
                };
                Ok(value)
            }
            DrawingPart::Crop(id) => {
                let shape = self.shape(id)?.clone();
                let (picture_w, picture_h, crop_left, crop_top, crop_right, crop_bottom) = crop_picture_geometry(&shape)?;
                let offset_x = (crop_right - crop_left) / 2.0;
                let offset_y = (crop_bottom - crop_top) / 2.0;
                let value = match lower.as_str() {
                    "picturewidth" => Some(Value::Single(picture_w as f32)),
                    "pictureheight" => Some(Value::Single(picture_h as f32)),
                    "pictureoffsetx" => Some(Value::Single(offset_x as f32)),
                    "pictureoffsety" => Some(Value::Single(offset_y as f32)),
                    "shapewidth" => Some(Value::Single(shape.width as f32)),
                    "shapeheight" => Some(Value::Single(shape.height as f32)),
                    "shapeleft" => Some(Value::Single(shape.left as f32)),
                    "shapetop" => Some(Value::Single(shape.top as f32)),
                    "creator" => Some(Value::Integer(1_480_803_660)),
                    "application" => Some(Value::String("Microsoft Excel".into())),
                    _ => None,
                };
                Ok(value)
            }
            DrawingPart::ChartObjects(sheet) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.charts_on(sheet).len() as i64)),
                "parent" => Some(self.object(HostObject::Worksheet(sheet))),
                _ => None,
            }),
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) => self.shape_get(part, id, &lower),
            DrawingPart::Fill(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "forecolor" => Some(self.part(DrawingPart::FillColor(id))),
                    "visible" => Some(mso(shape.fill_visible)),
                    "transparency" => Some(Value::Single(shape.transparency as f32)),
                    // A solid fill, measured: Type msoFillSolid, the gradient,
                    // pattern and texture kinds mixed (-2), no gradient
                    // variant, not turning with the shape; the angle is 440
                    // and the texture's own settings 5 (its name 70).
                    "type" => Some(Value::Integer(1)),
                    "gradientcolortype" | "gradientstyle" | "pattern" | "presetgradienttype" | "presettexture" | "texturetype" => {
                        Some(Value::Integer(-2))
                    }
                    "gradientvariant" | "rotatewithobject" => Some(Value::Integer(0)),
                    "gradientangle" => return Err(host_error(440, "a solid fill has no gradient angle")),
                    "texturename" => return Err(host_error(70, "a solid fill has no texture")),
                    "texturealignment" | "texturehorizontalscale" | "textureoffsetx" | "textureoffsety" | "texturetile" | "textureverticalscale" => {
                        return Err(host_error(5, "a solid fill has no texture"))
                    }
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    _ => None,
                })
            }
            DrawingPart::FillColor(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    // RGB is the default member: measured, a ForeColor read as
                    // a value is its colour.
                    "rgb" | "value" => Some(Value::Integer(shape.fill)),
                    "objectthemecolor" => Some(Value::Integer(shape.fill_theme.map_or(0, |theme| theme as i64))),
                    // Measured on a fresh shape (accent1): Type
                    // msoColorTypeScheme (2), SchemeColor 21, no brightness
                    // or tint, the FillFormat its Parent. An RGB colour is
                    // Type 1; SchemeColor is given only for accent1.
                    "type" => Some(Value::Integer(if shape.fill_theme.is_some() || shape.fill_scheme_color.is_some() { 2 } else { 1 })),
                    "schemecolor" => Some(Value::Integer(scheme_color_for_style_in_palette(
                        shape.fill_theme,
                        shape.fill_scheme_color,
                        shape.fill,
                        shape.fill_brightness,
                        &self.indexed_colour_palette(),
                    ))),
                    "brightness" | "tintandshade" => Some(Value::Single(shape.fill_brightness as f32)),
                    "parent" => Some(self.part(DrawingPart::Fill(id))),
                    _ => None,
                })
            }
            DrawingPart::Line(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "forecolor" => Some(self.part(DrawingPart::LineColor(id))),
                    "backcolor" => Some(self.part(DrawingPart::LineBackColor(id))),
                    "visible" => Some(mso(shape.line_visible)),
                    "weight" => Some(Value::Single(shape.line_weight as f32)),
                    "dashstyle" => Some(Value::Integer(shape.dash)),
                    "endarrowheadstyle" => Some(Value::Integer(shape.arrow_end)),
                    "beginarrowheadstyle" => Some(Value::Integer(shape.arrow_start)),
                    // DrawingML omits medium sizes; None in the record means
                    // the Excel default of msoArrowheadLength/WidthMedium.
                    "beginarrowheadlength" => Some(Value::Integer(shape.arrow_start_length.unwrap_or(2))),
                    "beginarrowheadwidth" => Some(Value::Integer(shape.arrow_start_width.unwrap_or(2))),
                    "endarrowheadlength" => Some(Value::Integer(shape.arrow_end_length.unwrap_or(2))),
                    "endarrowheadwidth" => Some(Value::Integer(shape.arrow_end_width.unwrap_or(2))),
                    "style" => Some(Value::Integer(1)),
                    "insetpen" => Some(Value::Integer(0)),
                    "pattern" => Some(Value::Integer(shape.line_pattern)),
                    "transparency" => Some(Value::Single(shape.line_transparency as f32)),
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    _ => None,
                })
            }
            DrawingPart::LineColor(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "rgb" | "value" => Some(Value::Integer(shape.line)),
                    "objectthemecolor" => Some(Value::Integer(shape.line_theme.map_or(0, |theme| theme as i64))),
                    "type" => Some(Value::Integer(if shape.line_theme.is_some() || shape.line_scheme_color.is_some() { 2 } else { 1 })),
                    "schemecolor" => Some(Value::Integer(scheme_color_for_style_in_palette(
                        shape.line_theme,
                        shape.line_scheme_color,
                        shape.line,
                        shape.line_brightness,
                        &self.indexed_colour_palette(),
                    ))),
                    "brightness" | "tintandshade" => Some(Value::Single(shape.line_brightness as f32)),
                    "parent" => Some(self.part(DrawingPart::Line(id))),
                    _ => None,
                })
            }
            DrawingPart::LineBackColor(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "rgb" | "value" => Some(Value::Integer(shape.line_back)),
                    "objectthemecolor" => Some(Value::Integer(shape.line_back_theme.map_or(0, |theme| theme as i64))),
                    "type" => Some(Value::Integer(if shape.line_back_theme.is_some() || shape.line_back_scheme_color.is_some() { 2 } else { 1 })),
                    "schemecolor" => Some(Value::Integer(scheme_color_for_style_in_palette(
                        shape.line_back_theme,
                        shape.line_back_scheme_color,
                        shape.line_back,
                        shape.line_back_brightness,
                        &self.indexed_colour_palette(),
                    ))),
                    "brightness" | "tintandshade" => Some(Value::Single(shape.line_back_brightness as f32)),
                    "parent" => Some(self.part(DrawingPart::Line(id))),
                    _ => None,
                })
            }
            DrawingPart::TextFrame(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "characters" => Some(self.part(DrawingPart::Characters(id, 1, None))),
                    "horizontalalignment" => Some(Value::Integer(shape.h_align)),
                    "verticalalignment" => Some(Value::Integer(shape.v_align)),
                    // Measured on a rectangle: AutoSize and AutoMargins are
                    // 1004 (a text box's are not measured and keep theirs).
                    "autosize" | "automargins" if matches!(shape.kind, ShapeKind::Auto(_)) => {
                        return Err(host_error(1004, "Unable to get the AutoSize property"))
                    }
                    "autosize" => Some(Value::Boolean(shape.auto_size)),
                    "marginleft" => Some(Value::Single(shape.margins.0 as f32)),
                    "margintop" => Some(Value::Single(shape.margins.1 as f32)),
                    "marginright" => Some(Value::Single(shape.margins.2 as f32)),
                    "marginbottom" => Some(Value::Single(shape.margins.3 as f32)),
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    // Measured: horizontal text overflowing both ways,
                    // reading order mixed.
                    "horizontaloverflow" | "verticaloverflow" | "orientation" => Some(Value::Integer(1)),
                    "readingorder" => Some(Value::Integer(-2)),
                    _ => None,
                })
            }
            DrawingPart::TextFrame2(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "textrange" => Some(self.part(DrawingPart::TextRange(id))),
                    "hastext" => Some(mso(!shape.text().is_empty())),
                    // msoAnchorTop 1, Middle 3, Bottom 4.
                    "verticalanchor" => Some(Value::Integer(match shape.v_align {
                        -4108 => 3,
                        -4107 => 4,
                        _ => 1,
                    })),
                    "wordwrap" => Some(mso(true)),
                    "autosize" => Some(Value::Integer(i64::from(shape.auto_size))),
                    "marginleft" => Some(Value::Single(shape.margins.0 as f32)),
                    "margintop" => Some(Value::Single(shape.margins.1 as f32)),
                    "marginright" => Some(Value::Single(shape.margins.2 as f32)),
                    "marginbottom" => Some(Value::Single(shape.margins.3 as f32)),
                    // Measured: anchored left, horizontal, turning with the
                    // shape, no path or warp, WordArt mixed.
                    "horizontalanchor" | "orientation" => Some(Value::Integer(1)),
                    "notextrotation" | "pathformat" | "warpformat" => Some(Value::Integer(0)),
                    "wordartformat" => Some(Value::Integer(-2)),
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    _ => None,
                })
            }
            DrawingPart::TextRange(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    // Text is the default member: measured, a TextRange read
                    // as a value is its text.
                    "text" | "value" => Some(Value::String(shape.text())),
                    "font" => Some(self.part(DrawingPart::Font2(id))),
                    "paragraphformat" => Some(self.part(DrawingPart::ParagraphFormat(id))),
                    "characters" => Some(self.part(DrawingPart::Characters(id, 1, None))),
                    "paragraphs" => Some(self.part(DrawingPart::Paragraphs(id))),
                    "length" => Some(Value::Integer(shape.text().chars().count() as i64)),
                    // Measured on "ab cd": Count 1 (one range), Start 1,
                    // LanguageID 1033, Parent the TextFrame2.
                    "count" | "start" => Some(Value::Integer(1)),
                    "languageid" => Some(Value::Integer(1033)),
                    "parent" => Some(self.part(DrawingPart::TextFrame2(id))),
                    _ => None,
                })
            }
            DrawingPart::ParagraphFormat(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    // The rest of a fresh paragraph, measured: baseline auto
                    // (5), far-east line breaking and hanging punctuation on,
                    // level 1, no indents or spacing, single lines, mixed
                    // direction, wrapping, the TextRange its Parent.
                    "baselinealignment" => Some(Value::Integer(5)),
                    "fareastlinebreaklevel" | "hangingpunctuation" | "linerulewithin" | "wordwrap" => Some(Value::Integer(-1)),
                    "indentlevel" => Some(Value::Integer(1)),
                    "firstlineindent" | "leftindent" | "rightindent" | "spaceafter" | "spacebefore" => Some(Value::Single(0.0)),
                    "spacewithin" => Some(Value::Single(1.0)),
                    "lineruleafter" | "linerulebefore" => Some(Value::Integer(0)),
                    "textdirection" => Some(Value::Integer(-2)),
                    "parent" => Some(self.part(DrawingPart::TextRange(id))),
                    // msoAlignLeft 1, Center 2, Right 3.
                    "alignment" => Some(Value::Integer(match shape.h_align {
                        -4108 => 2,
                        -4152 => 3,
                        _ => 1,
                    })),
                    _ => None,
                })
            }
            DrawingPart::Characters(id, start, length) => {
                let shape = self.shape(id)?.clone();
                let at = start.max(1) as usize - 1;
                let length = length.filter(|length| *length != 0);
                let units = shape.comment_at.and_then(|at| self.notes.iter().find(|note| note.at == at).map(Note::units))
                    .unwrap_or_else(|| shape.text().encode_utf16().collect());
                let requested = length.map_or(units.len(), |length| length as usize);
                let end = at.saturating_add(requested).min(units.len());
                let (text, count) = if shape.comment_at.is_some() {
                    let text = if units.is_empty() { Value::String(String::new()) }
                        else if at >= units.len() { Value::Error(2015) }
                        else { Value::from_utf16(units[at..end].to_vec()) };
                    (text, requested)
                } else {
                    if at >= units.len() && !units.is_empty() { return Err(host_error(1004, "Characters starts beyond the text")); }
                    // Drawing text ranges expand a partial surrogate pair to
                    // the complete character; legacy notes expose the units.
                    let original = shape.text();
                    let mut selected = String::new();
                    let mut offset = 0usize;
                    for ch in original.chars() {
                        let next = offset + ch.len_utf16();
                        if offset < end && next > at { selected.push(ch); }
                        offset = next;
                    }
                    let count = selected.encode_utf16().count();
                    (Value::String(selected), count)
                };
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(text),
                    "count" => Some(Value::Integer(count as i64)),
                    "font" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    _ => None,
                })
            }
            DrawingPart::CharactersFont(id, start, length) => {
                let shape = self.shape(id)?.clone();
                let answer = |value: Option<Value>| Some(value.unwrap_or(Value::Null));
                Ok(match lower.as_str() {
                    "name" => answer(shape.uniform_style(start, length, |s| s.name.clone()).map(Value::String)),
                    "size" => answer(shape.uniform_style(start, length, |s| s.size).map(Value::Double)),
                    "bold" => answer(shape.uniform_style(start, length, |s| s.bold).map(Value::Boolean)),
                    "italic" => answer(shape.uniform_style(start, length, |s| s.italic).map(Value::Boolean)),
                    "underline" => answer(shape.uniform_style(start, length, |s| s.underline).map(|u| {
                        Value::Integer(if u { UNDERLINE_SINGLE } else { UNDERLINE_NONE })
                    })),
                    "color" => answer(shape.uniform_style(start, length, |s| s.color).map(Value::Integer)),
                    "colorindex" => answer(
                        shape.uniform_style(start, length, |s| s.color).map(|c| Value::Integer(self.nearest_palette_index(c))),
                    ),
                    "schemecolor" => {
                        let palette = self.indexed_colour_palette();
                        answer(shape.uniform_style(start, length, |s| {
                            scheme_color_for_style_in_palette(
                                s.color_theme,
                                s.scheme_color,
                                s.color,
                                s.color_brightness,
                                &palette,
                            )
                        }).map(Value::Integer))
                    }
                    // `TextRange.Font.Fill.ForeColor.RGB` reaches the colour
                    // the long way round.
                    "fill" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    "forecolor" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    "rgb" => answer(shape.uniform_style(start, length, |s| s.color).map(Value::Integer)),
                    // Measured: a shape's white text is msoThemeColorLight1
                    // (2), a box's black text Dark1 (1).
                    "objectthemecolor" => answer(shape.uniform_style(start, length, |s| s.color_theme)
                        .map(|theme| Value::Integer(theme.map_or(0, |theme| theme as i64)))),
                    "brightness" => answer(shape.uniform_style(start, length, |s| s.color_brightness)
                        .map(|brightness| Value::Single(brightness as f32))),
                    "tintandshade" => Some(Value::Single(0.0)),
                    "type" => answer(shape.uniform_style(start, length, |s| s.color_theme.is_some() || s.scheme_color.is_some())
                        .map(|theme| Value::Integer(if theme { 2 } else { 1 }))),
                    _ => None,
                })
            }
            // Font2 reads the characters' dress and answers it as Office does:
            // measured on fresh text, Bold and Italic msoFalse, Size Single 11,
            // the Latin face for Name and NameAscii, 游ゴシック east, +mn-cs
            // complex, no caps, strike, kerning or spacing, and the
            // TextRange its Parent.
            DrawingPart::Font2(id) => {
                let inner = DrawingPart::CharactersFont(id, 1, None);
                let read = |host: &mut Self, what: &str| host.drawing_get(inner, what);
                Ok(match lower.as_str() {
                    "bold" | "italic" => match read(self, &lower)? {
                        Some(Value::Boolean(held)) => Some(mso(held)),
                        other => other,
                    },
                    "size" => match read(self, "size")? {
                        Some(Value::Double(size)) => Some(Value::Single(size as f32)),
                        other => other,
                    },
                    "name" | "nameascii" => read(self, "name")?,
                    "namefareast" => Some(Value::String("游ゴシック".to_string())),
                    "namecomplexscript" => Some(Value::String("+mn-cs".to_string())),
                    "nameother" => Some(Value::String(String::new())),
                    "allcaps" | "caps" | "doublestrikethrough" | "equalize" | "smallcaps" | "softedgeformat" | "strike"
                    | "strikethrough" | "subscript" | "superscript" => Some(Value::Integer(0)),
                    "underlinestyle" => match read(self, "underline")? {
                        Some(Value::Integer(UNDERLINE_SINGLE)) => Some(Value::Integer(1)),
                        _ => Some(Value::Integer(0)),
                    },
                    "kerning" | "spacing" => Some(Value::Single(0.0)),
                    "parent" => Some(self.part(DrawingPart::TextRange(id))),
                    _ => read(self, &lower)?,
                })
            }
            DrawingPart::Chart(id) => self.chart_get(id, &lower),
            DrawingPart::SeriesCollection(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.chart(id)?.series.len() as i64)),
                "parent" => Some(self.part(DrawingPart::Chart(id))),
                _ => None,
            }),
            DrawingPart::Series(id, number) => {
                let series = self.series(id, number)?.clone();
                Ok(match lower.as_str() {
                    "name" => Some(Value::String(series.name)),
                    "formula" => Some(Value::String(self.series_formula(id, number)?)),
                    "values" => Some(Self::variant_row(series.values.iter().map(|v| Value::Double(*v)).collect())),
                    "xvalues" => Some(Self::variant_row(
                        series.xs.iter().map(|x| match x.parse::<f64>() {
                            Ok(number) => Value::Double(number),
                            Err(_) => Value::String(x.clone()),
                        }).collect(),
                    )),
                    "points" => Some(self.part(DrawingPart::Points(id, number))),
                    "hasdatalabels" => Some(Value::Boolean(series.has_labels)),
                    "datalabels" => Some(self.part(DrawingPart::DataLabels(id, number))),
                    "charttype" => Some(Value::Integer(series.chart_type.unwrap_or(self.chart(id)?.chart_type))),
                    "axisgroup" => Some(Value::Integer(1)),
                    "plotorder" => Some(Value::Integer(number as i64)),
                    // The plain answers of a column series: measured.
                    "barshape" | "explosion" => Some(Value::Integer(0)),
                    "haserrorbars" | "invertifnegative" => Some(Value::Boolean(false)),
                    "hasleaderlines" => Some(Value::Boolean(true)),
                    "markersize" => Some(Value::Integer(5)),
                    "markerstyle" => Some(Value::Integer(-4142)),
                    "picturetype" => Some(Value::Integer(1)),
                    "type" => Some(Value::Integer(3)),
                    "smooth" => return Err(host_error(1004, "a column series has no Smooth")),
                    "format" => Some(self.part(DrawingPart::SeriesFormat(id, number))),
                    "interior" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                    "border" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    _ => None,
                })
            }
            DrawingPart::SeriesFormat(id, number) => Ok(match lower.as_str() {
                "fill" => Some(self.part(DrawingPart::SeriesFill(id, number))),
                "line" => Some(self.part(DrawingPart::SeriesFill(id, number))),
                _ => None,
            }),
            DrawingPart::SeriesFill(id, number) => Ok(match lower.as_str() {
                "forecolor" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                "visible" => Some(mso(true)),
                _ => None,
            }),
            DrawingPart::SeriesColor(id, number) => {
                let series = self.series(id, number)?;
                Ok(match lower.as_str() {
                    "rgb" | "color" => Some(Value::Integer(series.color.unwrap_or(THEME_COLOURS[ACCENT1_INDEX]))),
                    "colorindex" => Some(Value::Integer(self.nearest_palette_index(series.color.unwrap_or(THEME_COLOURS[ACCENT1_INDEX])))),
                    _ => None,
                })
            }
            // A column series' point, measured: S1P2 by name, no explosion,
            // label or shadow, marker 5 and none, the series its Parent;
            // SecondaryPlot 1004 outside a pie of pie.
            DrawingPart::Point(id, number, at) => {
                let series = self.series(id, number)?.clone();
                Ok(match lower.as_str() {
                    "name" => Some(Value::String(format!("S{number}P{at}"))),
                    "explosion" => Some(Value::Integer(0)),
                    "hasdatalabel" => Some(Value::Boolean(series.has_labels)),
                    "invertifnegative" | "shadow" => Some(Value::Boolean(false)),
                    "markersize" => Some(Value::Integer(5)),
                    "markerstyle" => Some(Value::Integer(-4142)),
                    "picturetype" => Some(Value::Integer(1)),
                    "parent" => Some(self.part(DrawingPart::Series(id, number))),
                    "secondaryplot" => return Err(host_error(1004, "the point is in no pie of pie")),
                    _ => None,
                })
            }
            DrawingPart::Points(id, number) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.series(id, number)?.values.len() as i64)),
                _ => None,
            }),
            DrawingPart::DataLabels(id, number) => {
                let series = self.series(id, number)?;
                Ok(match lower.as_str() {
                    "count" => Some(Value::Integer(if series.has_labels { series.values.len() as i64 } else { 0 })),
                    "showvalue" => Some(Value::Boolean(series.has_labels)),
                    // Measured on a column series' labels: 文字列 S1, General
                    // linked to the source, outside the end, the series their
                    // Parent, and nothing shown but the value.
                    "name" => Some(Value::String(format!("文字列 S{number}"))),
                    "numberformat" => Some(Value::String("General".to_string())),
                    "numberformatlinked" => Some(Value::Boolean(true)),
                    "position" => Some(Value::Integer(2)),
                    "parent" => Some(self.part(DrawingPart::Series(id, number))),
                    "showbubblesize" | "showcategoryname" | "showlegendkey" | "showpercentage" | "showseriesname" => {
                        Some(Value::Boolean(false))
                    }
                    _ => None,
                })
            }
            DrawingPart::ChartTitle(id) => {
                let chart = self.chart(id)?.clone();
                let shown = if chart.title.is_empty() {
                    match chart.series.as_slice() {
                        [only] => only.name.clone(),
                        _ => "グラフ タイトル".to_string(),
                    }
                } else {
                    chart.title
                };
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(Value::String(shown)),
                    "characters" => Some(self.part(DrawingPart::ChartTitle(id))),
                    // Measured: タイトル, centred both ways, horizontal, in the
                    // layout at its automatic place, no shadow.
                    "name" => Some(Value::String("タイトル".to_string())),
                    "horizontalalignment" | "verticalalignment" => Some(Value::Integer(-4108)),
                    "orientation" => Some(Value::Integer(-4128)),
                    "readingorder" => Some(Value::Integer(-5002)),
                    "position" => Some(Value::Integer(-4105)),
                    "includeinlayout" => Some(Value::Boolean(true)),
                    "shadow" => Some(Value::Boolean(false)),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    _ => None,
                })
            }
            DrawingPart::Legend(id) => Ok(match lower.as_str() {
                "position" => Some(Value::Integer(self.chart(id)?.legend_position)),
                // Measured: 凡例, in the layout, no shadow, the chart its Parent.
                "name" => Some(Value::String("凡例".to_string())),
                "includeinlayout" => Some(Value::Boolean(true)),
                "shadow" => Some(Value::Boolean(false)),
                "parent" => Some(self.part(DrawingPart::Chart(id))),
                _ => None,
            }),
            DrawingPart::Axes(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(2)),
                "parent" => Some(self.part(DrawingPart::Chart(id))),
                _ => None,
            }),
            DrawingPart::Axis(id, which) => {
                let chart = self.chart(id)?.clone();
                let axis = &chart.axes[usize::from(which == 2)];
                Ok(match lower.as_str() {
                    "hastitle" => Some(Value::Boolean(axis.has_title)),
                    "axistitle" => {
                        if !axis.has_title {
                            return Err(host_error(-2_147_467_259, "the axis has no title"));
                        }
                        Some(self.part(DrawingPart::AxisTitle(id, which)))
                    }
                    "minimumscale" if which != 1 => Some(Value::Double(axis.min.unwrap_or(0.0))),
                    "maximumscale" if which != 1 => Some(Value::Double(axis.max.unwrap_or(0.0))),
                    "minimumscaleisauto" if which != 1 => Some(Value::Boolean(axis.min.is_none())),
                    "maximumscaleisauto" if which != 1 => Some(Value::Boolean(axis.max.is_none())),
                    "categorynames" if which == 1 => Some(Self::variant_row(
                        chart.series.first().map(|s| s.xs.iter().map(|x| Value::String(x.clone())).collect()).unwrap_or_default(),
                    )),
                    "type" => Some(Value::Integer(which)),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    // The value axis of a fresh column chart, measured:
                    // primary, crossing automatically at 0, linear, outside
                    // major ticks, labels next to the axis, major gridlines
                    // only, no display unit; the category members are 440.
                    "axisgroup" if which == 2 => Some(Value::Integer(1)),
                    "crosses" if which == 2 => Some(Value::Integer(-4105)),
                    "crossesat" if which == 2 => Some(Value::Double(0.0)),
                    "displayunit" | "minortickmark" if which == 2 => Some(Value::Integer(-4142)),
                    "hasdisplayunitlabel" | "hasminorgridlines" | "reverseplotorder" if which == 2 => Some(Value::Boolean(false)),
                    "hasmajorgridlines" if which == 2 => Some(Value::Boolean(true)),
                    "majortickmark" if which == 2 => Some(Value::Integer(3)),
                    "scaletype" if which == 2 => Some(Value::Integer(-4132)),
                    "ticklabelposition" if which == 2 => Some(Value::Integer(4)),
                    "axisbetweencategories" | "categorytype" if which == 2 => {
                        return Err(host_error(440, "a value axis has no categories"))
                    }
                    // The category axis, measured: primary, crossing
                    // automatically, between categories, an automatic
                    // category type, outside major ticks, labels next to the
                    // axis, no gridlines; the scale members are 440.
                    "axisgroup" if which == 1 => Some(Value::Integer(1)),
                    "axisbetweencategories" if which == 1 => Some(Value::Boolean(true)),
                    "categorytype" | "crosses" if which == 1 => Some(Value::Integer(-4105)),
                    "crossesat" if which == 1 => Some(Value::Double(0.0)),
                    "hasmajorgridlines" | "hasminorgridlines" | "reverseplotorder" if which == 1 => Some(Value::Boolean(false)),
                    "majortickmark" if which == 1 => Some(Value::Integer(3)),
                    "minortickmark" if which == 1 => Some(Value::Integer(-4142)),
                    "ticklabelposition" if which == 1 => Some(Value::Integer(4)),
                    "displayunit" | "hasdisplayunitlabel" | "maximumscaleisauto" | "minimumscale" | "minimumscaleisauto"
                    | "maximumscale" | "scaletype"
                        if which == 1 =>
                    {
                        return Err(host_error(440, "a category axis has no scale"))
                    }
                    _ => None,
                })
            }
            DrawingPart::AxisTitle(id, which) => {
                let chart = self.chart(id)?;
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(Value::String(chart.axes[usize::from(which == 2)].title.clone())),
                    "characters" => Some(self.part(DrawingPart::AxisTitle(id, which))),
                    _ => None,
                })
            }
            // Measured: the chart area is the object's box less five points
            // each way.
            DrawingPart::ChartArea(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "width" => Some(Value::Double(shape.width - 5.0)),
                    "height" => Some(Value::Double(shape.height - 5.0)),
                    "format" => Some(self.part(DrawingPart::ChartArea(id))),
                    "fill" | "interior" | "forecolor" => Some(self.part(DrawingPart::ChartArea(id))),
                    "rgb" | "color" => Some(Value::Integer(WHITE)),
                    // Measured: グラフ, square corners, no shadow.
                    "name" => Some(Value::String("グラフ".to_string())),
                    "roundedcorners" | "shadow" => Some(Value::Boolean(false)),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    _ => None,
                })
            }
            // Measured on a 200 by 100 chart: the plot's inside left is 16.8.
            // The rest of the plot's frame is not measured and stands in.
            DrawingPart::PlotArea(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "format" | "fill" | "interior" | "forecolor" => Some(self.part(DrawingPart::PlotArea(id))),
                    "rgb" | "color" => Some(Value::Integer(WHITE)),
                    "insideleft" => Some(Value::Double(16.8025196850394)),
                    // Measured: プロット, at its automatic place.
                    "name" => Some(Value::String("プロット".to_string())),
                    "position" => Some(Value::Integer(-4105)),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    "insidetop" => Some(Value::Double(7.2)),
                    "insidewidth" => Some(Value::Double((shape.width - 40.0).max(0.0))),
                    "insideheight" => Some(Value::Double((shape.height - 40.0).max(0.0))),
                    "left" => Some(Value::Double(5.0)),
                    "top" => Some(Value::Double(5.0)),
                    "width" => Some(Value::Double((shape.width - 15.0).max(0.0))),
                    "height" => Some(Value::Double((shape.height - 15.0).max(0.0))),
                    _ => None,
                })
            }
            DrawingPart::Adjustments(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shape(id)?.adjusts.len() as i64)),
                _ => None,
            }),
            // Measured: "a" & Chr(10) & "b" & vbCr & "c" is two paragraphs --
            // a line feed breaks a line, a carriage return a paragraph; and
            // both read back as line feeds.
            DrawingPart::Paragraphs(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shape(id)?.paragraph_count as i64)),
                "text" => Some(Value::String(self.shape(id)?.text())),
                _ => None,
            }),
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                match lower.as_str() {
                    "count" => Ok(Some(Value::Integer(ids.len() as i64))),
                    "shaperange" => Ok(Some(self.part(part))),
                    _ => match ids.first() {
                        Some(id) => self.drawing_get(DrawingPart::Shape(*id), name),
                        None => Ok(None),
                    },
                }
            }
        }
    }

    fn shape_get(&mut self, part: DrawingPart, id: u64, lower: &str) -> Result<Option<Value>, String> {
        let shape = self.shape(id)?.clone();
        if matches!(part, DrawingPart::GroupObject(_)) {
            return Ok(match lower {
                "name" => Some(Value::String(shape.name)),
                "index" => Some(Value::Integer(self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1))),
                "left" => Some(Value::Double(shape.left)),
                "top" => Some(Value::Double(shape.top)),
                "width" => Some(Value::Double(shape.width)),
                "height" => Some(Value::Double(shape.height)),
                "visible" => Some(Value::Boolean(shape.visible)),
                "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
                "application" => Some(self.object(HostObject::Application)),
                "creator" => Some(Value::Integer(1_480_803_660)),
                "shaperange" => Some(self.shape_range_object(vec![id])),
                _ => None,
            });
        }
        if !shape.group_children.is_empty() {
            return Ok(match lower {
                "name" => Some(Value::String(shape.name)),
                "type" => Some(Value::Integer(6)),
                "groupitems" => Some(self.part(DrawingPart::GroupItems(id))),
                "pictureformat" => Some(self.part(DrawingPart::PictureFormat(id))),
                "left" => Some(Value::Single(shape.left as f32)),
                "top" => Some(Value::Single(shape.top as f32)),
                "width" => Some(Value::Single(shape.width as f32)),
                "height" => Some(Value::Single(shape.height as f32)),
                "rotation" => Some(Value::Single(shape.rotation as f32)),
                "child" => Some(mso(true)),
                "visible" => Some(mso(shape.visible)),
                "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
                "shaperange" => Some(self.shape_range_object(vec![id])),
                "application" => Some(self.object(HostObject::Application)),
                "creator" => Some(Value::Integer(1_480_803_660)),
                "index" => Some(Value::Integer(self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1))),
                "parentgroup" => return Err(host_error(1004, "the shape is in no group")),
                "id" => Some(Value::Integer(shape.id as i64 + 1)),
                _ => None,
            });
        }
        if shape.group_parent_id.is_some() {
            match lower {
                "parentgroup" => return Ok(shape.group_parent_id.map(|id| self.part(DrawingPart::Shape(id)))),
                "child" => return Ok(Some(mso(true))),
                _ => {}
            }
        }
        // A ChartObject is the older object: measured, its place is a Double,
        // Visible a Boolean, ZOrder its place, and it prints, has square
        // corners, no shadow and no chart protection.
        if matches!(part, DrawingPart::ChartObject(_)) {
            match lower {
                "left" => return Ok(Some(Value::Double(shape.left))),
                "top" => return Ok(Some(Value::Double(shape.top))),
                "width" => return Ok(Some(Value::Double(shape.width))),
                "height" => return Ok(Some(Value::Double(shape.height))),
                "visible" => return Ok(Some(Value::Boolean(shape.visible))),
                "zorder" => {
                    let at = self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1);
                    return Ok(Some(Value::Integer(at)));
                }
                "printobject" => return Ok(Some(Value::Boolean(true))),
                "protectchartobject" | "roundedcorners" | "shadow" => return Ok(Some(Value::Boolean(false))),
                _ => {}
            }
        }
        Ok(match lower {
            "name" => Some(Value::String(shape.name)),
            // Singles, as Excel hands them: measured, TypeName of a shape's
            // Left, Top, Width, Height and Rotation is Single.
            "left" => Some(Value::Single(shape.left as f32)),
            "top" => Some(Value::Single(shape.top as f32)),
            "width" => Some(Value::Single(shape.width as f32)),
            "height" => Some(Value::Single(shape.height as f32)),
            "rotation" => {
                let group_rotation = shape.group_parent_id
                    .and_then(|parent| self.shape(parent).ok())
                    .map_or(0.0, |parent| parent.rotation);
                Some(Value::Single((shape.rotation + group_rotation) as f32))
            }
            // The plain answers of a shape on its own: measured on an oval,
            // BackgroundStyle 0, BlackWhiteMode 1 (automatic), Child and
            // HasSmartArt msoFalse, Locked True, Title "", and GroupItems
            // and ParentGroup 1004 outside a group.
            "backgroundstyle" => Some(Value::Integer(0)),
            "blackwhitemode" => Some(Value::Integer(1)),
            "child" | "hassmartart" => Some(mso(false)),
            "locked" => Some(Value::Boolean(true)),
            "title" => Some(Value::String(String::new())),
            "groupitems" | "parentgroup" => return Err(host_error(1004, "the shape is in no group")),
            "pictureformat" => Some(self.part(DrawingPart::PictureFormat(id))),
            "application" => Some(Value::String("Microsoft Excel".to_string())),
            "visible" => Some(mso(shape.visible)),
            "placement" => Some(Value::Integer(shape.placement)),
            "onaction" => Some(Value::String(shape.on_action)),
            "alternativetext" => Some(Value::String(shape.alt_text)),
            "lockaspectratio" => Some(mso(shape.lock_aspect)),
            "horizontalflip" => Some(mso(shape.flip_h)),
            "verticalflip" => Some(mso(shape.flip_v)),
            "id" => Some(Value::Integer(shape.id as i64 + 1)),
            "zorderposition" => Some(Value::Integer(
                self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1),
            )),
            // msoAutoShape 1, msoChart 3, msoLine 9, msoPicture 13, msoTextBox 17.
            "connector" => Some(mso(matches!(shape.kind, ShapeKind::Line) || shape.connector != 0)),
            "type" if shape.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some()) => Some(Value::Integer(7)),
            "type" => Some(Value::Integer(match shape.kind {
                _ if shape.comment_at.is_some() => 4,
                _ if shape.connector != 0 => 1,
                ShapeKind::Auto(_) => 1,
                ShapeKind::TextBox => 17,
                ShapeKind::Line => 9,
                ShapeKind::Chart(_) => 3,
                ShapeKind::Picture => 13,
                ShapeKind::Other => 24,
            })),
            // Measured: a connector's AutoShapeType is msoShapeMixed, -2.
            "autoshapetype" if shape.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some()) => Some(Value::Integer(-2)),
            "autoshapetype" => Some(Value::Integer(match shape.kind {
                _ if shape.comment_at.is_some() => 1,
                ShapeKind::Auto(kind) => kind,
                ShapeKind::TextBox => 1,
                _ => -2,
            })),
            "haschart" => Some(mso(matches!(shape.kind, ShapeKind::Chart(_)))),
            "adjustments" => Some(self.part(DrawingPart::Adjustments(id))),
            "oleformat" if shape.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some()) => {
                Some(self.part(DrawingPart::OleFormat(id)))
            }
            "oleformat" => return Err(host_error(438, "Object doesn't support this property or method")),
            "fill" => Some(self.part(DrawingPart::Fill(id))),
            "line" | "border" => Some(self.part(DrawingPart::Line(id))),
            "textframe" => Some(self.part(DrawingPart::TextFrame(id))),
            "textframe2" => Some(self.part(DrawingPart::TextFrame2(id))),
            "chart" => {
                if !matches!(shape.kind, ShapeKind::Chart(_)) {
                    return Err(host_error(1004, "the shape holds no chart"));
                }
                Some(self.part(DrawingPart::Chart(id)))
            }
            "topleftcell" => {
                let at = self.cell_under(shape.sheet, shape.left, shape.top);
                Some(self.object(HostObject::Range(CellRange::single(at))))
            }
            "bottomrightcell" => {
                let at = self.cell_under(shape.sheet, shape.left + shape.width, shape.top + shape.height);
                Some(self.object(HostObject::Range(CellRange::single(at))))
            }
            "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
            "shaperange" => Some(self.part(part)),
            "index" => Some(Value::Integer(
                self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1),
            )),
            _ => None,
        })
    }

    fn chart_get(&mut self, id: u64, lower: &str) -> Result<Option<Value>, String> {
        let shape = self.shape(id)?.clone();
        let chart = self.chart(id)?.clone();
        Ok(match lower {
            "charttype" => Some(Value::Integer(chart.chart_type)),
            "hastitle" => Some(Value::Boolean(chart.has_title)),
            // Measured: `ChartTitle.Text` reads the one series' name on a
            // chart whose HasTitle is False.
            "charttitle" => Some(self.part(DrawingPart::ChartTitle(id))),
            "haslegend" => Some(Value::Boolean(chart.has_legend)),
            "legend" => Some(self.part(DrawingPart::Legend(id))),
            "seriescollection" => Some(self.part(DrawingPart::SeriesCollection(id))),
            "axes" => Some(self.part(DrawingPart::Axes(id))),
            "chartarea" => Some(self.part(DrawingPart::ChartArea(id))),
            "plotarea" => Some(self.part(DrawingPart::PlotArea(id))),
            "chartstyle" => Some(Value::Integer(chart.style)),
            // Measured: the chart is `Sheet1 グラフ 5` for the object `Chart 5`
            // -- the sheet, the Japanese word, and the number.
            "name" => Some(Value::String(format!(
                "{} グラフ {}",
                self.workbook.sheets[shape.sheet].name,
                shape.name.rsplit(' ').next().unwrap_or("")
            ))),
            "parent" => Some(self.part(DrawingPart::ChartObject(id))),
            // An embedded chart's own Visible and CodeName are 1004, and
            // HasAxis wants its axis: measured, 440 asked bare.
            "visible" | "codename" => return Err(host_error(1004, "an embedded chart has no such property")),
            "hasaxis" => return Err(host_error(440, "HasAxis wants an axis")),
            // The plain answers of a fresh column chart: measured.
            "autoscaling" | "plotvisibleonly" => Some(Value::Boolean(true)),
            "hasdatatable" | "rightangleaxes" | "showallfieldbuttons" | "showdatalabelsovermaximum" => Some(Value::Boolean(false)),
            "barshape" | "plotby" => Some(Value::Integer(0)),
            "depthpercent" | "heightpercent" => Some(Value::Integer(100)),
            "displayblanksas" => Some(Value::Integer(1)),
            "elevation" => Some(Value::Integer(15)),
            "gapdepth" => Some(Value::Integer(150)),
            "perspective" => Some(Value::Integer(30)),
            "rotation" => Some(Value::Integer(20)),
            _ => None,
        })
    }

    // ---- writing ------------------------------------------------------------

    pub(super) fn drawing_set(&mut self, part: DrawingPart, name: &str, value: Value) -> Result<bool, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Crop(id) => {
                let requested = any_number(&value).ok_or_else(|| host_error(1004, "Invalid crop value"))? as f32;
                let requested = f64::from(requested);
                if !requested.is_finite() {
                    return Err(host_error(1004, "Invalid crop value"));
                }
                let shape_before = self.shape(id)?.clone();
                let (picture_w, picture_h, crop_left, crop_top, crop_right, crop_bottom) = crop_picture_geometry(&shape_before)?;
                let offset_x = (crop_right - crop_left) / 2.0;
                let offset_y = (crop_bottom - crop_top) / 2.0;
                let (mut left, mut top, mut width, mut height) = (shape_before.left, shape_before.top, shape_before.width, shape_before.height);
                let (mut new_left_crop, mut new_top_crop, mut new_right_crop, mut new_bottom_crop) = (crop_left, crop_top, crop_right, crop_bottom);
                let (mut new_picture_w, mut new_picture_h) = (picture_w, picture_h);
                match lower.as_str() {
                    "picturewidth" if requested > 0.0 => {
                        new_picture_w = requested;
                        new_left_crop = (requested - width) / 2.0 - offset_x;
                        new_right_crop = (requested - width) / 2.0 + offset_x;
                    }
                    "pictureheight" if requested > 0.0 => {
                        new_picture_h = requested;
                        new_top_crop = (requested - height) / 2.0 - offset_y;
                        new_bottom_crop = (requested - height) / 2.0 + offset_y;
                    }
                    "pictureoffsetx" => {
                        new_left_crop = (picture_w - width) / 2.0 - requested;
                        new_right_crop = (picture_w - width) / 2.0 + requested;
                    }
                    "pictureoffsety" => {
                        new_top_crop = (picture_h - height) / 2.0 - requested;
                        new_bottom_crop = (picture_h - height) / 2.0 + requested;
                    }
                    "shapewidth" if requested > 0.0 => {
                        width = requested;
                        new_right_crop = picture_w - width - new_left_crop;
                    }
                    "shapeheight" if requested > 0.0 => {
                        height = requested;
                        new_bottom_crop = picture_h - height - new_top_crop;
                    }
                    "shapeleft" => {
                        new_left_crop += requested - left;
                        left = requested;
                        new_right_crop = picture_w - width - new_left_crop;
                    }
                    "shapetop" => {
                        new_top_crop += requested - top;
                        top = requested;
                        new_bottom_crop = picture_h - height - new_top_crop;
                    }
                    "picturewidth" | "pictureheight" | "shapewidth" | "shapeheight" => {
                        return Err(host_error(1004, "Invalid picture dimensions"));
                    }
                    _ => return Ok(false),
                }
                let shape = self.shape_mut(id)?;
                shape.left = left;
                shape.top = top;
                shape.width = width;
                shape.height = height;
                let Some(DrawingKind::Picture { format, .. }) = shape.original.as_mut().map(|drawing| &mut drawing.kind) else {
                    return Err(host_error(438, "Object doesn't support this property or method"));
                };
                format.crop = (
                    (new_left_crop / new_picture_w * 100_000.0).round() as i32,
                    (new_top_crop / new_picture_h * 100_000.0).round() as i32,
                    (new_right_crop / new_picture_w * 100_000.0).round() as i32,
                    (new_bottom_crop / new_picture_h * 100_000.0).round() as i32,
                );
                self.redraw(id)?;
                Ok(true)
            }
            DrawingPart::PictureFormat(id) => {
                let crop_extent = self.shape(id)?.original.as_ref().and_then(|drawing| match &drawing.kind {
                    DrawingKind::Picture { bytes, format } => picture_source_size_points(bytes, format),
                    _ => None,
                }).unwrap_or_else(|| {
                    let shape = self.shape(id).expect("shape was resolved above");
                    (shape.width, shape.height)
                });
                let taken = {
                    let shape = self.shape_mut(id)?;
                    let Some(oxicells_core::ir::Drawing { kind: DrawingKind::Picture { format, .. }, .. }) = shape.original.as_mut() else {
                        if matches!(lower.as_str(), "brightness" | "contrast" | "colortype" | "cropleft" | "croptop" | "cropright" | "cropbottom" | "transparencycolor" | "transparentbackground") {
                            return Err(host_error(438, "Object doesn't support this property or method"));
                        }
                        return Ok(false);
                    };
                    match lower.as_str() {
                        "brightness" | "contrast" => {
                            let narrowed = f64::from(transparency_number(&value)? as f32);
                            if !(0.0..=1.0).contains(&narrowed) {
                                return Err(host_error(1004, "Invalid picture adjustment"));
                            }
                            let raw = ((narrowed - 0.5) * 200_000.0).round() as i32;
                            if lower == "brightness" { format.brightness = raw; } else { format.contrast = raw; }
                            if format.brightness == 70_000 && format.contrast == -70_000 { format.color_type = 4; }
                            else if format.color_type == 4 { format.color_type = 1; }
                            true
                        }
                        "colortype" => {
                            let kind = excel_enum_variant(&value, 1, 4, "Invalid picture color type")
                                .map_err(|_| host_error(1004, "Invalid picture color type"))?;
                            format.color_type = kind;
                            match kind {
                                4 => { format.brightness = 70_000; format.contrast = -70_000; }
                                _ => { format.brightness = 0; format.contrast = 0; }
                            }
                            true
                        }
                        "cropleft" | "croptop" | "cropright" | "cropbottom" => {
                            let requested = transparency_number(&value)?;
                            if !requested.is_finite() {
                                return Err(host_error(1004, "Invalid crop value"));
                            }
                            let axis = if matches!(lower.as_str(), "cropleft" | "cropright") { 0 } else { 1 };
                            let extent = if axis == 0 { crop_extent.0 } else { crop_extent.1 };
                            if !extent.is_finite() || extent <= 0.0 {
                                return Err(host_error(1004, "Invalid picture dimensions"));
                            }
                            let points = requested.min(extent);
                            let raw = ((points / extent) * 100_000.0).round() as i32;
                            match lower.as_str() {
                                "cropleft" => format.crop.0 = raw,
                                "croptop" => format.crop.1 = raw,
                                "cropright" => format.crop.2 = raw,
                                _ => format.crop.3 = raw,
                            }
                            true
                        }
                        "transparencycolor" => {
                            let number = transparency_number(&value)?;
                            if !number.is_finite() {
                                return Err(host_error(1004, "Invalid transparency color"));
                            }
                            let rounded = number.round_ties_even();
                            if !(i32::MIN as f64..=i32::MAX as f64).contains(&rounded) {
                                return Err(host_error(1004, "Invalid transparency color"));
                            }
                            let packed = rounded as i32;
                            // Excel coerces the sentinel Long value -2147483648
                            // to black (0) rather than removing clrChange.
                            let packed = if packed == i32::MIN { 0 } else { packed };
                            format.transparency_color = Some(colour_from_packed(i64::from(packed)));
                            true
                        }
                        "transparentbackground" => {
                            format.transparent_background = mso_asked(&value, "PictureFormat.TransparentBackground")?;
                            true
                        }
                        _ => false,
                    }
                };
                if taken { self.redraw(id)?; }
                Ok(taken)
            }
            DrawingPart::OleBorder(id) | DrawingPart::OleInterior(id) => {
                let border = matches!(&part, DrawingPart::OleBorder(_));
                let parsed_color = if lower == "color" {
                    let number = if matches!(value, Value::Null) {
                        if border { Some(16_711_935.0) } else { Some(0.0) }
                    } else { Some(transparency_number(&value)?) };
                    let number = number.expect("OLE color coercion supplies a number");
                    let rounded = number.round_ties_even();
                    if !number.is_finite() || !(i32::MIN as f64..=i32::MAX as f64).contains(&rounded) {
                        return Err(oxivba_core::host_error_described(6, "オーバーフローしました。"));
                    }
                    if border && rounded == i32::MAX as f64 {
                        return Err(oxivba_core::host_error_described(1004, "未知の実行時エラーです。"));
                    }
                    let packed = rounded as i32;
                    let normalized = if border && packed == -1 { 16_711_935 } else { packed & 0x00FF_FFFF };
                    let normalized = if border {
                        let palette = self.indexed_colour_palette();
                        let index = self.nearest_palette_index(i64::from(normalized));
                        if (1..=56).contains(&index) { palette[index as usize - 1] } else { normalized.into() }
                    } else { normalized.into() };
                    Some(colour_from_packed(normalized))
                } else { None };
                let parsed_weight = if lower == "weight" {
                    let invalid = || oxivba_core::host_error_described(1004, "Border クラスの Weight プロパティを設定できません。");
                    if matches!(value, Value::Null) { return Err(invalid()); }
                    let number = transparency_number(&value)?;
                    let rounded = number.round_ties_even();
                    if !number.is_finite() || !(i64::MIN as f64..=i64::MAX as f64).contains(&rounded) {
                        return Err(invalid());
                    }
                    let rounded = rounded as i64;
                    let weight = match rounded { 1 | 2 | 4 | -4138 => rounded, 3 => -4138, _ => return Err(invalid()) };
                    Some(weight)
                } else { None };
                let valid = if border { lower == "color" || lower == "weight" } else { lower == "color" };
                if !valid { return Ok(false); }
                let sheet = self.shape(id)?.sheet;
                let drawing = self.shape_mut(id)?.original.as_mut()
                    .filter(|drawing| drawing.ole_prog_id.is_some())
                    .ok_or_else(|| host_error(438, "Object doesn't support this property or method"))?;
                if let Some(color) = parsed_color {
                    if border { drawing.ole_border_color = Some(color); }
                    else { drawing.ole_interior_color = Some(color); }
                }
                if let Some(weight) = parsed_weight { drawing.ole_border_weight = Some(weight); }
                self.mirror_drawings(sheet);
                Ok(true)
            }
            DrawingPart::OleObject(id) => {
                if matches!(lower.as_str(), "autoload" | "enabled" | "locked" | "printobject" | "shadow") {
                    let Value::Boolean(flag) = value else { return Err(host_error(13, "Type mismatch")); };
                    let sheet = self.shape(id)?.sheet;
                    {
                        let drawing = self.shape_mut(id)?.original.as_mut()
                            .filter(|drawing| drawing.ole_prog_id.is_some())
                            .ok_or_else(|| host_error(438, "Object doesn't support this property or method"))?;
                        match lower.as_str() {
                            "autoload" => drawing.ole_auto_load = Some(flag),
                            "enabled" => drawing.ole_enabled = Some(flag),
                            "locked" => drawing.ole_locked = Some(flag),
                            "printobject" => drawing.ole_print_object = Some(flag),
                            "shadow" => drawing.ole_shadow = Some(flag),
                            _ => unreachable!(),
                        }
                    }
                    self.mirror_drawings(sheet);
                    Ok(true)
                } else if matches!(lower.as_str(), "name" | "left" | "top" | "width" | "height" | "placement" | "visible" | "onaction") {
                    if matches!(lower.as_str(), "width" | "height") {
                        if let Some(requested) = any_number(&value) {
                            let shape = self.shape(id)?;
                            let (old_width, old_height) = (shape.width, shape.height);
                            let requested = ((requested.max(0.0) / 0.75).round() * 0.75).max(0.0);
                            let (width, height) = if lower == "width" {
                                let height = if old_width > 0.0 { requested * old_height / old_width } else { old_height };
                                (requested, ((height / 0.75).round() * 0.75).max(0.0))
                            } else {
                                let width = if old_height > 0.0 { requested * old_width / old_height } else { old_width };
                                (((width / 0.75).round() * 0.75).max(0.0), requested)
                            };
                            self.drawing_set(DrawingPart::Shape(id), "width", Value::Double(width))?;
                            self.drawing_set(DrawingPart::Shape(id), "height", Value::Double(height))
                        } else {
                            self.drawing_set(DrawingPart::Shape(id), name, value)
                        }
                    } else if matches!(lower.as_str(), "left" | "top") {
                        if let Some(requested) = any_number(&value) {
                            let requested = (requested.max(0.0) / 0.75).round() * 0.75;
                            self.drawing_set(DrawingPart::Shape(id), name, Value::Double(requested))
                        } else {
                            self.drawing_set(DrawingPart::Shape(id), name, value)
                        }
                    } else {
                        self.drawing_set(DrawingPart::Shape(id), name, value)
                    }
                } else {
                    Ok(false)
                }
            }
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) | DrawingPart::GroupObject(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "name" => {
                            let Value::String(named) = &value else {
                                return Err("Shape.Name takes text".to_string());
                            };
                            shape.name = named.clone();
                            true
                        }
                        "left" | "top" | "width" | "height" | "rotation" => {
                            let Some(number) = any_number(&value) else {
                                return Err(format!("Shape.{name} takes a number"));
                            };
                            match lower.as_str() {
                                "left" => shape.left = number.max(0.0),
                                "top" => shape.top = number.max(0.0),
                                "width" => shape.width = number.max(0.0),
                                "height" => shape.height = number.max(0.0),
                                _ => shape.rotation = number.rem_euclid(360.0),
                            }
                            true
                        }
                        "visible" => {
                            shape.visible = mso_asked(&value, "Shape.Visible")?;
                            true
                        }
                        "placement" => {
                            shape.placement = any_whole_number(&value).unwrap_or(1);
                            true
                        }
                        "onaction" => {
                            shape.on_action = match &value {
                                Value::String(text) => text.clone(),
                                _ => String::new(),
                            };
                            true
                        }
                        "alternativetext" => {
                            shape.alt_text = match &value {
                                Value::String(text) => text.clone(),
                                _ => String::new(),
                            };
                            true
                        }
                        "lockaspectratio" => {
                            shape.lock_aspect = mso_asked(&value, "Shape.LockAspectRatio")?;
                            true
                        }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Fill(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "visible" => {
                            shape.fill_visible = mso_asked(&value, "FillFormat.Visible")?;
                            true
                        }
                        "transparency" => {
                            shape.transparency = excel_shape_transparency(transparency_number(&value)?)?;
                            true
                        }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::FillColor(id) | DrawingPart::LineColor(id) | DrawingPart::LineBackColor(id) => {
                let fill = matches!(part, DrawingPart::FillColor(_));
                let back = matches!(part, DrawingPart::LineBackColor(_));
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "rgb" => {
                            let Some(colour) = color_number(&value, "ColorFormat.RGB")? else {
                                return Ok(true);
                            };
                            if fill {
                                shape.fill = colour as i64;
                                shape.fill_theme = None;
                                shape.fill_scheme_color = None;
                                shape.fill_brightness = 0.0;
                                shape.fill_visible = true;
                            } else if back {
                                if shape.line_pattern != -2 {
                                    shape.line_back = colour as i64;
                                    shape.line_back_theme = None;
                                    shape.line_back_scheme_color = None;
                                    shape.line_back_brightness = 0.0;
                                    shape.line_visible = true;
                                }
                            } else {
                                shape.line = colour as i64;
                                shape.line_theme = None;
                                shape.line_scheme_color = None;
                                shape.line_brightness = 0.0;
                                shape.line_shade = None;
                                shape.line_visible = true;
                            }
                            true
                        }
                        "objectthemecolor" => {
                            let theme = theme_colour(&value)?;
                            if fill {
                                shape.fill = office_theme_colour_from(&shape.theme_palette, theme);
                                shape.fill_theme = Some(theme);
                                shape.fill_scheme_color = None;
                                shape.fill_brightness = 0.0;
                            } else if back {
                                if shape.line_pattern != -2 {
                                    shape.line_back = office_theme_colour_from(&shape.theme_palette, theme);
                                    shape.line_back_theme = Some(theme);
                                    shape.line_back_scheme_color = None;
                                    shape.line_back_brightness = 0.0;
                                    shape.line_visible = true;
                                }
                            } else {
                                shape.line = office_theme_colour_from(&shape.theme_palette, theme);
                                shape.line_theme = Some(theme);
                                shape.line_scheme_color = None;
                                shape.line_brightness = 0.0;
                                shape.line_shade = None;
                            }
                            true
                        }
                        "schemecolor" => {
                            let index = legacy_scheme_color(&value)?;
                            let color = LEGACY_SCHEME_RGB[index as usize];
                            if fill {
                                shape.fill = color;
                                shape.fill_theme = None;
                                shape.fill_scheme_color = Some(index);
                                shape.fill_brightness = 0.0;
                                shape.fill_visible = true;
                            } else if back {
                                if shape.line_pattern != -2 {
                                    shape.line_back = color;
                                    shape.line_back_theme = None;
                                    shape.line_back_scheme_color = Some(index);
                                    shape.line_back_brightness = 0.0;
                                    shape.line_visible = true;
                                }
                            } else {
                                shape.line = color;
                                shape.line_theme = None;
                                shape.line_scheme_color = Some(index);
                                shape.line_brightness = 0.0;
                                shape.line_shade = None;
                                shape.line_visible = true;
                            }
                            true
                        }
                        "brightness" => {
                            let brightness = brightness_asked(&value)?;
                            if fill {
                                let base = shape.fill_theme.map_or(shape.fill, |theme| office_theme_colour_from(&shape.theme_palette, theme));
                                shape.fill = color_format_brightness(base, brightness);
                                shape.fill_brightness = brightness;
                                shape.fill_scheme_color = None;
                            } else if back {
                                if shape.line_pattern != -2 {
                                    let base = shape.line_back_theme.map_or(shape.line_back, |theme| office_theme_colour_from(&shape.theme_palette, theme));
                                    shape.line_back = color_format_brightness(base, brightness);
                                    shape.line_back_brightness = brightness;
                                    shape.line_back_scheme_color = None;
                                }
                            } else {
                                let base = shape.line_theme.map_or(shape.line, |theme| office_theme_colour_from(&shape.theme_palette, theme));
                                shape.line = color_format_brightness(base, brightness);
                                shape.line_brightness = brightness;
                                shape.line_scheme_color = None;
                                shape.line_shade = None;
                            }
                            true
                        }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Line(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "visible" => {
                            shape.line_visible = mso_asked(&value, "LineFormat.Visible")?;
                            true
                        }
                        "weight" => {
                            shape.line_weight = excel_line_weight(transparency_number(&value)?)?;
                            shape.line_visible = true;
                            true
                        }
                        "dashstyle" => {
                            shape.dash = excel_line_dash_style(&value)?;
                            true
                        }
                        "pattern" => {
                            shape.line_pattern = excel_enum_variant(&value, 1, 54, "Invalid line pattern")?;
                            shape.line = 0;
                            shape.line_theme = None;
                            shape.line_scheme_color = None;
                            shape.line_shade = None;
                            true
                        }
                        "endarrowheadstyle" => {
                            shape.arrow_end = excel_arrowhead_style(&value)?;
                            true
                        }
                        "transparency" => {
                            shape.line_transparency = excel_shape_transparency(transparency_number(&value)?)?;
                            true
                        }
                        "beginarrowheadstyle" => {
                            shape.arrow_start = excel_arrowhead_style(&value)?;
                            true
                        }
                        "beginarrowheadlength" => {
                            shape.arrow_start_length = Some(excel_enum_variant(&value, 1, 3, "Invalid arrowhead length")?);
                            true
                        }
                        "beginarrowheadwidth" => {
                            shape.arrow_start_width = Some(excel_enum_variant(&value, 1, 3, "Invalid arrowhead width")?);
                            true
                        }
                        "endarrowheadlength" => {
                            shape.arrow_end_length = Some(excel_enum_variant(&value, 1, 3, "Invalid arrowhead length")?);
                            true
                        }
                        "endarrowheadwidth" => {
                            shape.arrow_end_width = Some(excel_enum_variant(&value, 1, 3, "Invalid arrowhead width")?);
                            true
                        }
                        "style" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::TextFrame(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "horizontalalignment" => {
                            shape.h_align = any_whole_number(&value).unwrap_or(-4131);
                            true
                        }
                        "verticalalignment" => {
                            shape.v_align = any_whole_number(&value).unwrap_or(-4160);
                            true
                        }
                        "autosize" => {
                            shape.auto_size = style_face_boolean(&value, "TextFrame.AutoSize")?.unwrap_or(false);
                            if shape.auto_size {
                                shape.fit_to_text();
                            }
                            true
                        }
                        "marginleft" => { shape.margins.0 = any_number(&value).unwrap_or(7.2); true }
                        "margintop" => { shape.margins.1 = any_number(&value).unwrap_or(3.6); true }
                        "marginright" => { shape.margins.2 = any_number(&value).unwrap_or(7.2); true }
                        "marginbottom" => { shape.margins.3 = any_number(&value).unwrap_or(3.6); true }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::TextFrame2(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "verticalanchor" => {
                            shape.v_align = match any_whole_number(&value) {
                                Some(3) => -4108,
                                Some(4) => -4107,
                                _ => -4160,
                            };
                            true
                        }
                        "wordwrap" | "autosize" => true,
                        "marginleft" => { shape.margins.0 = any_number(&value).unwrap_or(7.2); true }
                        "margintop" => { shape.margins.1 = any_number(&value).unwrap_or(3.6); true }
                        "marginright" => { shape.margins.2 = any_number(&value).unwrap_or(7.2); true }
                        "marginbottom" => { shape.margins.3 = any_number(&value).unwrap_or(3.6); true }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::ParagraphFormat(id) => {
                if lower == "alignment" {
                    let shape = self.shape_mut(id)?;
                    shape.h_align = match any_whole_number(&value) {
                        Some(2) => -4108,
                        Some(3) => -4152,
                        _ => -4131,
                    };
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::TextRange(id) | DrawingPart::Characters(id, ..) => {
                if lower == "text" || lower == "caption" {
                    let text = match &value {
                        Value::String(text) => text.clone(),
                        other => text_of_value(other),
                    };
                    match part {
                        DrawingPart::Characters(_, start, length) => self.replace_shape_characters(id, start, length, &value)?,
                        _ => { self.shape_mut(id)?.set_text(&text); self.redraw(id)?; },
                    }
                    return Ok(true);
                }
                Ok(false)
            }
            // Font2 writes the characters' dress; msoTrue (-1) is True there.
            DrawingPart::Font2(id) => self.drawing_set(DrawingPart::CharactersFont(id, 1, None), name, value),
            DrawingPart::CharactersFont(id, start, length) => {
                let palette = self.indexed_colour_palette();
                let taken = {
                    let shape = self.shape_mut(id)?;
                    if shape.runs.is_empty() {
                        // Dress the text a macro will write: the style is kept
                        // on an empty run until the text arrives.
                        let style = shape.default_style();
                        shape.runs.push(ShapeRunRecord { text: String::new(), style, paragraph_end: false });
                    }
                    let (first, last) = shape.split_runs(start, length);
                    let mut taken = true;
                    for run in &mut shape.runs[first..last] {
                        match lower.as_str() {
                            "name" => {
                                if let Value::String(named) = &value {
                                    run.style.name = named.clone();
                                }
                            }
                            "size" => run.style.size = any_number(&value).unwrap_or(run.style.size),
                            "bold" => run.style.bold = style_face_boolean(&value, "Font.Bold")?.unwrap_or(false),
                            "italic" => run.style.italic = style_face_boolean(&value, "Font.Italic")?.unwrap_or(false),
                            "underline" => {
                                run.style.underline = match &value {
                                    Value::Boolean(flag) => *flag,
                                    other => any_whole_number(other).is_some_and(|n| n != UNDERLINE_NONE && n != 0),
                                }
                            }
                            "color" | "rgb" => {
                                if let Some(colour) = color_number(&value, "Font.Color")? {
                                    run.style.color = colour as i64;
                                    run.style.color_theme = None;
                                    run.style.color_brightness = 0.0;
                                    run.style.scheme_color = None;
                                }
                            }
                            "colorindex" => {
                                if let Some(colour) = palette_choice_in(&value, COLOUR_AUTOMATIC, "Font.ColorIndex", &palette)? {
                                    run.style.color = colour_to_packed(Some(&colour)).unwrap_or(0);
                                    run.style.color_theme = None;
                                    run.style.color_brightness = 0.0;
                                    run.style.scheme_color = None;
                                }
                            }
                            "objectthemecolor" => {
                                let theme = theme_colour(&value)?;
                                run.style.color = office_theme_colour_from(&shape.theme_palette, theme);
                                run.style.color_theme = Some(theme);
                                run.style.color_brightness = 0.0;
                                run.style.scheme_color = None;
                            }
                            "schemecolor" => {
                                let index = legacy_scheme_color(&value)?;
                                run.style.color = LEGACY_SCHEME_RGB[index as usize];
                                run.style.color_theme = None;
                                run.style.color_brightness = 0.0;
                                run.style.scheme_color = Some(index);
                            }
                            "brightness" => {
                                let brightness = brightness_asked(&value)?;
                                let base = run.style.color_theme.map_or(run.style.color, |theme| office_theme_colour_from(&shape.theme_palette, theme));
                                run.style.color = color_format_brightness(base, brightness);
                                run.style.color_brightness = brightness;
                                run.style.scheme_color = None;
                            }
                            _ => taken = false,
                        }
                    }
                    taken
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Chart(id) => {
                let taken = {
                    let chart = self.chart_mut(id)?;
                    match lower.as_str() {
                        "charttype" => {
                            let Some(kind) = any_whole_number(&value) else {
                                return Err("Chart.ChartType takes an xlChartType".to_string());
                            };
                            chart.chart_type = kind;
                            for series in &mut chart.series {
                                series.chart_type = None;
                            }
                            if chart.title_auto && chart_is_pie(kind) && chart.series.len() == 1 {
                                chart.has_title = true;
                            }
                            chart.title_auto = false;
                            true
                        }
                        "hastitle" => {
                            chart.has_title = style_face_boolean(&value, "Chart.HasTitle")?.unwrap_or(false);
                            if !chart.has_title {chart.title.clear();}
                            if chart.has_title && chart.title.is_empty() {
                                chart.title = match chart.series.as_slice() {
                                    [only] => only.name.clone(),
                                    _ => "グラフ タイトル".to_string(),
                                };
                            }
                            true
                        }
                        "haslegend" => {
                            chart.has_legend = style_face_boolean(&value, "Chart.HasLegend")?.unwrap_or(false);
                            true
                        }
                        "chartstyle" => {
                            chart.style = any_whole_number(&value).unwrap_or(chart.style);
                            true
                        }
                        "name" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Series(id, number) => {
                let taken = match lower.as_str() {
                    "name" => {
                        let series = self.series_mut(id, number)?;
                        match &value {
                            // `="S2"` is a name written as a formula.
                            Value::String(text) => {
                                let bare = text.strip_prefix('=').unwrap_or(text);
                                series.name = bare.trim_matches('"').to_string();
                                series.name_ref = None;
                            }
                            Value::Object(_) => {
                                if let Some(range) = self.cells_of_range_object(&value) {
                                    let text = self.range_texts(range).into_iter().next().unwrap_or_default();
                                    let reference = self.address_of(range);
                                    let series = self.series_mut(id, number)?;
                                    series.name = text;
                                    series.name_ref = Some(reference);
                                }
                            }
                            other => {
                                let text = text_of_value(other);
                                self.series_mut(id, number)?.name = text;
                            }
                        }
                        true
                    }
                    "values" => {
                        if let Some(range) = self.cells_of_range_object(&value) {
                            let numbers = self.range_numbers(range);
                            let reference = self.address_of(range);
                            let series = self.series_mut(id, number)?;
                            series.values = numbers;
                            series.values_ref = Some(reference);
                            if series.xs.len() != series.values.len() {
                                series.xs = (1..=series.values.len()).map(|n| n.to_string()).collect();
                            }
                        } else if let Value::Array(listed) = &value {
                            let series = self.series_mut(id, number)?;
                            series.values = listed.values.iter().map(|v| any_number(v).unwrap_or(0.0)).collect();
                            series.values_ref = None;
                        }
                        true
                    }
                    "xvalues" => {
                        if let Some(range) = self.cells_of_range_object(&value) {
                            let texts = self.range_texts(range);
                            let reference = self.address_of(range);
                            let series = self.series_mut(id, number)?;
                            series.xs = texts;
                            series.x_ref = Some(reference);
                        } else if let Value::Array(listed) = &value {
                            let series = self.series_mut(id, number)?;
                            series.xs = listed.values.iter().map(text_of_value).collect();
                            series.x_ref = None;
                        }
                        true
                    }
                    "hasdatalabels" => {
                        let flag = style_face_boolean(&value, "Series.HasDataLabels")?.unwrap_or(false);
                        self.series_mut(id, number)?.has_labels = flag;
                        true
                    }
                    // Measured: one series' ChartType is the chart's when it
                    // is the only one.
                    "charttype" => {
                        let kind = any_whole_number(&value).unwrap_or(51);
                        let single = self.chart(id)?.series.len() == 1;
                        self.series_mut(id, number)?.chart_type = Some(kind);
                        if single {
                            self.chart_mut(id)?.chart_type = kind;
                        }
                        true
                    }
                    "axisgroup" | "plotorder" | "smooth" | "markerstyle" | "markersize" => true,
                    _ => false,
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::SeriesColor(id, number) => {
                if lower == "rgb" || lower == "color" {
                    if let Some(colour) = color_number(&value, "ColorFormat.RGB")? {
                        self.series_mut(id, number)?.color = Some(colour as i64);
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                if lower == "colorindex" {
                    if let Some(colour) = self.palette_choice(&value, COLOUR_AUTOMATIC, "Interior.ColorIndex")? {
                        self.series_mut(id, number)?.color = colour_to_packed(Some(&colour));
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::SeriesFill(..) | DrawingPart::SeriesFormat(..) => {
                Ok(matches!(lower.as_str(), "visible" | "transparency" | "weight"))
            }
            DrawingPart::DataLabels(id, number) => {
                if matches!(lower.as_str(), "showvalue" | "showcategoryname" | "showpercentage" | "position" | "numberformat") {
                    if lower == "showvalue" {
                        let flag = style_face_boolean(&value, "DataLabels.ShowValue")?.unwrap_or(false);
                        self.series_mut(id, number)?.has_labels = flag;
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::ChartTitle(id) => {
                if lower == "text" || lower == "caption" {
                    let chart = self.chart_mut(id)?;
                    chart.title = text_of_value(&value);
                    chart.has_title = !chart.title.is_empty();
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::Legend(id) => {
                if lower == "position" {
                    self.chart_mut(id)?.legend_position = any_whole_number(&value).unwrap_or(-4152);
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(matches!(lower.as_str(), "includeinlayout" | "font"))
            }
            DrawingPart::Axis(id, which) => {
                let at = usize::from(which == 2);
                let taken = {
                    let chart = self.chart_mut(id)?;
                    let axis = &mut chart.axes[at];
                    match lower.as_str() {
                        "hastitle" => {
                            axis.has_title = style_face_boolean(&value, "Axis.HasTitle")?.unwrap_or(false);
                            if axis.has_title && axis.title.is_empty() {
                                axis.title = "軸ラベル".to_string();
                            }
                            true
                        }
                        "minimumscale" => { axis.min = any_number(&value); true }
                        "maximumscale" => { axis.max = any_number(&value); true }
                        "minimumscaleisauto" => {
                            if style_face_boolean(&value, "Axis.MinimumScaleIsAuto")?.unwrap_or(false) {
                                axis.min = None;
                            }
                            true
                        }
                        "maximumscaleisauto" => {
                            if style_face_boolean(&value, "Axis.MaximumScaleIsAuto")?.unwrap_or(false) {
                                axis.max = None;
                            }
                            true
                        }
                        "majorunit" | "minorunit" | "hasmajorgridlines" | "hasminorgridlines" | "tickLabelPosition"
                        | "ticklabelposition" | "reverseplotorder" | "crosses" | "crossesat" | "majorunitisauto" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::AxisTitle(id, which) => {
                if lower == "text" || lower == "caption" {
                    let chart = self.chart_mut(id)?;
                    chart.axes[usize::from(which == 2)].title = text_of_value(&value);
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::ChartArea(_) | DrawingPart::PlotArea(_) => {
                Ok(matches!(lower.as_str(), "rgb" | "color" | "colorindex" | "visible" | "width" | "height" | "left" | "top" | "insideleft" | "insidetop" | "insidewidth" | "insideheight"))
            }
            DrawingPart::Adjustments(id) => {
                if lower == "item" {
                    return Ok(false);
                }
                let _ = id;
                Ok(false)
            }
            DrawingPart::Paragraphs(_) => Ok(false),
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                let mut any = false;
                for id in ids {
                    if self.drawing_set(DrawingPart::Shape(id), name, value.clone())? {
                        any = true;
                    }
                }
                Ok(any)
            }
            DrawingPart::OleObjects(_) | DrawingPart::OleObjectsSnapshot(..)
                if matches!(lower.as_str(), "visible" | "left" | "top" | "width" | "height" | "enabled" | "locked" | "placement" | "printobject") => {
                let ids = self.ole_objects_for_part(part);
                if matches!(lower.as_str(), "left" | "top") {
                    let Some(anchor) = ids.last().copied() else { return Ok(false); };
                    let Some(requested) = any_number(&value) else {
                        return self.drawing_set(DrawingPart::OleObject(anchor), name, value);
                    };
                    let requested = (requested / 0.75).round() * 0.75;
                    let anchor_shape = self.shape(anchor)?;
                    let current = if lower == "left" { anchor_shape.left } else { anchor_shape.top };
                    let delta = requested - current;
                    for id in ids {
                        let shape = self.shape(id)?;
                        let position = if lower == "left" { shape.left } else { shape.top };
                        let target = ((position + delta) / 0.75).round() * 0.75;
                        self.drawing_set(DrawingPart::OleObject(id), name, Value::Double(target))?;
                    }
                    return Ok(true);
                }
                let mut any = false;
                for id in ids {
                    any |= self.drawing_set(DrawingPart::OleObject(id), name, value.clone())?;
                }
                Ok(any)
            }
            DrawingPart::OleObjects(_) | DrawingPart::OleObjectsSnapshot(..)
                if matches!(lower.as_str(), "shadow" | "autoload" | "onaction" | "sourcename") =>
            {
                Err(host_error_described(
                    1004,
                    format!("OLEObjects \u{30af}\u{30e9}\u{30b9}\u{306e} {name} \u{30d7}\u{30ed}\u{30d1}\u{30c6}\u{30a3}\u{3092}\u{8a2d}\u{5b9a}\u{3067}\u{304d}\u{307e}\u{305b}\u{3093}\u{3002}"),
                ))
            }
            DrawingPart::GroupItems(_) => Ok(false),
            DrawingPart::OleFormat(_) | DrawingPart::Shapes(_) | DrawingPart::OleObjects(_) | DrawingPart::OleObjectsSnapshot(..) | DrawingPart::ChartObjects(_) | DrawingPart::SeriesCollection(_) | DrawingPart::Points(..) | DrawingPart::Point(..) | DrawingPart::Axes(_) => Ok(false),
        }
    }

    // ---- calling ------------------------------------------------------------

    pub(super) fn drawing_call(&mut self, part: DrawingPart, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Shapes(sheet) => match lower.as_str() {
                "item" | "_default" => {
                    let listed = self.shapes_on(sheet);
                    let index = args.first().ok_or_else(|| "Shapes.Item takes an index or a name".to_string())?;
                    let id = self.pick_shape(&listed, index)?;
                    Ok(Some(self.part(DrawingPart::Shape(id))))
                }
                "addshape" => self.add_shape(sheet, args).map(Some),
                "addtextbox" => self.add_textbox(sheet, args).map(Some),
                "addlabel" => self.add_textbox(sheet, args).map(Some),
                "addline" => self.add_line(sheet, args).map(Some),
                // Measured: AddConnector names its shape by the connector's
                // type -- "Straight Arrow Connector", "Elbow Connector",
                // "Curved Connector" -- and it is an AutoShape (Type 1), where
                // AddLine's is "Straight Connector" and a line (9). Both
                // answer Connector True.
                "addconnector" => {
                    let kind = args.first().and_then(any_whole_number).unwrap_or(1);
                    let base = match kind {
                        1 => "Straight Arrow Connector",
                        2 => "Elbow Connector",
                        3 => "Curved Connector",
                        _ => return Err(host_error(5, "that is not a connector type")),
                    };
                    let rest: Vec<Value> = args.iter().skip(1).cloned().collect();
                    let shape = self.add_line(sheet, &rest)?;
                    if let Value::Object(object) = &shape {
                        if let Some(DrawingPart::Shape(id)) = self.drawing_part(object) {
                            let number = self.shape_mut(id)?.name.rsplit(' ').next().unwrap_or("1").to_string();
                            let record = self.shape_mut(id)?;
                            record.connector = kind;
                            record.name = format!("{base} {number}");
                        }
                    }
                    Ok(Some(shape))
                }
                "addchart2" | "addchart" => {
                    // AddChart2 Style, XlChartType, Left, Top, Width, Height;
                    // AddChart XlChartType, Left, Top, Width, Height.
                    let skip = usize::from(lower == "addchart2");
                    let chart_type = args.get(skip).and_then(any_whole_number).unwrap_or(51);
                    let number = |index: usize| args.get(skip + 1 + index).and_then(any_number).unwrap_or(0.0);
                    let (left, top, width, height) = (number(0), number(1), number(2).max(1.0), number(3).max(1.0));
                    let (width, height) = if args.get(skip + 3).is_none() { (360.0, 216.0) } else { (width, height) };
                    let id = self.add_chart(sheet, left, top, width, height, chart_type)?;
                    Ok(Some(self.part(DrawingPart::Shape(id))))
                }
                "addpicture" => Err(host_error(1004, "the browser cannot read a picture from a path")),
                "range" => {
                    let listed = self.shapes_on(sheet);
                    let names: Vec<Value> = match args.first() {
                        Some(Value::Array(listed)) => listed.values.clone(),
                        Some(one) => vec![one.clone()],
                        None => return Err(host_error(1004, "Shapes.Range takes a name or an array of names")),
                    };
                    let mut ids = Vec::new();
                    for name in &names {
                        ids.push(self.pick_shape(&listed, name)?);
                    }
                    Ok(Some(self.shape_range_object(ids)))
                }
                "selectall" => {
                    self.shape_selection = self.shapes_on(sheet);
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.shapes_on(sheet).len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::OleObjects(sheet) | DrawingPart::OleObjectsSnapshot(sheet, _) => match lower.as_str() {
                "add" => Err(host_error(1004, "OLEObjects.Add requires the active file system")),
                "item" | "_default" => {
                    let listed = self.ole_objects_for_part(part);
                    let index = args.first().ok_or_else(|| oxivba_core::host_error_described(
                        1004,
                        "OLEObjects クラスの Item プロパティを取得できません。",
                    ))?;
                    let id = self.pick_ole_object(&listed, index)?;
                    Ok(Some(self.part(DrawingPart::OleObject(id))))
                }
                "count" => Ok(Some(Value::Integer(self.ole_objects_for_part(part).len() as i64))),
                "duplicate" => {
                    let originals = self.ole_objects_for_part(part);
                    for id in originals.iter().copied() {
                        self.duplicate_shape(id, None)?;
                    }
                    Ok(Some(self.ole_objects_snapshot(sheet, originals)))
                }
                "group" => {
                    let id = self.group_shapes(sheet, self.ole_objects_for_part(part))?;
                    Ok(Some(self.part(DrawingPart::GroupObject(id))))
                }
                "delete" => {
                    for id in self.ole_objects_for_part(part) {
                        self.delete_shape(id)?;
                    }
                    Ok(Some(Value::Empty))
                }
                "select" => {
                    self.shape_selection = self.ole_objects_for_part(part);
                    Ok(Some(Value::Empty))
                }
                "bringtofront" | "sendtoback" => {
                    let ids = self.ole_objects_for_part(part);
                    if lower == "bringtofront" {
                        for id in ids {
                            self.move_shape_in_z_order(id, true)?;
                        }
                    } else {
                        for id in ids.into_iter().rev() {
                            self.move_shape_in_z_order(id, false)?;
                        }
                    }
                    Ok(Some(Value::Boolean(true)))
                }
                _ => Ok(None),
            },
            DrawingPart::ChartObjects(sheet) => match lower.as_str() {
                "item" | "_default" => {
                    let listed = self.charts_on(sheet);
                    let index = args.first().ok_or_else(|| "ChartObjects.Item takes an index or a name".to_string())?;
                    let id = self.pick_shape(&listed, index)?;
                    Ok(Some(self.part(DrawingPart::ChartObject(id))))
                }
                "add" => {
                    let (left, top, width, height) = Self::placed(args, 0, "ChartObjects.Add")?;
                    let id = self.add_chart(sheet, left, top, width, height, 51)?;
                    // Measured: a chart made by ChartObjects.Add has
                    // ChartStyle 2, where AddChart2's default is 201.
                    if let ShapeKind::Chart(chart) = &mut self.shape_mut(id)?.kind {
                        chart.style = 2;
                    }
                    Ok(Some(self.part(DrawingPart::ChartObject(id))))
                }
                "delete" => {
                    for id in self.charts_on(sheet) {
                        self.delete_shape(id)?;
                    }
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.charts_on(sheet).len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::OleObject(id) => match lower.as_str() {
                "select" => {
                    self.shape_selection = vec![id];
                    Ok(Some(Value::Null))
                }
                "copy" => {
                    self.shape_clipboard = Some(id);
                    self.clipboard = None;
                    self.pending_cut = None;
                    Ok(Some(Value::Boolean(true)))
                }
                "cut" => {
                    self.shape_clipboard = Some(id);
                    Ok(Some(Value::Boolean(true)))
                }
                "copypicture" => {
                    let bad_picture = || oxivba_core::host_error_described(
                        1004,
                        "OLEObject クラスの CopyPicture プロパティを取得できません。",
                    );
                    if args.len() > 2 {
                        return Err(oxivba_core::host_error(450, "Wrong number of arguments or invalid property assignment"));
                    }
                    let appearance_value = args.first().cloned().filter(|value| !matches!(value, Value::Missing));
                    let appearance_number = match appearance_value.as_ref() {
                        None => 1.0,
                        Some(Value::Boolean(value)) => if *value { 1.0 } else { 0.0 },
                        Some(value) => numeric_text(value).ok_or_else(bad_picture)?,
                    };
                    let appearance = appearance_number.round_ties_even() as i64;
                    let format = match args.get(1) {
                        None | Some(Value::Missing) => -4147,
                        Some(value) => numeric_text(value).ok_or_else(bad_picture)?.round_ties_even() as i64,
                    };
                    let boolean_success = appearance_number > 2.5 && appearance_number < 3.0;
                    if (!matches!(appearance, 1 | 2) && !boolean_success)
                        || !matches!(format, -4147 | 2)
                        || (appearance == 2 && format == 2)
                    {
                        return Err(bad_picture());
                    }
                    self.shape_clipboard = Some(id);
                    self.clipboard = None;
                    self.pending_cut = None;
                    let result = match appearance_value {
                        Some(_) if boolean_success => Value::Boolean(true),
                        Some(value) => value,
                        None => Value::Empty,
                    };
                    Ok(Some(result))
                }
                "delete" => {
                    self.delete_shape(id)?;
                    Ok(Some(Value::Empty))
                }
                "duplicate" => {
                    let duplicate = self.duplicate_shape(id, None)?;
                    Ok(Some(self.part(DrawingPart::OleObject(duplicate))))
                }
                "bringtofront" | "sendtoback" => {
                    self.move_shape_in_z_order(id, lower == "bringtofront")?;
                    Ok(Some(Value::Null))
                }
                "zorder" => Err(oxivba_core::host_error_described(1004, "OLEObject クラスの ZOrder メソッドが失敗しました。")),
                "update" => Err(oxivba_core::host_error_described(1004, "OLEObject クラスの Update メソッドが失敗しました。")),
                "activate" | "verb" => {
                    let prog_id = self.shape(id)?.original.as_ref().and_then(|drawing| drawing.ole_prog_id.as_deref()).unwrap_or_default();
                    if prog_id == "Package" {
                        Err(oxivba_core::host_error_described(1004, "このオブジェクトの作成元アプリケーションを起動できません。"))
                    } else {
                        Ok(Some(Value::Null))
                    }
                }
                _ => Ok(None),
            },
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) | DrawingPart::GroupObject(id) => match lower.as_str() {
                "groupitems" => {
                    let children = self.shape(id)?.group_children.clone();
                    if children.is_empty() {
                        Err(host_error(1004, "the shape is in no group"))
                    } else if let Some(index) = args.first() {
                        let child = self.pick_group_item(&children, index)?;
                        Ok(Some(self.part(DrawingPart::Shape(child))))
                    } else {
                        Ok(Some(self.part(DrawingPart::GroupItems(id))))
                    }
                }
                "parentgroup" => match self.shape(id)?.group_parent_id {
                    Some(parent) => Ok(Some(self.part(DrawingPart::Shape(parent)))),
                    None => Err(host_error(1004, "the shape is in no group")),
                },
                "delete" => {
                    self.delete_shape(id)?;
                    Ok(Some(Value::Empty))
                }
                "duplicate" => {
                    let new_id = self.duplicate_shape(id, None)?;
                    Ok(Some(self.part(DrawingPart::Shape(new_id))))
                }
                "copy" => {
                    self.shape_clipboard = Some(id);
                    self.clipboard = None;
                    self.pending_cut = None;
                    Ok(Some(Value::Boolean(true)))
                }
                "cut" => {
                    self.shape_clipboard = Some(id);
                    Ok(Some(Value::Boolean(true)))
                }
                "select" | "activate" => {
                    self.shape_selection = vec![id];
                    Ok(Some(Value::Boolean(true)))
                }
                "zorder" => Ok(Some(Value::Empty)),
                "incrementleft" | "incrementtop" | "incrementrotation" => {
                    let by = args.first().and_then(any_number).unwrap_or(0.0);
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "incrementleft" => shape.left = (shape.left + by).max(0.0),
                        "incrementtop" => shape.top = (shape.top + by).max(0.0),
                        _ => shape.rotation = (shape.rotation + by).rem_euclid(360.0),
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "scalewidth" | "scaleheight" => {
                    let factor = args.first().ok_or_else(|| {
                        host_error(450, "ScaleWidth/ScaleHeight requires Factor and RelativeToOriginalSize")
                    }).and_then(transparency_number)?;
                    let relative = args.get(1).ok_or_else(|| {
                        host_error(450, "ScaleWidth/ScaleHeight requires Factor and RelativeToOriginalSize")
                    }).and_then(|value| mso_asked(value, "RelativeToOriginalSize"))?;
                    let scale = match args.get(2) {
                        None | Some(Value::Missing) => 0,
                        Some(value) => excel_enum_variant(value, 0, 2, "Invalid shape scale origin")
                            .map_err(|_| oxivba_core::host_error_described(-2_147_024_809, "指定された値は境界を超えています。"))?,
                    };
                    let shape = self.shape_mut(id)?;
                    let original_size = if relative {
                        match (&shape.kind, shape.original.as_ref().map(|drawing| &drawing.kind)) {
                            (ShapeKind::Picture, Some(DrawingKind::Picture { bytes, format })) => picture_source_size_points(bytes, format)
                                .ok_or_else(|| host_error(1004, "the picture's original size is unavailable"))?,
                            (ShapeKind::Picture, _) => return Err(host_error(1004, "the picture's original size is unavailable")),
                            _ => return Err(oxivba_core::host_error_described(
                                -2_147_024_809,
                                "RelativeToOriginalSize 引数は、ピクチャまたは OLE オブジェクトにしか適用できません。",
                            )),
                        }
                    } else {
                        (shape.width, shape.height)
                    };
                    if lower == "scalewidth" {
                        let old = shape.width;
                        let new = (original_size.0 * factor).max(0.0);
                        shape.left = (shape.left + match scale {
                            0 => 0.0,
                            1 => (old - new) / 2.0,
                            _ => old - new,
                        }).max(0.0);
                        shape.width = new;
                    } else {
                        let old = shape.height;
                        let new = (original_size.1 * factor).max(0.0);
                        shape.top = (shape.top + match scale {
                            0 => 0.0,
                            1 => (old - new) / 2.0,
                            _ => old - new,
                        }).max(0.0);
                        shape.height = new;
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "bringtofront" | "sendtoback" => Ok(Some(Value::Empty)),
                "adjustments" => {
                    if args.is_empty() || matches!(args, [Value::Missing]) {
                        return Ok(Some(self.part(DrawingPart::Adjustments(id))));
                    }
                    self.drawing_call(DrawingPart::Adjustments(id), "item", args)
                }
                "textframe" => Ok(Some(self.part(DrawingPart::TextFrame(id)))),
                _ => Ok(None),
            },
            DrawingPart::Characters(id, start, length) if matches!(lower.as_str(), "delete" | "insert") => {
                if lower == "delete" && !args.is_empty() { return Err(host_error(450, "Characters.Delete takes no arguments")); }
                if lower == "insert" && (args.len() != 1 || matches!(args[0], Value::Missing)) {
                    return Err(host_error(450, "Characters.Insert requires one argument"));
                }
                let value = if lower == "insert" { args[0].clone() } else { Value::String(String::new()) };
                let comment = self.shape(id)?.comment_at;
                self.replace_shape_characters(id, start, length, &value)?;
                let count = comment.and_then(|at| self.note_at(at).map(|index| self.notes[index].units().len())).unwrap_or_else(|| self.shape(id).map(|shape| shape.text().encode_utf16().count()).unwrap_or(0)) as i64;
                let text = text_of_value(&value);
                Ok(Some(if comment.is_some() { Value::Integer(count) } else if lower == "insert" { Value::String(text) } else { Value::Empty }))
            }
            DrawingPart::Characters(..) if matches!(lower.as_str(), "clear" | "insertbefore" | "insertafter") => Ok(None),
            DrawingPart::TextFrame(id) | DrawingPart::TextRange(id) | DrawingPart::Characters(id, ..) => match lower.as_str() {
                "characters" => {
                    let start = args.first().and_then(any_whole_number).unwrap_or(1).max(1) as u32;
                    let length = args.get(1).and_then(any_whole_number).map(|n| n.max(0) as u32);
                    Ok(Some(self.part(DrawingPart::Characters(id, start, length))))
                }
                "delete" | "clear" => {
                    self.shape_mut(id)?.set_text("");
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "insertafter" | "insertbefore" => {
                    let added = args.first().map(text_of_value).unwrap_or_default();
                    let shape = self.shape_mut(id)?;
                    let text = if lower == "insertafter" { format!("{}{added}", shape.text()) } else { format!("{added}{}", shape.text()) };
                    shape.set_text(&text);
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                _ => Ok(None),
            },
            DrawingPart::Chart(id) => match lower.as_str() {
                "setsourcedata" => {
                    let Some(range) = args.first().and_then(|value| self.cells_of_range_object(value)) else {
                        return Err(host_error(1004, "SetSourceData takes a Range"));
                    };
                    // xlRows 1, xlColumns 2.
                    let by_rows = args.get(1).and_then(any_whole_number) == Some(1);
                    self.set_source_data(id, range, by_rows)?;
                    Ok(Some(Value::Empty))
                }
                "axes" => {
                    let which = args.first().and_then(any_whole_number).unwrap_or(1);
                    if !matches!(which, 1 | 2) {
                        return Err(host_error(1004, "Axes takes xlCategory or xlValue"));
                    }
                    Ok(Some(self.part(DrawingPart::Axis(id, which))))
                }
                "seriescollection" => match args.first() {
                    None | Some(Value::Missing) => Ok(Some(self.part(DrawingPart::SeriesCollection(id)))),
                    Some(index) => {
                        let number = self.series_number(id, index)?;
                        Ok(Some(self.part(DrawingPart::Series(id, number))))
                    }
                },
                "hasaxis" => Ok(Some(Value::Boolean(true))),
                // Measured: ApplyLayout 1 gives the chart a title.
                "applylayout" => {
                    let chart = self.chart_mut(id)?;
                    chart.has_title = true;
                    if chart.title.is_empty() {
                        chart.title = "グラフ タイトル".to_string();
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "export" => Err(host_error(1004, "the browser cannot write a file")),
                "delete" => {
                    self.delete_shape(id)?;
                    Ok(Some(Value::Empty))
                }
                "refresh" | "activate" | "select" | "deselect" | "setelement" | "applychartTemplate" | "applycharttemplate" | "clearToMatchStyle" | "cleartomatchstyle" => Ok(Some(Value::Empty)),
                "location" => Ok(Some(self.part(DrawingPart::Chart(id)))),
                _ => Ok(None),
            },
            DrawingPart::SeriesCollection(id) => match lower.as_str() {
                "item" | "_default" => {
                    let index = args.first().ok_or_else(|| "SeriesCollection takes an index".to_string())?;
                    let number = self.series_number(id, index)?;
                    Ok(Some(self.part(DrawingPart::Series(id, number))))
                }
                "newseries" => {
                    let chart = self.chart_mut(id)?;
                    let number = chart.series.len() + 1;
                    chart.series.push(SeriesRecord {
                        name: format!("系列{number}"),
                        name_ref: None,
                        values_ref: None,
                        x_ref: None,
                        values: Vec::new(),
                        xs: Vec::new(),
                        color: None,
                        chart_type: None,
                        has_labels: false,
                    });
                    self.redraw(id)?;
                    Ok(Some(self.part(DrawingPart::Series(id, number))))
                }
                "add" => {
                    if let Some(range) = args.first().and_then(|value| self.cells_of_range_object(value)) {
                        let numbers = self.range_numbers(range);
                        let reference = self.address_of(range);
                        let chart = self.chart_mut(id)?;
                        let number = chart.series.len() + 1;
                        chart.series.push(SeriesRecord {
                            name: format!("系列{number}"),
                            name_ref: None,
                            values_ref: Some(reference),
                            x_ref: None,
                            xs: (1..=numbers.len()).map(|n| n.to_string()).collect(),
                            values: numbers,
                            color: None,
                            chart_type: None,
                            has_labels: false,
                        });
                        self.redraw(id)?;
                    }
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.chart(id)?.series.len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::Series(id, number) => match lower.as_str() {
                "delete" => {
                    let chart = self.chart_mut(id)?;
                    if number >= 1 && number <= chart.series.len() {
                        chart.series.remove(number - 1);
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "points" => match args.first() {
                    None | Some(Value::Missing) => Ok(Some(self.part(DrawingPart::Points(id, number)))),
                    Some(wanted) => {
                        let count = self.series(id, number)?.values.len();
                        let at = any_whole_number(wanted).unwrap_or(0);
                        if at < 1 || at as usize > count {
                            return Err(host_error(1004, "there is no such point"));
                        }
                        Ok(Some(self.part(DrawingPart::Point(id, number, at as usize))))
                    }
                },
                "datalabels" => Ok(Some(self.part(DrawingPart::DataLabels(id, number)))),
                "applydatalabels" => {
                    self.series_mut(id, number)?.has_labels = true;
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "select" => Ok(Some(Value::Boolean(true))),
                // `Values(2)` and `XValues(2)` index the arrays they answer.
                "values" | "xvalues" => {
                    let whole = self.drawing_get(DrawingPart::Series(id, number), &lower)?;
                    match (args.first().and_then(any_whole_number), whole) {
                        (Some(index), Some(Value::Array(array))) => Ok(Some(
                            array
                                .values
                                .get(index.max(1) as usize - 1)
                                .cloned()
                                .ok_or_else(|| host_error(9, "no such point"))?,
                        )),
                        (_, whole) => Ok(whole),
                    }
                }
                _ => Ok(None),
            },
            DrawingPart::Adjustments(id) => match lower.as_str() {
                "item" | "_default" => {
                    let index = args.first().and_then(any_whole_number).unwrap_or(1);
                    let shape = self.shape(id)?;
                    shape
                        .adjusts
                        .get(index.max(1) as usize - 1)
                        .map(|value| Some(Value::Double(*value)))
                        .ok_or_else(|| host_error(1004, "there is no such adjustment"))
                }
                _ => Ok(None),
            },
            DrawingPart::Paragraphs(id) => match lower.as_str() {
                "item" | "_default" => Ok(Some(self.part(DrawingPart::TextRange(id)))),
                _ => Ok(None),
            },
            DrawingPart::Axes(id) => match lower.as_str() {
                "item" | "_default" => {
                    let which = args.first().and_then(any_whole_number).unwrap_or(1);
                    if !matches!(which, 1 | 2) {
                        return Err(host_error(1004, "Axes takes xlCategory or xlValue"));
                    }
                    Ok(Some(self.part(DrawingPart::Axis(id, which))))
                }
                _ => Ok(None),
            },
            DrawingPart::Points(id, number) => match lower.as_str() {
                "item" | "_default" | "count" => Ok(Some(Value::Integer(self.series(id, number)?.values.len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::ChartTitle(id) => match lower.as_str() {
                "characters" => Ok(Some(self.part(DrawingPart::ChartTitle(id)))),
                "select" | "delete" => Ok(Some(Value::Empty)),
                _ => Ok(None),
            },
            DrawingPart::AxisTitle(id, which) => match lower.as_str() {
                "characters" => Ok(Some(self.part(DrawingPart::AxisTitle(id, which)))),
                _ => Ok(None),
            },
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                match lower.as_str() {
                    "select" => {
                        self.shape_selection = ids;
                        Ok(Some(Value::Boolean(true)))
                    }
                    "item" | "_default" => {
                        let at = args.first().and_then(any_whole_number).unwrap_or(1).max(1) as usize;
                        ids.get(at - 1)
                            .map(|id| Ok(Some(self.part(DrawingPart::Shape(*id)))))
                            .unwrap_or_else(|| Err(host_error(1004, "there is no such shape")))
                    }
                    "count" => Ok(Some(Value::Integer(ids.len() as i64))),
                    "shaperange" => Ok(Some(self.part(part))),
                    "delete" => {
                        for id in ids {
                            self.delete_shape(id)?;
                        }
                        self.shape_selection.clear();
                        Ok(Some(Value::Empty))
                    }
                    "group" => {
                        let sheet = ids.first().and_then(|id| self.shape(*id).ok()).map(|shape| shape.sheet)
                            .ok_or_else(|| host_error(1004, "There are no shapes to group."))?;
                        let id = self.group_shapes(sheet, ids)?;
                        Ok(Some(self.part(DrawingPart::GroupObject(id))))
                    }
                    _ => match ids.first() {
                        Some(id) => self.drawing_call(DrawingPart::Shape(*id), name, args),
                        None => Ok(None),
                    },
                }
            }
            DrawingPart::GroupItems(id) => match lower.as_str() {
                "item" | "_default" => {
                    let children = self.shape(id)?.group_children.clone();
                    let child = args.first().ok_or_else(|| "GroupItems.Item takes an index or a name".to_string())?;
                    let child_id = self.pick_group_item(&children, child)?;
                    Ok(Some(self.part(DrawingPart::Shape(child_id))))
                }
                "count" => Ok(Some(Value::Integer(self.shape(id)?.group_children.len() as i64))),
                "range" => {
                    let children = self.shape(id)?.group_children.clone();
                    let args = args.first().cloned().into_iter().collect::<Vec<_>>();
                    let ids = if let Some(Value::Array(array)) = args.first() {
                        array.values.iter().map(|value| self.pick_group_item(&children, value)).collect::<Result<Vec<_>, _>>()?
                    } else if let Some(value) = args.first() {
                        vec![self.pick_group_item(&children, value)?]
                    } else {
                        return Err(host_error(1004, "GroupItems.Range takes an index or an array"));
                    };
                    Ok(Some(self.shape_range_object(ids)))
                }
                _ => Ok(None),
            },
            DrawingPart::PictureFormat(id) => match lower.as_str() {
                "incrementbrightness" | "incrementcontrast" => {
                    if args.len() != 1 {
                        return Err(host_error(450, "Wrong number of arguments or invalid property assignment"));
                    }
                    let delta = transparency_number(&args[0])? as f32;
                    if !delta.is_finite() {
                        return Err(host_error(1004, "Invalid picture adjustment"));
                    }
                    let (property, current) = {
                        let shape = self.shape(id)?;
                        let Some(oxicells_core::ir::Drawing { kind: DrawingKind::Picture { format, .. }, .. }) = shape.original.as_ref() else {
                            return Err(host_error(438, "Object doesn't support this property or method"));
                        };
                        let (property, raw) = if lower == "incrementbrightness" {
                            ("brightness", format.brightness)
                        } else {
                            ("contrast", format.contrast)
                        };
                        (property, 0.5 + f64::from(raw) / 200_000.0)
                    };
                    let next = (current as f32 + delta).clamp(0.0, 1.0);
                    self.drawing_set(DrawingPart::PictureFormat(id), property, Value::Single(next))?;
                    Ok(Some(Value::Empty))
                }
                _ => Ok(None),
            },
            DrawingPart::Legend(_) | DrawingPart::DataLabels(..) | DrawingPart::ChartArea(_) | DrawingPart::PlotArea(_) => match lower.as_str() {
                "select" | "delete" | "clearformats" => Ok(Some(Value::Empty)),
                _ => Ok(None),
            },
            _ => Ok(None),
        }
    }

    /// A series' `Values`, `XValues` or `Name` given a Range: the range is
    /// what is kept, and its cells are read now.
    pub(super) fn drawing_set_range(&mut self, part: DrawingPart, name: &str, range: CellRange) -> Result<Option<bool>, String> {
        let DrawingPart::Series(id, number) = part else {
            return Ok(None);
        };
        let reference = self.address_of(range);
        match name.to_ascii_lowercase().as_str() {
            "values" => {
                let numbers = self.range_numbers(range);
                let series = self.series_mut(id, number)?;
                series.values = numbers;
                series.values_ref = Some(reference);
                if series.xs.len() != series.values.len() {
                    series.xs = (1..=series.values.len()).map(|n| n.to_string()).collect();
                }
            }
            "xvalues" => {
                let texts = self.range_texts(range);
                let series = self.series_mut(id, number)?;
                series.xs = texts;
                series.x_ref = Some(reference);
            }
            "name" => {
                let text = self.range_texts(range).into_iter().next().unwrap_or_default();
                let series = self.series_mut(id, number)?;
                series.name = text;
                series.name_ref = Some(reference);
            }
            _ => return Ok(None),
        }
        self.redraw(id)?;
        Ok(Some(true))
    }

    fn series_number(&self, id: u64, index: &Value) -> Result<usize, String> {
        let chart = self.chart(id)?;
        match index {
            Value::String(name) => chart
                .series
                .iter()
                .position(|series| series.name.eq_ignore_ascii_case(name))
                .map(|at| at + 1)
                .ok_or_else(|| host_error(1004, format!("there is no series called {name}"))),
            value => match any_whole_number(value) {
                Some(number) if number >= 1 && (number as usize) <= chart.series.len() => Ok(number as usize),
                _ => Err(host_error(1004, "there is no such series")),
            },
        }
    }
}

/// The name Excel gives a preset shape, by `msoAutoShapeType`.
fn auto_shape_name(kind: i64) -> &'static str {
    auto_shape_label(kind)
}

fn chart_is_pie(kind: i64) -> bool {
    matches!(kind, 5 | 69 | 70 | 71 | -4102 | -4120 | 80)
}

/// Text for what a macro hands a text property that is not text.
fn text_of_value(value: &Value) -> String {
    match value {
        Value::String(text) => text.clone(),
        Value::Empty | Value::Null | Value::Missing | Value::Nothing => String::new(),
        Value::Boolean(flag) => if *flag { "True".to_string() } else { "False".to_string() },
        other => match any_number(other) {
            Some(number) if number.fract() == 0.0 => format!("{}", number as i64),
            Some(number) => number.to_string(),
            None => String::new(),
        },
    }
}

/// The theme table's index for accent1.
const ACCENT1_INDEX: usize = 4;

#[cfg(test)]
mod picture_extent_tests {
    use super::picture_original_size_points;

    fn assert_size(bytes: &[u8], expected: (f64, f64)) {
        let actual = picture_original_size_points(bytes).expect("recognized picture format");
        assert!((actual.0 - expected.0).abs() < 0.0001, "width {actual:?} != {expected:?}");
        assert!((actual.1 - expected.1).abs() < 0.0001, "height {actual:?} != {expected:?}");
    }

    #[test]
    fn bitmap_extent_uses_pixels_per_meter() {
        let mut bytes = vec![0u8; 54];
        bytes[..2].copy_from_slice(b"BM");
        bytes[14..18].copy_from_slice(&40u32.to_le_bytes());
        bytes[18..22].copy_from_slice(&120i32.to_le_bytes());
        bytes[22..26].copy_from_slice(&80i32.to_le_bytes());
        bytes[38..42].copy_from_slice(&3780i32.to_le_bytes());
        bytes[42..46].copy_from_slice(&3780i32.to_le_bytes());
        assert_size(&bytes, (120.0 * 72.0 / (3780.0 * 0.0254), 80.0 * 72.0 / (3780.0 * 0.0254)));
    }

    #[test]
    fn png_extent_uses_physical_pixel_density() {
        let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
        bytes.extend_from_slice(&13u32.to_be_bytes());
        bytes.extend_from_slice(b"IHDR");
        bytes.extend_from_slice(&120u32.to_be_bytes());
        bytes.extend_from_slice(&80u32.to_be_bytes());
        bytes.extend_from_slice(&[8, 2, 0, 0, 0]);
        bytes.extend_from_slice(&[0; 4]);
        bytes.extend_from_slice(&9u32.to_be_bytes());
        bytes.extend_from_slice(b"pHYs");
        bytes.extend_from_slice(&3779u32.to_be_bytes());
        bytes.extend_from_slice(&3779u32.to_be_bytes());
        bytes.push(1);
        bytes.extend_from_slice(&[0; 4]);
        assert_size(&bytes, (120.0 * 72.0 / (3779.0 * 0.0254), 80.0 * 72.0 / (3779.0 * 0.0254)));
    }

    #[test]
    fn jpeg_extent_uses_jfif_dpi() {
        let mut bytes = vec![0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10];
        bytes.extend_from_slice(b"JFIF\0\x01\x02\x01\x00\x60\x00\x60\0\0");
        bytes.extend_from_slice(&[0xff, 0xc0, 0x00, 0x0b, 8, 0, 80, 0, 120, 1, 1, 0x11, 0]);
        assert_size(&bytes, (90.0, 60.0));
    }

    #[test]
    fn jpeg_extent_uses_exif_density_when_jfif_has_no_density() {
        let book = oxicells_core::parse_xlsx(include_bytes!("../../../tests/fixtures/vba/scale_picture_original_exif_jpeg.xlsx")).unwrap();
        let bytes = book.sheets[0].drawings.iter().find_map(|drawing| match &drawing.kind {
            oxicells_core::ir::DrawingKind::Picture { bytes, .. } => Some(bytes.as_slice()),
            _ => None,
        }).expect("fixture picture bytes");
        assert_size(bytes, (60.0, 60.0));
    }

    #[test]
    fn gif_extent_uses_default_96_dpi() {
        let mut bytes = b"GIF89a".to_vec();
        bytes.extend_from_slice(&120u16.to_le_bytes());
        bytes.extend_from_slice(&80u16.to_le_bytes());
        assert_size(&bytes, (90.0, 60.0));
    }
}
