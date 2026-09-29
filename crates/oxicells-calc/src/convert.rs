// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! CONVERT: a number from one unit to another of the same kind, the units
//! and their factors as Excel lists them, a metric or binary prefix allowed
//! where Excel allows one.

use crate::value::{ExcelError, Value};

#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    Mass,
    Distance,
    Time,
    Pressure,
    Force,
    Energy,
    Power,
    Magnetism,
    Temperature,
    Volume,
    Area,
    Information,
    Speed,
}

/// (name, kind, size in the kind's base unit, takes a prefix)
const UNITS: &[(&str, Kind, f64, bool)] = &[
    // Mass, in kilograms.
    ("g", Kind::Mass, 1e-3, true),
    // A slug is a pound-force over a foot a second squared: measured,
    // 32.1740485564304 lbm.
    ("sg", Kind::Mass, 0.453_592_37 * 9.806_65 / 0.3048, false),
    ("lbm", Kind::Mass, 0.453_592_37, false),
    ("u", Kind::Mass, 1.660_538_782e-27, true),
    ("ozm", Kind::Mass, 0.028_349_523_125, false),
    ("grain", Kind::Mass, 6.479_891e-5, false),
    ("cwt", Kind::Mass, 45.359_237, false),
    ("shweight", Kind::Mass, 45.359_237, false),
    ("uk_cwt", Kind::Mass, 50.802_345_44, false),
    ("lcwt", Kind::Mass, 50.802_345_44, false),
    ("hweight", Kind::Mass, 50.802_345_44, false),
    ("stone", Kind::Mass, 6.350_293_18, false),
    ("ton", Kind::Mass, 907.184_74, false),
    ("uk_ton", Kind::Mass, 1_016.046_908_8, false),
    ("LTON", Kind::Mass, 1_016.046_908_8, false),
    ("brton", Kind::Mass, 1_016.046_908_8, false),
    // Distance, in metres.
    ("m", Kind::Distance, 1.0, true),
    ("mi", Kind::Distance, 1_609.344, false),
    ("Nmi", Kind::Distance, 1_852.0, false),
    ("in", Kind::Distance, 0.0254, false),
    ("ft", Kind::Distance, 0.3048, false),
    ("yd", Kind::Distance, 0.9144, false),
    ("ang", Kind::Distance, 1e-10, true),
    ("ell", Kind::Distance, 1.143, false),
    ("ly", Kind::Distance, 9.460_730_472_580_8e15, true),
    ("parsec", Kind::Distance, 3.085_677_581_281_55e16, true),
    ("pc", Kind::Distance, 3.085_677_581_281_55e16, true),
    ("Picapt", Kind::Distance, 0.0254 / 72.0, false),
    ("Pica", Kind::Distance, 0.0254 / 72.0, false),
    ("pica", Kind::Distance, 0.0254 / 6.0, false),
    ("survey_mi", Kind::Distance, 1_609.347_218_694_44, false),
    // Time, in seconds.
    ("yr", Kind::Time, 31_557_600.0, false),
    ("day", Kind::Time, 86_400.0, false),
    ("d", Kind::Time, 86_400.0, false),
    ("hr", Kind::Time, 3_600.0, false),
    ("mn", Kind::Time, 60.0, false),
    ("min", Kind::Time, 60.0, false),
    ("sec", Kind::Time, 1.0, true),
    ("s", Kind::Time, 1.0, true),
    // Pressure, in pascals.
    ("Pa", Kind::Pressure, 1.0, true),
    ("p", Kind::Pressure, 1.0, true),
    ("atm", Kind::Pressure, 101_325.0, true),
    ("at", Kind::Pressure, 101_325.0, true),
    ("mmHg", Kind::Pressure, 133.322, true),
    ("psi", Kind::Pressure, 6_894.757_293_168_36, false),
    ("Torr", Kind::Pressure, 133.322_368_421_053, false),
    // Force, in newtons.
    ("N", Kind::Force, 1.0, true),
    ("dyn", Kind::Force, 1e-5, true),
    ("dy", Kind::Force, 1e-5, true),
    ("lbf", Kind::Force, 4.448_221_615_260_5, false),
    ("pond", Kind::Force, 9.806_65e-3, true),
    // Energy, in joules.
    ("J", Kind::Energy, 1.0, true),
    ("e", Kind::Energy, 1e-7, true),
    ("c", Kind::Energy, 4.184, true),
    ("cal", Kind::Energy, 4.1868, true),
    ("eV", Kind::Energy, 1.602_176_487e-19, true),
    ("ev", Kind::Energy, 1.602_176_487e-19, true),
    // A horsepower for an hour: measured, 1 HPh is 745.69987158227 Wh.
    ("HPh", Kind::Energy, 745.699_871_582_27 * 3_600.0, false),
    ("hh", Kind::Energy, 745.699_871_582_27 * 3_600.0, false),
    ("Wh", Kind::Energy, 3_600.0, true),
    ("wh", Kind::Energy, 3_600.0, true),
    ("flb", Kind::Energy, 1.355_817_948_331_4, false),
    ("BTU", Kind::Energy, 1_055.055_852_62, false),
    ("btu", Kind::Energy, 1_055.055_852_62, false),
    // Power, in watts.
    ("HP", Kind::Power, 745.699_871_582_27, false),
    ("h", Kind::Power, 745.699_871_582_27, false),
    ("PS", Kind::Power, 735.498_75, false),
    ("W", Kind::Power, 1.0, true),
    ("w", Kind::Power, 1.0, true),
    // Magnetism, in teslas.
    ("T", Kind::Magnetism, 1.0, true),
    ("ga", Kind::Magnetism, 1e-4, true),
    // Volume, in litres, the unit Excel reckons them in: measured,
    // uk_pt to pt 1.20094992550486 comes out that way and not in cubic metres.
    ("tsp", Kind::Volume, 4.928_921_593_75e-3, false),
    ("tspm", Kind::Volume, 5e-3, false),
    ("tbs", Kind::Volume, 1.478_676_478_125e-2, false),
    ("oz", Kind::Volume, 2.957_352_956_25e-2, false),
    ("cup", Kind::Volume, 0.236_588_236_5, false),
    ("pt", Kind::Volume, 0.473_176_473, false),
    ("us_pt", Kind::Volume, 0.473_176_473, false),
    ("uk_pt", Kind::Volume, 0.568_261_25, false),
    ("qt", Kind::Volume, 0.946_352_946, false),
    ("uk_qt", Kind::Volume, 1.136_522_5, false),
    ("gal", Kind::Volume, 3.785_411_784, false),
    ("uk_gal", Kind::Volume, 4.546_09, false),
    ("l", Kind::Volume, 1.0, true),
    ("L", Kind::Volume, 1.0, true),
    ("lt", Kind::Volume, 1.0, true),
    ("ang3", Kind::Volume, 1e-27, true),
    ("ang^3", Kind::Volume, 1e-27, true),
    ("barrel", Kind::Volume, 158.987_294_928, false),
    ("bushel", Kind::Volume, 35.239_070_166_88, false),
    ("ft3", Kind::Volume, 28.316_846_592, false),
    ("ft^3", Kind::Volume, 28.316_846_592, false),
    ("in3", Kind::Volume, 0.016_387_064, false),
    ("in^3", Kind::Volume, 0.016_387_064, false),
    ("m3", Kind::Volume, 1_000.0, true),
    ("m^3", Kind::Volume, 1_000.0, true),
    ("mi3", Kind::Volume, 4_168_181_825_440.58, false),
    ("mi^3", Kind::Volume, 4_168_181_825_440.58, false),
    ("yd3", Kind::Volume, 764.554_857_984, false),
    ("yd^3", Kind::Volume, 764.554_857_984, false),
    ("Nmi3", Kind::Volume, 6_352_182_208_000.0, false),
    ("Nmi^3", Kind::Volume, 6_352_182_208_000.0, false),
    ("GRT", Kind::Volume, 2_831.684_659_2, false),
    ("regton", Kind::Volume, 2_831.684_659_2, false),
    ("MTON", Kind::Volume, 1_132.673_863_68, false),
    // Area, in square metres.
    ("uk_acre", Kind::Area, 4_046.856_422_4, false),
    ("us_acre", Kind::Area, 4_046.872_609_874_25, false),
    ("ang2", Kind::Area, 1e-20, true),
    ("ang^2", Kind::Area, 1e-20, true),
    ("ar", Kind::Area, 100.0, true),
    ("ft2", Kind::Area, 0.092_903_04, false),
    ("ft^2", Kind::Area, 0.092_903_04, false),
    ("ha", Kind::Area, 10_000.0, false),
    ("in2", Kind::Area, 6.4516e-4, false),
    ("in^2", Kind::Area, 6.4516e-4, false),
    ("m2", Kind::Area, 1.0, true),
    ("m^2", Kind::Area, 1.0, true),
    ("Morgen", Kind::Area, 2_500.0, false),
    ("mi2", Kind::Area, 2_589_988.110_336, false),
    ("mi^2", Kind::Area, 2_589_988.110_336, false),
    ("Nmi2", Kind::Area, 3_429_904.0, false),
    ("Nmi^2", Kind::Area, 3_429_904.0, false),
    ("yd2", Kind::Area, 0.836_127_36, false),
    ("yd^2", Kind::Area, 0.836_127_36, false),
    // Information, in bits.
    ("bit", Kind::Information, 1.0, true),
    ("byte", Kind::Information, 8.0, true),
    // Speed, in metres a second.
    ("admkn", Kind::Speed, 0.514_773_333_333_333, false),
    ("kn", Kind::Speed, 1_852.0 / 3_600.0, false),
    ("m/h", Kind::Speed, 1.0 / 3_600.0, true),
    ("m/hr", Kind::Speed, 1.0 / 3_600.0, true),
    ("m/s", Kind::Speed, 1.0, true),
    ("m/sec", Kind::Speed, 1.0, true),
    ("mph", Kind::Speed, 0.447_04, false),
];

const PREFIXES: &[(&str, f64)] = &[
    ("Y", 1e24), ("Z", 1e21), ("E", 1e18), ("P", 1e15), ("T", 1e12), ("G", 1e9), ("M", 1e6), ("k", 1e3),
    ("h", 1e2), ("da", 1e1), ("e", 1e1), ("d", 1e-1), ("c", 1e-2), ("m", 1e-3), ("u", 1e-6), ("n", 1e-9),
    ("p", 1e-12), ("f", 1e-15), ("a", 1e-18), ("z", 1e-21), ("y", 1e-24),
];

const BINARY: &[(&str, f64)] = &[
    ("ki", 1024.0), ("Mi", 1_048_576.0), ("Gi", 1_073_741_824.0), ("Ti", 1_099_511_627_776.0),
    ("Pi", 1_125_899_906_842_624.0), ("Ei", 1_152_921_504_606_846_976.0), ("Zi", 1.180_591_620_717_411_3e21),
    ("Yi", 1.208_925_819_614_629_2e24),
];

const TEMPERATURES: &[&str] = &["C", "cel", "F", "fah", "K", "kel", "Rank", "Reau"];

/// A unit's kind and size, a prefix counted in; the unit exactly as
/// written first, since "mi" is a mile and not a milli-inch.
fn unit(name: &str) -> Option<(Kind, f64)> {
    if let Some((_, kind, size, _)) = UNITS.iter().find(|(held, ..)| *held == name) {
        return Some((*kind, *size));
    }
    // Squared and cubed metric units take the prefix squared or cubed.
    for (prefix, scale) in BINARY {
        if let Some(rest) = name.strip_prefix(prefix) {
            if let Some((_, kind, size, true)) = UNITS.iter().find(|(held, ..)| *held == rest) {
                if *kind == Kind::Information {
                    return Some((*kind, size * scale));
                }
            }
        }
    }
    for (prefix, scale) in PREFIXES {
        if let Some(rest) = name.strip_prefix(prefix) {
            if let Some((_, kind, size, true)) = UNITS.iter().find(|(held, ..)| *held == rest) {
                let power = if rest.ends_with('2') { 2 } else if rest.ends_with('3') { 3 } else { 1 };
                return Some((*kind, size * scale.powi(power)));
            }
        }
    }
    None
}

/// A temperature in kelvin, and back.
fn kelvin(name: &str, value: f64, into: bool) -> Option<f64> {
    let (scale, offset) = match name {
        "C" | "cel" => (1.0, 273.15),
        "F" | "fah" => (5.0 / 9.0, 459.67 * 5.0 / 9.0),
        "K" | "kel" => (1.0, 0.0),
        "Rank" => (5.0 / 9.0, 0.0),
        "Reau" => (1.25, 273.15),
        _ => return None,
    };
    Some(if into { value * scale + offset } else { (value - offset) / scale })
}

/// Temperatures may take a metric prefix too ("mK"); the plain name wins.
fn temperature(name: &str) -> Option<(&str, f64)> {
    if TEMPERATURES.contains(&name) {
        return Some((name, 1.0));
    }
    for (prefix, scale) in PREFIXES {
        if let Some(rest) = name.strip_prefix(prefix) {
            if matches!(rest, "K" | "kel") {
                return Some((rest, *scale));
            }
        }
    }
    None
}

pub(crate) fn convert(value: f64, from: &str, to: &str) -> Result<Value, ExcelError> {
    if let (Some((from_name, from_scale)), Some((to_name, to_scale))) = (temperature(from), temperature(to)) {
        let k = kelvin(from_name, value * from_scale, true).ok_or(ExcelError::NA)?;
        let out = kelvin(to_name, k, false).ok_or(ExcelError::NA)? / to_scale;
        return Ok(Value::Number(out));
    }
    let (Some((from_kind, from_size)), Some((to_kind, to_size))) = (unit(from), unit(to)) else {
        return Err(ExcelError::NA);
    };
    if from_kind != to_kind {
        return Err(ExcelError::NA);
    }
    Ok(Value::Number(value * from_size / to_size))
}
