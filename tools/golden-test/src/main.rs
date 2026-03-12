//! Golden Test Harness — parse government Office documents and report success rate.
//!
//! Usage:
//!     golden-test <documents-dir>
//!
//! Reads all .docx/.xlsx/.pptx files from the directory tree,
//! attempts to parse each with Oxi, and generates a report.

use serde::Serialize;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

#[derive(Debug, Serialize)]
struct TestResult {
    filename: String,
    format: String,
    size_bytes: u64,
    success: bool,
    error: Option<String>,
    parse_time_ms: u64,
    /// Number of pages/sheets/slides parsed
    element_count: usize,
}

#[derive(Debug, Serialize)]
struct Report {
    total: usize,
    success: usize,
    failure: usize,
    success_rate: f64,
    by_format: FormatBreakdown,
    total_time_secs: f64,
    failures: Vec<TestResult>,
}

#[derive(Debug, Serialize)]
struct FormatBreakdown {
    docx: FormatStats,
    xlsx: FormatStats,
    pptx: FormatStats,
}

#[derive(Debug, Serialize)]
struct FormatStats {
    total: usize,
    success: usize,
    success_rate: f64,
}

fn find_files(dir: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                files.extend(find_files(&path));
            } else if let Some(ext) = path.extension() {
                let ext = ext.to_string_lossy().to_lowercase();
                if ext == "docx" || ext == "xlsx" || ext == "pptx" {
                    files.push(path);
                }
            }
        }
    }
    files
}

fn test_file(path: &Path) -> TestResult {
    let filename = path.file_name().unwrap_or_default().to_string_lossy().to_string();
    let ext = path
        .extension()
        .unwrap_or_default()
        .to_string_lossy()
        .to_lowercase();
    let size_bytes = fs::metadata(path).map(|m| m.len()).unwrap_or(0);

    let data = match fs::read(path) {
        Ok(d) => d,
        Err(e) => {
            return TestResult {
                filename,
                format: ext,
                size_bytes,
                success: false,
                error: Some(format!("Read error: {e}")),
                parse_time_ms: 0,
                element_count: 0,
            };
        }
    };

    let start = Instant::now();

    let (success, error, element_count) = match ext.as_str() {
        "docx" => match oxidocs_core::parse_docx(&data) {
            Ok(doc) => (true, None, doc.pages.len()),
            Err(e) => (false, Some(format!("{e}")), 0),
        },
        "xlsx" => match oxicells_core::parse_xlsx(&data) {
            Ok(wb) => (true, None, wb.sheets.len()),
            Err(e) => (false, Some(format!("{e}")), 0),
        },
        "pptx" => match oxislides_core::parse_pptx(&data) {
            Ok(pres) => (true, None, pres.slides.len()),
            Err(e) => (false, Some(format!("{e}")), 0),
        },
        _ => (false, Some("Unknown format".to_string()), 0),
    };

    let elapsed = start.elapsed();

    TestResult {
        filename,
        format: ext,
        size_bytes,
        success,
        error,
        parse_time_ms: elapsed.as_millis() as u64,
        element_count,
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let dir = if args.len() > 1 {
        PathBuf::from(&args[1])
    } else {
        PathBuf::from("./documents")
    };

    if !dir.exists() {
        eprintln!("Directory not found: {}", dir.display());
        std::process::exit(1);
    }

    let files = find_files(&dir);
    let total = files.len();
    println!("=== Oxi Golden Test ===");
    println!("Documents: {total}");
    println!();

    let overall_start = Instant::now();
    let mut results: Vec<TestResult> = Vec::new();

    for (i, file) in files.iter().enumerate() {
        let result = test_file(file);
        let status = if result.success { "OK" } else { "FAIL" };
        let err_msg = result.error.as_deref().unwrap_or("");
        println!(
            "[{:>4}/{:>4}] {:4} {} {} ({} bytes, {}ms) {}",
            i + 1,
            total,
            status,
            result.format,
            result.filename,
            result.size_bytes,
            result.parse_time_ms,
            err_msg
        );
        results.push(result);
    }

    let total_time = overall_start.elapsed();

    // Compute statistics
    let success = results.iter().filter(|r| r.success).count();
    let failure = total - success;
    let success_rate = if total > 0 {
        (success as f64 / total as f64) * 100.0
    } else {
        0.0
    };

    let docx_total = results.iter().filter(|r| r.format == "docx").count();
    let xlsx_total = results.iter().filter(|r| r.format == "xlsx").count();
    let pptx_total = results.iter().filter(|r| r.format == "pptx").count();
    let docx_success = results.iter().filter(|r| r.format == "docx" && r.success).count();
    let xlsx_success = results.iter().filter(|r| r.format == "xlsx" && r.success).count();
    let pptx_success = results.iter().filter(|r| r.format == "pptx" && r.success).count();

    let failures: Vec<TestResult> = results.into_iter().filter(|r| !r.success).collect();

    let report = Report {
        total,
        success,
        failure,
        success_rate,
        by_format: FormatBreakdown {
            docx: FormatStats {
                total: docx_total,
                success: docx_success,
                success_rate: if docx_total == 0 { 0.0 } else { (docx_success as f64 / docx_total as f64) * 100.0 },
            },
            xlsx: FormatStats {
                total: xlsx_total,
                success: xlsx_success,
                success_rate: if xlsx_total == 0 { 0.0 } else { (xlsx_success as f64 / xlsx_total as f64) * 100.0 },
            },
            pptx: FormatStats {
                total: pptx_total,
                success: pptx_success,
                success_rate: if pptx_total == 0 { 0.0 } else { (pptx_success as f64 / pptx_total as f64) * 100.0 },
            },
        },
        total_time_secs: total_time.as_secs_f64(),
        failures,
    };

    println!();
    println!("══════════════════════════════════════════");
    println!("  Oxi Golden Test Report");
    println!("══════════════════════════════════════════");
    println!("  Total:        {}", report.total);
    println!("  Success:      {}", report.success);
    println!("  Failure:      {}", report.failure);
    println!(
        "  Success Rate: {:.1}%",
        report.success_rate
    );
    println!();
    println!(
        "  DOCX: {}/{} ({:.1}%)",
        report.by_format.docx.success,
        report.by_format.docx.total,
        report.by_format.docx.success_rate
    );
    println!(
        "  XLSX: {}/{} ({:.1}%)",
        report.by_format.xlsx.success,
        report.by_format.xlsx.total,
        report.by_format.xlsx.success_rate
    );
    println!(
        "  PPTX: {}/{} ({:.1}%)",
        report.by_format.pptx.success,
        report.by_format.pptx.total,
        report.by_format.pptx.success_rate
    );
    println!("  Time:         {:.1}s", report.total_time_secs);
    println!("══════════════════════════════════════════");

    if !report.failures.is_empty() {
        println!();
        println!("Failed files:");
        for f in &report.failures {
            println!(
                "  {} ({}) — {}",
                f.filename,
                f.format,
                f.error.as_deref().unwrap_or("unknown")
            );
        }
    }

    // Save JSON report
    let report_path = dir.join("golden_test_report.json");
    if let Ok(json) = serde_json::to_string_pretty(&report) {
        let _ = fs::write(&report_path, &json);
        println!();
        println!("Report saved: {}", report_path.display());
    }
}
