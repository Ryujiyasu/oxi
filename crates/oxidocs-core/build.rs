use std::path::Path;

fn main() {
    // S244 (2026-05-24): declare check-cfg for has_local_font_metrics to
    // silence "unexpected cfg condition name" warnings (4 sites in
    // src/font/mod.rs).
    println!("cargo::rustc-check-cfg=cfg(has_local_font_metrics)");
    let metrics_path = Path::new("src/font/data/font_metrics_compact.json");
    if metrics_path.exists() {
        println!("cargo:rustc-cfg=has_local_font_metrics");
    }
    // Rebuild when ANY font data file changes (all are embedded via include_str!)
    println!("cargo:rerun-if-changed=src/font/data/");
}
