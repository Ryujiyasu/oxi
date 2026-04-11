use std::path::Path;

fn main() {
    let metrics_path = Path::new("src/font/data/font_metrics_compact.json");
    if metrics_path.exists() {
        println!("cargo:rustc-cfg=has_local_font_metrics");
    }
    // Rebuild when ANY font data file changes (all are embedded via include_str!)
    println!("cargo:rerun-if-changed=src/font/data/");
}
