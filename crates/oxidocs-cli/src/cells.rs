// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

mod cells_server;

fn main() {
    let mut args = std::env::args().skip(1);
    match args.next().as_deref() {
        Some("serve") if args.next().is_none() => {
            // Windows executables have a small default main-thread stack.
            // Nested native VBA execution needs more room in debug builds too.
            let result = std::thread::Builder::new()
                .name("oxicells-server".into())
                .stack_size(16 * 1024 * 1024)
                .spawn(|| cells_server::serve(std::io::stdin().lock(), std::io::stdout().lock()))
                .map_err(|error| error.to_string())
                .and_then(|worker| {
                    worker
                        .join()
                        .unwrap_or_else(|panic| std::panic::resume_unwind(panic))
                });
            if let Err(error) = result {
                eprintln!("{error}");
                std::process::exit(2);
            }
        }
        Some("--version") => println!("oxicells {}", env!("CARGO_PKG_VERSION")),
        _ => {
            eprintln!("Usage: oxicells serve | --version");
            std::process::exit(2);
        }
    }
}
