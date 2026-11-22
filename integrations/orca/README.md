# Oxi Slides for Orca

Source of the [Orca](https://www.onorca.dev) plugin **Oxi Slides** (`ryujiyasu.oxi-slides`), a read-only `.pptx` viewer. Orca opens a `.pptx` in the plugin's editor tab, and the plugin shows the slides with Oxi's WebAssembly engine. Orca has no viewer of its own for `.pptx`. What users see is described in [`plugin/README.md`](plugin/README.md).

## How it is built

The plugin is `web/index.html` (the page the web editor and the VS Code extension use), patched by [`build.mjs`](build.mjs):

- **Engine:** `crates/oxidocs-wasm` built with `--features slides,external-font-catalog`, the PowerPoint bindings without the Excel, VBA, PDF and hanko code (about 41 MB instead of 97 MB). The page imports the module as a namespace, so the exports this build lacks are `undefined` instead of a load error.
- **Engine parts:** Orca's isolated frames read WebAssembly whole and refuse files over 16 MiB. The module is split into 15 MiB slices, and the page joins them before starting the engine. The community catalog judges files by their content: it refused gzipped parts as archives (0.1.0–0.1.1) and a bare slice without a WebAssembly header as an unknown binary (0.1.2), both as `binary-files`. So each `oxidocs_engine.N.wasm` is a valid WebAssembly module of its own (checked with `WebAssembly.validate` at build time) whose only section is a custom section, `oxi-engine-part`, holding one slice. It also refuses `.ico`, so the page's favicon links are dropped. The plugin is about 40 MiB, under Orca's 50 MiB limit.
- **Viewer:** the title bar and ribbon are hidden, the main slide is drawn read-only like the thumbnails, and a slide-show button is added to the status bar.
- **Orca bridge:** the page asks for the deck with `orca.document.session()` (`contentKind: "bytes"`) and redraws on `onReplace`, staying on the same slide. It never calls `change()`, so Orca never saves anything from the plugin.

## Build and try it

```sh
cd crates/oxidocs-wasm
wasm-pack build --target web --release --out-dir pkg-slides -- --features slides,external-font-catalog
cd ../..
node integrations/orca/build.mjs            # writes integrations/orca/dist
orca plugin check integrations/orca/dist
orca plugin dev load integrations/orca/dist # prints a fingerprint to approve
orca plugin shot integrations/orca/dist --editor path/to/deck.pptx
```

On Linux, `orca` here is Orca's CLI (`~/.config/orca/linux-orca-cli-shim/orca`), not the GNOME screen reader. Start Orca and its CLI without `ELECTRON_RUN_AS_NODE` in the environment: VS Code's terminals set it, and it makes Electron run as plain Node.

## Publishing

Community plugins are listed from their own repository. `dist/` is committed to [Ryujiyasu/orca-oxi-slides](https://github.com/Ryujiyasu/orca-oxi-slides), which has the `orca-plugin` topic. A `vX.Y.Z` tag there publishes that version on [onorca.dev/plugins](https://www.onorca.dev/plugins). Bump `version` in [`plugin/orca-plugin.json`](plugin/orca-plugin.json) for every release.

## Licence

The glue in this folder is MIT OR Apache-2.0. The plugin it builds contains the engine, which is MPL-2.0, so its manifest says `MPL-2.0` and it ships the MPL text as `LICENSE`.
