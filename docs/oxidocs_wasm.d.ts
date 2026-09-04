/* tslint:disable */
/* eslint-disable */

/**
 * Where the engine breaks one paragraph's lines, for a box `width_pt` wide.
 *
 * The browser has been wrapping slide text with its own text layout, so what
 * a person edits on screen is not what the renderer -- or PowerPoint --
 * produces. This is the engine's own break test
 * (`pptx-master-unit-break-law`: each glyph's advance rounded to 1/8 pt, then
 * summed), running where there is no font system.
 *
 * Returns `null` when the metrics tables cannot measure the text, which is
 * most families: the tables carry the faces that were measured from their
 * real files, and a guess here would break lines PowerPoint keeps. A caller
 * that gets null must say the wrap is the browser's, not the engine's.
 *
 * `runs` is the paragraph's runs, so a line carrying a bold word is measured
 * per run rather than wholly in one weight.
 */
export function break_slide_paragraph(text: string, font_size: number, family: string, bold: boolean, italic: boolean, width_pt: number, runs: any, advances: any): any;

/**
 * Build a .docx from a content structure.
 * `content`: JS array of block objects:
 *   { type: "paragraph", runs: [{text, bold?, italic?, underline?, strikethrough?, font_family?, font_size?, color?}], alignment?, heading_level?, line_height? }
 *   { type: "table", rows: [[{text, bold?}]] }
 */
export function build_docx(content: any): Uint8Array;

/**
 * Build a .docx from content, using a template docx for styles/theme/numbering.
 * Preserves original formatting while replacing document content.
 */
export function build_docx_with_template(content: any, template: Uint8Array): Uint8Array;

/**
 * Create a blank .docx file and return it as bytes.
 * Can be used to create a new document from scratch.
 */
export function create_blank_docx(): Uint8Array;

/**
 * Generate a PDF from scratch with the given text content.
 * Returns the PDF bytes.
 */
export function create_pdf(title: string, text: string): Uint8Array;

/**
 * Convert a .docx file to PDF bytes.
 * Parses the docx, runs layout, and converts positioned elements to PDF.
 */
export function docx_to_pdf(data: Uint8Array): Uint8Array;

/**
 * Edit a .docx file and return the modified bytes.
 *
 * `data`: original .docx bytes
 * `edits`: JS array of `{paragraph_index, run_index, new_text}` objects
 *
 * Returns the modified .docx as `Uint8Array`.
 */
export function edit_docx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Apply structural edits to a .docx file.
 *
 * `data`: original .docx bytes
 * `edits`: JS array of edit operation objects. Each object has a `type` field:
 *
 * Text operations:
 *   { type: "set_run_text", paragraph_index, run_index, new_text }
 *   { type: "insert_paragraph", index, text, style?, para_style? }
 *   { type: "delete_paragraph", index }
 *   { type: "insert_run", paragraph_index, run_index, text, style? }
 *   { type: "delete_run", paragraph_index, run_index }
 *
 * Formatting:
 *   { type: "set_run_format", paragraph_index, run_index, style }
 *   { type: "set_paragraph_format", paragraph_index, style }
 *
 * Tables:
 *   { type: "insert_table", index, rows, cols, content?, col_widths_pt? }
 *   { type: "insert_table_row", table_index, row_index, cells }
 *   { type: "delete_table_row", table_index, row_index }
 *   { type: "set_cell_text", table_index, row, col, text }
 *
 * Images:
 *   { type: "insert_image", index, data (base64), width_pt, height_pt, content_type }
 *
 * style (RunProps): { bold?, italic?, underline?, font_family?, font_size?, color?, highlight? }
 * para_style (ParaProps): { alignment?, space_before?, space_after?, line_spacing?, indent_left?, style_id? }
 */
export function edit_docx_advanced(data: Uint8Array, edits: any): Uint8Array;

/**
 * Edit a .pptx file and return the modified bytes.
 */
export function edit_pptx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Edit a .pptx and break paragraphs in it, returning the modified bytes.
 *
 * `edits` replaces run text; `splits` cuts a paragraph in two at a character
 * offset, which is what Enter means. Both are applied to the same save, and
 * the text edit lands first so a split counts the characters the file will
 * actually carry.
 */
export function edit_pptx_with_splits(data: Uint8Array, edits: any, splits: any, merges: any, formats: any): Uint8Array;

/**
 * Fast text edit + re-layout using cached document (skips docx parse).
 * Returns layout result. Also updates the cached docx bytes.
 */
export function edit_text_and_relayout(paragraph_index: number, run_index: number, new_text: string): any;

/**
 * Edit a .xlsx file and return the modified bytes.
 */
export function edit_xlsx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Save a workbook a VBA run has changed back into the .xlsx it came from.
 *
 * `run_spreadsheet_vba` hands back the whole workbook rather than a list of
 * edits, so the difference against the original file is worked out here.
 * Everything the difference covers is written: values and formulas, but also
 * fills and fonts, merges, row heights and column widths, what is hidden,
 * frozen panes, the filter, the defined names and which sheets are shown.
 */
export function edit_xlsx_from_workbook(data: Uint8Array, workbook: any): Uint8Array;

/**
 * Renders a number the way a sheet shows it under `format`.
 *
 * The browser used to carry its own reading of number formats; this is the
 * one the rest of the engine uses, so a cell reads the same wherever it is
 * shown.
 */
export function format_cell_number(value: number, format: string): string;

/**
 * Generate a hanko stamp SVG.
 *
 * `config`: JS object with StampConfig fields:
 *   { name: "山田", style: "Round"|"Square"|"Oval", size: 100, date?: "2026.03.13" }
 */
export function generate_hanko_svg(config: any): string;

export function init(): void;

/**
 * Load a document, cache it, and return layout result.
 * Subsequent calls to `edit_text_and_relayout` will reuse the cached parse.
 */
export function layout_document(data: Uint8Array): any;

/**
 * Lay out one text shape the way the engine does, for the browser to draw.
 *
 * `shape` is a `Shape` from `parse_presentation`, `paragraphs` its
 * paragraphs, and `master` / `ph_levels` the inherited outline levels. The
 * answer is one entry per LINE -- its text, where it starts, its baseline,
 * and which paragraph and character offset it came from, so a click can be
 * mapped back to a run.
 *
 * `complete` says whether EVERY paragraph was measured. A shape that comes
 * back incomplete must not be drawn as if it were the engine's layout: the
 * measured tables cover 17 of the 142 families the corpora name, and the rest
 * are still the browser's own wrap.
 */
export function layout_slide_shape(shape: any, paragraphs: any, master: any, ph_levels: any, default_family: string, advances: any): any;

/**
 * List executable Sub and Function entry points in VBA source.
 */
export function list_spreadsheet_vba_procedures(source: string): any;

export function parse_document(data: Uint8Array): any;

/**
 * Parse a PDF file and return its structure as a JS object.
 */
export function parse_pdf(data: Uint8Array): any;

export function parse_presentation(data: Uint8Array): any;

export function parse_spreadsheet(data: Uint8Array): any;

/**
 * Extract all text from a PDF as a single string.
 */
export function pdf_extract_text(data: Uint8Array): string;

/**
 * Verify signatures in a PDF. Returns an array of signature info objects.
 */
export function pdf_verify_signatures(data: Uint8Array): any;

/**
 * Preview a hanko stamp SVG with default config for the given name.
 */
export function preview_hanko(name: string): string;

/**
 * Reads the macros in an `.xlsm` / `.xlam` / `.docm` and says what they could
 * reach. Nothing is executed.
 *
 * Returns an error only when the bytes are not an Office package. A package
 * with no macros in it is not an error: it answers with an empty report,
 * because "there are no macros" is exactly what a caller wants to hear.
 */
export function read_macro_safety(_package: Uint8Array): any;

/**
 * Recalculate every formula in a workbook and hand the workbook back.
 *
 * The browser holds a sheet as this IR while it is being edited, so this is
 * how a typed `=A1+B1` gets an answer: without it the editor would have to
 * write the file out and read it back to find out what it had just computed.
 * Cross-sheet references resolve, because the whole workbook goes across.
 */
export function recalculate_spreadsheet(workbook: any, now?: number | null): any;

/**
 * Execute VBA source against an OxiCells workbook IR.
 */
export function run_spreadsheet_vba(workbook: any, source: string, procedure: string, args: any, active_sheet: number, file_name?: string | null): any;

/**
 * Write comments into a .docx (adds only). See `update_docx_comments` for
 * the full operation set (add + remove + resolve).
 */
export function set_docx_comments(data: Uint8Array, comments: any): Uint8Array;

/**
 * Put rows or columns into a sheet, or take them out, and hand the workbook
 * back.
 *
 * `at` counts rows from one and columns from zero, as the IR does. A negative
 * `count` takes them out.
 *
 * The whole workbook goes across because the change reaches all of it: a
 * formula on any sheet that names this one has to follow the rows it names.
 * That is a cost worth paying here in a way it would not be per keystroke —
 * nobody inserts a row sixty times a second.
 */
export function shift_band(workbook: any, sheet: string, rows: boolean, at: number, count: number): any;

/**
 * What the compiled tables say one character advances, in EM units.
 *
 * The page measures faces itself; this lets it check that measuring against a
 * face the tables also carry, instead of trusting a canvas it never verified.
 */
export function slide_face_advance(family: string, bold: boolean, italic: boolean, ch: string): number | undefined;

/**
 * Whether the engine can measure this family at all, so a caller can tell a
 * person which text on the page is laid out by the engine and which is not.
 */
export function slide_family_measurable(family: string): boolean;

/**
 * Where each character of one laid-out line starts, in points from its `x`.
 *
 * `line` is a `PlacedLine` from [`layout_slide_shape`] and `advances` the same
 * measured faces that call was given. The answer is the engine's own placement
 * -- each advance on the master unit, which is what PowerPoint measures and
 * draws on (see `glyph_offsets_pt`) -- so a page that draws from it puts the
 * glyphs where the break already assumed they were.
 *
 * Null when a face cannot be measured; the caller then has to fall back to its
 * own measuring and should say the answer is not the engine's.
 */
export function slide_glyph_offsets(line: any, advances: any): any;

/**
 * Move a formula's relative references as Excel does when a cell is copied.
 *
 * The browser needs this to fill a formula down a column: `=B2+C2` dragged one
 * row must become `=B3+C3`, while `=$B$2` must not move. A formula the engine
 * cannot parse comes back as an error rather than unchanged, because copying
 * it verbatim would keep relative references Excel would have moved — which
 * looks like it worked and is wrong.
 */
export function translate_formula(formula: string, rows: number, columns: number): string;

/**
 * Apply a batch of comment operations to a .docx:
 * { add: [ { author, initials?, date?, text, paragraph_index, char_start,
 *            char_end, resolved?, parent_index?, parent_para_id? } ],
 *   remove_ids: [ "w:id", … ],
 *   set_resolved: [ { para_id, done } ] }
 * Adds write word/comments.xml + commentsExtended.xml (threads via
 * paraIdParent, resolved via w15:done) and range markers in document.xml;
 * removals strip all three.
 */
export function update_docx_comments(data: Uint8Array, ops: any): Uint8Array;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly break_slide_paragraph: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: any, j: any) => [number, number, number];
    readonly build_docx: (a: any) => [number, number, number, number];
    readonly build_docx_with_template: (a: any, b: number, c: number) => [number, number, number, number];
    readonly create_blank_docx: () => [number, number];
    readonly create_pdf: (a: number, b: number, c: number, d: number) => [number, number];
    readonly docx_to_pdf: (a: number, b: number) => [number, number, number, number];
    readonly edit_docx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_docx_advanced: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_pptx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_pptx_with_splits: (a: number, b: number, c: any, d: any, e: any, f: any) => [number, number, number, number];
    readonly edit_text_and_relayout: (a: number, b: number, c: number, d: number) => [number, number, number];
    readonly edit_xlsx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_xlsx_from_workbook: (a: number, b: number, c: any) => [number, number, number, number];
    readonly format_cell_number: (a: number, b: number, c: number) => [number, number];
    readonly generate_hanko_svg: (a: any) => [number, number, number, number];
    readonly layout_document: (a: number, b: number) => [number, number, number];
    readonly layout_slide_shape: (a: any, b: any, c: any, d: any, e: number, f: number, g: any) => [number, number, number];
    readonly list_spreadsheet_vba_procedures: (a: number, b: number) => [number, number, number];
    readonly parse_document: (a: number, b: number) => [number, number, number];
    readonly parse_pdf: (a: number, b: number) => [number, number, number];
    readonly parse_presentation: (a: number, b: number) => [number, number, number];
    readonly parse_spreadsheet: (a: number, b: number) => [number, number, number];
    readonly pdf_extract_text: (a: number, b: number) => [number, number, number, number];
    readonly pdf_verify_signatures: (a: number, b: number) => [number, number, number];
    readonly preview_hanko: (a: number, b: number) => [number, number];
    readonly read_macro_safety: (a: number, b: number) => [number, number, number];
    readonly recalculate_spreadsheet: (a: any, b: number, c: number) => [number, number, number];
    readonly run_spreadsheet_vba: (a: any, b: number, c: number, d: number, e: number, f: any, g: number, h: number, i: number) => [number, number, number];
    readonly set_docx_comments: (a: number, b: number, c: any) => [number, number, number, number];
    readonly shift_band: (a: any, b: number, c: number, d: number, e: number, f: number) => [number, number, number];
    readonly slide_face_advance: (a: number, b: number, c: number, d: number, e: number, f: number) => number;
    readonly slide_family_measurable: (a: number, b: number) => number;
    readonly slide_glyph_offsets: (a: any, b: any) => [number, number, number];
    readonly translate_formula: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly update_docx_comments: (a: number, b: number, c: any) => [number, number, number, number];
    readonly init: () => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
