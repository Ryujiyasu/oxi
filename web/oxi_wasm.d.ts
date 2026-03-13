/* tslint:disable */
/* eslint-disable */

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
 * Edit a .docx file and return the modified bytes.
 *
 * `data`: original .docx bytes
 * `edits`: JS array of `{paragraph_index, run_index, new_text}` objects
 *
 * Returns the modified .docx as `Uint8Array`.
 */
export function edit_docx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Edit a .pptx file and return the modified bytes.
 */
export function edit_pptx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Edit a .xlsx file and return the modified bytes.
 */
export function edit_xlsx(data: Uint8Array, edits: any): Uint8Array;

/**
 * Generate a hanko stamp SVG.
 *
 * `config`: JS object with StampConfig fields:
 *   { name: "山田", style: "Round"|"Square"|"Oval", size: 100, date?: "2026.03.13" }
 */
export function generate_hanko_svg(config: any): string;

export function init(): void;

export function layout_document(data: Uint8Array): any;

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

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly create_blank_docx: () => [number, number];
    readonly create_pdf: (a: number, b: number, c: number, d: number) => [number, number];
    readonly edit_docx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_pptx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly edit_xlsx: (a: number, b: number, c: any) => [number, number, number, number];
    readonly generate_hanko_svg: (a: any) => [number, number, number, number];
    readonly layout_document: (a: number, b: number) => [number, number, number];
    readonly parse_document: (a: number, b: number) => [number, number, number];
    readonly parse_pdf: (a: number, b: number) => [number, number, number];
    readonly parse_presentation: (a: number, b: number) => [number, number, number];
    readonly parse_spreadsheet: (a: number, b: number) => [number, number, number];
    readonly pdf_extract_text: (a: number, b: number) => [number, number, number, number];
    readonly pdf_verify_signatures: (a: number, b: number) => [number, number, number];
    readonly preview_hanko: (a: number, b: number) => [number, number];
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
