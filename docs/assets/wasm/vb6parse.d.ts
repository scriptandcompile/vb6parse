/* tslint:disable */
/* eslint-disable */

/**
 * Chroma subsampling format
 */
export enum ChromaSampling {
    /**
     * Both vertically and horizontally subsampled.
     */
    Cs420 = 0,
    /**
     * Horizontally subsampled.
     */
    Cs422 = 1,
    /**
     * Not subsampled.
     */
    Cs444 = 2,
    /**
     * Monochrome.
     */
    Cs400 = 3,
}

/**
 * Initializes the panic hook for better error messages in the browser console.
 */
export function init_panic_hook(): void;

/**
 * Parses VB6 code and returns a `PlaygroundOutput` object containing tokens, CST, and errors.
 *
 * # Errors
 *
 * So far we do not correctly handle errors and failures and just panic but this must eventually
 * be converted into an error value.
 *
 * # Panics
 *
 * Currently, we are doing minimal error recovery and checking for the playground as this
 * is an attempt to get the system up and working well enough to demonstrate the possibilities.
 * As is, we can produce a panic if the input can not be tokenized.
 */
export function parse_vb6_code(code: string, _file_type: string): any;

/**
 * Tokenizes VB6 code and returns a list of `TokenInfo` objects for quick preview.
 *
 * # Errors
 *
 * So far we do not correctly handle errors and failures and just panic but this must eventually
 * be converted into an error value.
 *
 * # Panics
 *
 * Currently, we are doing minimal error recovery and checking for the playground as this
 * is an attempt to get the system up and working well enough to demonstrate the possibilities.
 * As is, we can produce a panic if the input can not be tokenized.
 */
export function tokenize_vb6_code(code: string): any;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly init_panic_hook: () => void;
    readonly parse_vb6_code: (a: number, b: number, c: number, d: number) => [number, number, number];
    readonly tokenize_vb6_code: (a: number, b: number) => [number, number, number];
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_externrefs: WebAssembly.Table;
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
