/* tslint:disable */
/* eslint-disable */

/**
 * WASMコンバーターのメインインスタンス
 * JavaScriptからこのオブジェクトを作成して使用します。
 */
export class WasmConverter {
    free(): void;
    [Symbol.dispose](): void;
    /**
     * 外部フォントデータを追加（実行時にフォントを読み込み）
     * コンパイル後でも外部からフォントを追加できます。
     * TTFまたはOTFフォーマットのバイト列を受け付けます。
     * @param name フォント名（例: "NotoSansJP-Regular", "Meiryo"）
     * @param data フォントファイルのバイト列（Uint8Array）
     */
    addFont(name: string, data: Uint8Array): void;
    /**
     * ファイルを画像に変換してZIPで返す
     * @param filename ファイル名
     * @param data ファイルのバイト列
     * @param dpi 画像の解像度（デフォルト: 150）
     * @returns ZIPバイト列（各ページがPNG画像）
     */
    convertToImagesZip(filename: string, data: Uint8Array, dpi?: number | null): Uint8Array;
    /**
     * ファイルをJSON形式のドキュメントモデルに変換（デバッグ用）
     * @param filename ファイル名
     * @param data ファイルのバイト列
     * @returns ドキュメントモデルのJSON文字列
     */
    convertToJson(filename: string, data: Uint8Array): string;
    /**
     * ファイルをPDFに変換
     * @param filename ファイル名（拡張子でフォーマットを判定）
     * @param data ファイルのバイト列
     * @returns PDFバイト列
     */
    convertToPdf(filename: string, data: Uint8Array): Uint8Array;
    /**
     * 読み込まれた外部フォントの数を取得
     */
    externalFontCount(): number;
    /**
     * いずれかのフォントが利用可能かどうか（外部フォント含む）
     */
    hasAnyFont(): boolean;
    /**
     * 日本語内蔵フォントが利用可能かどうか
     */
    hasJapaneseFont(): boolean;
    /**
     * 利用可能なフォント名の一覧をJSON配列で取得
     */
    listFonts(): string;
    /**
     * 新しいコンバーターインスタンスを作成
     */
    constructor();
    /**
     * 外部フォントを削除
     * @param name 削除するフォント名
     */
    removeFont(name: string): void;
    /**
     * サポートされているフォーマット一覧をJSON文字列で取得
     */
    supportedFormats(): string;
}

/**
 * 簡易変換関数（インスタンスなしで使用可能）
 * @param filename ファイル名
 * @param data ファイルのバイト列
 * @param output_format "pdf" または "images_zip"
 * @returns 変換結果のバイト列
 */
export function convertDocument(filename: string, data: Uint8Array, output_format: string): Uint8Array;

/**
 * バージョン情報を取得
 */
export function getVersion(): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_wasmconverter_free: (a: number, b: number) => void;
    readonly convertDocument: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly getVersion: () => [number, number];
    readonly wasmconverter_addFont: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmconverter_convertToImagesZip: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly wasmconverter_convertToJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly wasmconverter_convertToPdf: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly wasmconverter_externalFontCount: (a: number) => number;
    readonly wasmconverter_hasAnyFont: (a: number) => number;
    readonly wasmconverter_hasJapaneseFont: (a: number) => number;
    readonly wasmconverter_listFonts: (a: number) => [number, number];
    readonly wasmconverter_new: () => number;
    readonly wasmconverter_removeFont: (a: number, b: number, c: number) => void;
    readonly wasmconverter_supportedFormats: (a: number) => [number, number];
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
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
