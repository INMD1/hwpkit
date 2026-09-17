interface LibreOfficeDocOptions {
    /** Executable path; arguments are passed directly, without a shell. */
    executable?: string;
    timeoutMs?: number;
}
/** Local DOC → DOCX bridge. Requires LibreOffice, never uploads documents. */
declare function createLibreOfficeDocConverter(options?: LibreOfficeDocOptions): (data: Uint8Array) => Promise<Uint8Array>;

export { type LibreOfficeDocOptions, createLibreOfficeDocConverter };
