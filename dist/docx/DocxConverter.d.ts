export declare class DocxConverter {
    /**
     * Converts a DOCX file buffer to Markdown string.
     * @param arrayBuffer The ArrayBuffer of the DOCX file.
     * @returns Promise resolving to the Markdown string.
     */
    static toMarkdown(arrayBuffer: ArrayBuffer): Promise<string>;
}
