/**
 * Converts Markdown string to HWPX Blob in a browser environment.
 * @param markdown Markdown content string
 * @param images Record of image paths to Uint8Array content (e.g. {'img/foo.png': buffer})
 * @param referenceHwpx Uint8Array of the reference blank.hwpx file
 * @returns Promise<Uint8Array> The generated HWPX file content
 */
export declare function convertMarkdownToHwpx(markdown: string, images: Record<string, Uint8Array>, referenceHwpx: Uint8Array): Promise<Uint8Array>;
/**
 * Converts Markdown string to HML string in a browser environment.
 * @param markdown Markdown content string
 * @param images Record of image paths to Uint8Array content
 * @returns Promise<string> The generated HML string
 */
export declare function convertMarkdownToHml(markdown: string, images: Record<string, Uint8Array>): Promise<string>;
/**
 * Converts DOCX ArrayBuffer to HWPX Blob in a browser environment.
 * @param docxBuffer ArrayBuffer of the DOCX file
 * @param images Record of image paths to Uint8Array content
 * @param referenceHwpx Uint8Array of the reference blank.hwpx file
 * @returns Promise<Uint8Array> The generated HWPX file content
 */
export declare function convertDocxToHwpx(docxBuffer: ArrayBuffer, images: Record<string, Uint8Array>, referenceHwpx: Uint8Array): Promise<Uint8Array>;
/**
 * Converts DOCX ArrayBuffer to HML string in a browser environment.
 * @param docxBuffer ArrayBuffer of the DOCX file
 * @param images Record of image paths to Uint8Array content
 * @returns Promise<string> The generated HML string
 */
export declare function convertDocxToHml(docxBuffer: ArrayBuffer, images: Record<string, Uint8Array>): Promise<string>;
