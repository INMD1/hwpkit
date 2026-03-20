
import { HwpxConverter } from '../hwpx/HwpxConverter';
import { BrowserFileSystem, BrowserZipService, RemarkParser } from '../impl/BrowserServices';


import { HmlConverter } from '../hml/HmlConverter';

/**
 * Converts Markdown string to HWPX Blob in a browser environment.
 * @param markdown Markdown content string
 * @param images Record of image paths to Uint8Array content (e.g. {'img/foo.png': buffer})
 * @param referenceHwpx Uint8Array of the reference blank.hwpx file
 * @returns Promise<Uint8Array> The generated HWPX file content
 */
export async function convertMarkdownToHwpx(
    markdown: string,
    images: Record<string, Uint8Array>,
    referenceHwpx: Uint8Array
): Promise<Uint8Array> {

    // 1. Setup Services
    const fs = new BrowserFileSystem(images);
    const zip = new BrowserZipService();
    const parser = new RemarkParser();

    // 2. Parse
    const ast = await parser.parse(markdown);

    // 3. Convert
    // We need to provide a path for reference HWPX.
    // In BrowserFileSystem, we can store it in a virtual path.
    const refPath = '/virtual/blank.hwpx';
    fs.addFile(refPath, referenceHwpx);

    const converter = new HwpxConverter(ast, fs, zip);

    const outputPath = '/virtual/output.hwpx';
    await converter.run(refPath, outputPath);

    // 4. Retrieve
    return await fs.readFile(outputPath);
}

/**
 * Converts Markdown string to HML string in a browser environment.
 * @param markdown Markdown content string
 * @param images Record of image paths to Uint8Array content
 * @returns Promise<string> The generated HML string
 */
export async function convertMarkdownToHml(
    markdown: string,
    images: Record<string, Uint8Array>
): Promise<string> {
    const fs = new BrowserFileSystem(images);
    const converter = new HmlConverter(fs);
    return await converter.convert(markdown);
}

import { DocxConverter } from '../docx/DocxConverter';

/**
 * Converts DOCX ArrayBuffer to HWPX Blob in a browser environment.
 * @param docxBuffer ArrayBuffer of the DOCX file
 * @param images Record of image paths to Uint8Array content
 * @param referenceHwpx Uint8Array of the reference blank.hwpx file
 * @returns Promise<Uint8Array> The generated HWPX file content
 */
export async function convertDocxToHwpx(
    docxBuffer: ArrayBuffer,
    images: Record<string, Uint8Array>,
    referenceHwpx: Uint8Array
): Promise<Uint8Array> {
    const markdown = await DocxConverter.toMarkdown(docxBuffer);
    return await convertMarkdownToHwpx(markdown, images, referenceHwpx);
}

/**
 * Converts DOCX ArrayBuffer to HML string in a browser environment.
 * @param docxBuffer ArrayBuffer of the DOCX file
 * @param images Record of image paths to Uint8Array content
 * @returns Promise<string> The generated HML string
 */
export async function convertDocxToHml(
    docxBuffer: ArrayBuffer,
    images: Record<string, Uint8Array>
): Promise<string> {
    const markdown = await DocxConverter.toMarkdown(docxBuffer);
    return await convertMarkdownToHml(markdown, images);
}
