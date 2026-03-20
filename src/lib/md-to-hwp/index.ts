import { HmlConverter } from '../../hml/HmlConverter';
import { BrowserFileSystem } from '../../abstraction/BrowserFileSystem';

/**
 * Converts Markdown string or File/Blob to a HWP (HML XML) Blob.
 * Hancom Office supports rendering HML format flawlessly.
 * 
 * @param markdownInput string (markdown content), File, or Blob
 * @returns Promise<Blob> containing the HML (application/xml) data
 */
export async function convertMdToHwp(markdownInput: string | File | Blob): Promise<Blob> {
    let text = '';

    if (typeof markdownInput === 'string') {
        text = markdownInput;
    } else if (markdownInput instanceof File || markdownInput instanceof Blob) {
        text = await markdownInput.text();
    } else {
        throw new Error("Invalid input type. Expected string, File, or Blob.");
    }

    const fs = new BrowserFileSystem();
    const converter = new HmlConverter(fs);

    try {
        const hmlContent = await converter.convert(text);
        return new Blob([hmlContent], { type: 'application/xml' });
    } catch (err: any) {
        console.error("MD to HWP conversion error", err);
        throw new Error(`Failed to convert MD to HWP: ${err.message}`);
    }
}
