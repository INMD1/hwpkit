/**
 * Converts Markdown string or File/Blob to a HWP (HML XML) Blob.
 * Hancom Office supports rendering HML format flawlessly.
 *
 * @param markdownInput string (markdown content), File, or Blob
 * @returns Promise<Blob> containing the HML (application/xml) data
 */
export declare function convertMdToHwp(markdownInput: string | File | Blob): Promise<Blob>;
