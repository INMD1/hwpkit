
// @ts-ignore - mammoth는 별도 설치 필요
import mammoth from 'mammoth';

export class DocxConverter {
    /**
     * Converts a DOCX file buffer to Markdown string.
     * @param arrayBuffer The ArrayBuffer of the DOCX file.
     * @returns Promise resolving to the Markdown string.
     */
    static async toMarkdown(arrayBuffer: ArrayBuffer): Promise<string> {
        try {
            const mammothAny = mammoth as any;
            const result = await mammothAny.convertToMarkdown({ arrayBuffer: arrayBuffer });
            if (result.messages && result.messages.length > 0) {
                console.warn("Mammoth messages:", result.messages);
            }
            return result.value;
        } catch (error) {
            console.error("DOCX conversion failed:", error);
            throw new Error("Failed to convert DOCX to Markdown: " + (error as any).message);
        }
    }
}
