
import JSZip from 'jszip';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { IMarkdownParser, PandocAst, PandocNode } from '../abstraction/MarkdownParser';

// Browser File System (In-Memory)
export class BrowserFileSystem implements IFileSystem {
    private files: Map<string, Uint8Array> = new Map();

    constructor(initialFiles?: Record<string, Uint8Array>) {
        if (initialFiles) {
            for (const [path, content] of Object.entries(initialFiles)) {
                this.files.set(path, content);
            }
        }
    }

    async readFile(filePath: string): Promise<Uint8Array> {
        const data = this.files.get(filePath);
        if (!data) throw new Error(`File not found: ${filePath}`);
        return data;
    }

    async exists(filePath: string): Promise<boolean> {
        return this.files.has(filePath);
    }

    async writeFile(filePath: string, data: Uint8Array): Promise<void> {
        this.files.set(filePath, data);
    }

    // Helper to add files dynamically
    addFile(path: string, content: Uint8Array) {
        this.files.set(path, content);
    }
}

// Browser Zip Service (JSZip)
export class BrowserZipService implements IZipService {
    private zip: JSZip | null = null;

    async load(data: Uint8Array): Promise<void> {
        this.zip = await JSZip.loadAsync(data);
    }

    async readFile(entryName: string): Promise<string> {
        if (!this.zip) throw new Error("Zip not loaded");
        const file = this.zip.file(entryName);
        if (!file) throw new Error(`Entry ${entryName} not found`);
        return await file.async("string");
    }

    async readBuffer(entryName: string): Promise<Uint8Array> {
        if (!this.zip) throw new Error("Zip not loaded");
        const file = this.zip.file(entryName);
        if (!file) throw new Error(`Entry ${entryName} not found`);
        return await file.async("uint8array");
    }

    addFile(entryName: string, content: string | Uint8Array, options?: { compression: "STORE" | "DEFLATE" }): void {
        if (!this.zip) this.zip = new JSZip();
        this.zip.file(entryName, content, {
            compression: options?.compression || "DEFLATE"
        });
    }

    getEntries(): string[] {
        if (!this.zip) return [];
        const entries: string[] = [];
        this.zip.forEach((relativePath) => {
            entries.push(relativePath);
        });
        return entries;
    }

    async generate(): Promise<Uint8Array> {
        if (!this.zip) throw new Error("Zip not loaded");
        return await this.zip.generateAsync({ type: "uint8array" });
    }
}

// Remark Parser Adapter
export class RemarkParser implements IMarkdownParser {
    async parse(markdown: string): Promise<PandocAst> {
        const processor = unified().use(remarkParse).use(remarkGfm);
        const mdast = processor.parse(markdown);
        return this.transformToPandoc(mdast);
    }

    private transformToPandoc(mdast: any): PandocAst {
        return {
            'pandoc-api-version': [1, 23],
            meta: {},
            blocks: mdast.children.map((child: any) => this.transformBlock(child)).filter((x: any) => x !== null)
        };
    }

    private transformBlock(node: any): PandocNode | null {
        switch (node.type) {
            case 'paragraph':
                return { t: 'Para', c: this.transformInlines(node.children) };
            case 'heading':
                return {
                    t: 'Header',
                    c: [
                        node.depth,
                        ["", [], []], // Attr: [id, classes, kvs]
                        this.transformInlines(node.children)
                    ]
                };
            case 'code':
                return {
                    t: 'CodeBlock',
                    c: [
                        ["", [], []], // Attr
                        node.value
                    ]
                };
            case 'list': // unordered
                // Pandoc BulletList is list of list of blocks
                // Remark list has listItems.
                const items = node.children.map((li: any) => {
                    // li.children are blocks (e.g. paragraph)
                    return li.children.map((b: any) => this.transformBlock(b)).filter((x: any) => x !== null);
                });
                return { t: 'BulletList', c: items };
            case 'blockquote':
                // Not strictly supported in current HwpxConverter but let's map to Plain or similar if needed.
                // HwpxConverter doesn't handle BlockQuote explicitly in processBlocks.
                // Let's ignore or map to Para.
                // Recursively process children?
                // For safety, let's map to Para with some marker or just flatten.
                // HwpxConverter handles Para, Header, Plain, CodeBlock, BulletList.
                // Let's just flatten to Paras for now.
                if (node.children && node.children.length > 0) {
                    // Just take the first child if it's a paragraph
                    return this.transformBlock(node.children[0]);
                }
                return null;
            default:
                // Fallback for unsupported blocks
                return null;
        }
    }

    private transformInlines(nodes: any[]): any[] {
        if (!nodes) return [];
        let result: any[] = [];
        for (const node of nodes) {
            const transformed = this.transformInline(node);
            if (transformed) {
                if (Array.isArray(transformed)) {
                    result = result.concat(transformed);
                } else {
                    result.push(transformed);
                }
            }
        }
        return result;
    }

    private transformInline(node: any): any {
        switch (node.type) {
            case 'text':
                return { t: 'Str', c: node.value };
            case 'emphasis':
                return { t: 'Emph', c: this.transformInlines(node.children) };
            case 'strong':
                return { t: 'Strong', c: this.transformInlines(node.children) };
            case 'inlineCode':
                // Pandoc: Code [attr] string
                // HwpxConverter doesn't explicitly handle Code inline?
                // Checking HwpxConverter... it handles Str, Space, SoftBreak, Strong, Emph, Link, Image.
                // It does NOT handle Code (inline). 
                // Let's transform inlineCode to Str for now or Strong?
                // Wait, HwpxConverter.processInlines lines 352-383.
                // It does recursively process arrays.
                // If I return Str, it works.
                return { t: 'Str', c: node.value };
            case 'link':
                // Pandoc: Link [attr] [text] [target, title]
                return {
                    t: 'Link',
                    c: [
                        ["", [], []],
                        this.transformInlines(node.children),
                        [node.url, node.title || ""]
                    ]
                };
            case 'image':
                // Pandoc: Image [attr] [alt] [target, title]
                // Remark image: url, title, alt (string)
                return {
                    t: 'Image',
                    c: [
                        ["", [], []],
                        [{ t: 'Str', c: node.alt || "" }], // Alt text as inline
                        [node.url, node.title || ""]
                    ]
                };
            case 'break':
                return { t: 'SoftBreak' }; // Or HardBreak? Pandoc has SoftBreak and LineBreak.
            case 'space':
                // Remark typically merges spaces into text? NO.
                // If remark gives space?
                return { t: 'Space' };
            default:
                if (node.value) return { t: 'Str', c: node.value };
                if (node.children) return this.transformInlines(node.children);
                return null;
        }
    }
}
