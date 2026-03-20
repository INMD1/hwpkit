import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { IMarkdownParser, PandocAst } from '../abstraction/MarkdownParser';

export declare class BrowserFileSystem implements IFileSystem {
    private files;
    constructor(initialFiles?: Record<string, Uint8Array>);
    readFile(filePath: string): Promise<Uint8Array>;
    exists(filePath: string): Promise<boolean>;
    writeFile(filePath: string, data: Uint8Array): Promise<void>;
    addFile(path: string, content: Uint8Array): void;
}
export declare class BrowserZipService implements IZipService {
    private zip;
    load(data: Uint8Array): Promise<void>;
    readFile(entryName: string): Promise<string>;
    readBuffer(entryName: string): Promise<Uint8Array>;
    addFile(entryName: string, content: string | Uint8Array, options?: {
        compression: "STORE" | "DEFLATE";
    }): void;
    getEntries(): string[];
    generate(): Promise<Uint8Array>;
}
export declare class RemarkParser implements IMarkdownParser {
    parse(markdown: string): Promise<PandocAst>;
    private transformToPandoc;
    private transformBlock;
    private transformInlines;
    private transformInline;
}
