import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { IMarkdownParser, PandocAst } from '../abstraction/MarkdownParser';

export declare class NodeFileSystem implements IFileSystem {
    readFile(filePath: string): Promise<Uint8Array>;
    exists(filePath: string): Promise<boolean>;
    writeFile(filePath: string, data: Uint8Array): Promise<void>;
}
export declare class NodeZipService implements IZipService {
    private zip;
    load(data: Uint8Array): Promise<void>;
    readFile(entryName: string): Promise<string>;
    readBuffer(entryName: string): Promise<Uint8Array>;
    addFile(entryName: string, content: string | Uint8Array): void;
    getEntries(): string[];
    generate(): Promise<Uint8Array>;
}
export declare class PandocParser implements IMarkdownParser {
    private wrapper;
    constructor();
    parse(markdown: string): Promise<PandocAst>;
}
