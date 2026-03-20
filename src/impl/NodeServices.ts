
import * as fs from 'fs';
// import * as path from 'path'; // 미사용
// @ts-ignore - adm-zip은 Node.js 환경에서만 사용
import AdmZip from 'adm-zip';
import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { IMarkdownParser, PandocAst } from '../abstraction/MarkdownParser';
// @ts-ignore - PandocWrapper는 별도 구현 필요
import { PandocWrapper } from '../hwpx/PandocWrapper';

export class NodeFileSystem implements IFileSystem {
    async readFile(filePath: string): Promise<Uint8Array> {
        return new Promise((resolve, reject) => {
            fs.readFile(filePath, (err, data) => {
                if (err) reject(err);
                else resolve(new Uint8Array(data));
            });
        });
    }

    async exists(filePath: string): Promise<boolean> {
        return fs.existsSync(filePath);
    }

    async writeFile(filePath: string, data: Uint8Array): Promise<void> {
        return new Promise((resolve, reject) => {
            fs.writeFile(filePath, Buffer.from(data), (err) => {
                if (err) reject(err);
                else resolve();
            });
        });
    }
}

export class NodeZipService implements IZipService {
    private zip: AdmZip | null = null;

    async load(data: Uint8Array): Promise<void> {
        this.zip = new AdmZip(Buffer.from(data));
    }

    async readFile(entryName: string): Promise<string> {
        if (!this.zip) throw new Error("Zip not loaded");
        return this.zip.readAsText(entryName);
    }

    async readBuffer(entryName: string): Promise<Uint8Array> {
        if (!this.zip) throw new Error("Zip not loaded");
        const entry = this.zip.getEntry(entryName);
        if (!entry) throw new Error(`Entry ${entryName} not found`);
        return new Uint8Array(entry.getData());
    }

    addFile(entryName: string, content: string | Uint8Array): void {
        if (!this.zip) this.zip = new AdmZip();
        if (typeof content === 'string') {
            this.zip.addFile(entryName, Buffer.from(content, 'utf8'));
        } else {
            this.zip.addFile(entryName, Buffer.from(content));
        }
    }

    getEntries(): string[] {
        if (!this.zip) return [];
        return this.zip.getEntries().map((e: any) => e.entryName);
    }

    async generate(): Promise<Uint8Array> {
        if (!this.zip) throw new Error("Zip not loaded");
        return new Uint8Array(this.zip.toBuffer());
    }
}

export class PandocParser implements IMarkdownParser {
    private wrapper: PandocWrapper;

    constructor() {
        this.wrapper = new PandocWrapper();
    }

    async parse(markdown: string): Promise<PandocAst> {
        // PandocWrapper was updated to support string input via convertToAst
        // We need to cast the result to PandocAst
        return await this.wrapper.convertToAst(markdown);
    }
}
