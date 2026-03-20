export interface IZipService {
    load(data: Uint8Array): Promise<void>;
    readFile(entryName: string): Promise<string>;
    readBuffer(entryName: string): Promise<Uint8Array>;
    addFile(entryName: string, content: string | Uint8Array, options?: {
        compression: "STORE" | "DEFLATE";
    }): void;
    getEntries(): string[];
    generate(): Promise<Uint8Array>;
}
