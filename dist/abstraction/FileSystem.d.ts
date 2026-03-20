export interface IFileSystem {
    readFile(path: string): Promise<Uint8Array>;
    exists(path: string): Promise<boolean>;
    writeFile(path: string, data: Uint8Array): Promise<void>;
}
