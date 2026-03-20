
export interface IFileSystem {
    readFile(path: string): Promise<Uint8Array>;
    exists(path: string): Promise<boolean>;
    // Write support might be needed for tests or Node usage
    writeFile(path: string, data: Uint8Array): Promise<void>;
}
