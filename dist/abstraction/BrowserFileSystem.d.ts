import { IFileSystem } from './FileSystem';

export declare class BrowserFileSystem implements IFileSystem {
    readFile(path: string): Promise<Uint8Array>;
    exists(path: string): Promise<boolean>;
    writeFile(path: string, data: string | Uint8Array): Promise<void>;
}
