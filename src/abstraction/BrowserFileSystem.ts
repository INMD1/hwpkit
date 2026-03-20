
import { IFileSystem } from './FileSystem';

export class BrowserFileSystem implements IFileSystem {
    async readFile(path: string): Promise<Uint8Array> {
        console.warn(`[BrowserFileSystem] readFile not supported in browser: ${path}`);
        // 빈 배열 반환하여 에러 방지 (이미지 로드 실패 처리됨)
        return new Uint8Array(0);
    }

    async exists(path: string): Promise<boolean> {
        console.warn(`[BrowserFileSystem] exists check: ${path} -> false`);
        return false;
    }

    async writeFile(path: string, data: string | Uint8Array): Promise<void> {
        console.log(`[BrowserFileSystem] Writing to ${path}, size: ${data.length}`);
        // 실제 파일 시스템에는 쓰지 않음 (메모리나 가상 파일 시스템 구현 필요 시 확장)
    }
}
