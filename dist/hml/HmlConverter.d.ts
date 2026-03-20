import { IFileSystem } from '../abstraction/FileSystem';

export declare class HmlConverter {
    private fs;
    private BIN_ITEM_ENTRIES;
    private BIN_STORAGE_ENTRIES;
    private PLACEHOLDERS;
    private ph_counter;
    private borderFills;
    constructor(fileSystem: IFileSystem);
    private register_placeholder;
    private fetchImage;
    private getImageSize;
    private process_image;
    private process_code_block;
    private process_table_native;
    private processInline;
    convert(markdown: string): Promise<string>;
    private HEADER;
    private process_html_table;
    /**
     * CSS style 문자열에서 border/background 정보를 파싱하여
     * HML BORDERFILL XML을 생성합니다.
     */
    private parseCssToHmlBorderFill;
    private cssColorToHmlInt;
}
