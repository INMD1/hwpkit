import { default as JSZip } from 'jszip';
import { ConversionState } from './state';

export declare class DocxParser {
    private state;
    constructor(state: ConversionState);
    parseMetadata(zip: JSZip): Promise<void>;
    parseRelationships(zip: JSZip): Promise<void>;
    parseImages(zip: JSZip): Promise<void>;
    parseStyles(zip: JSZip): Promise<void>;
    parseNumbering(zip: JSZip): Promise<void>;
    parseFootnotes(zip: JSZip): Promise<void>;
    parseEndnotes(zip: JSZip): Promise<void>;
    parsePageSetup(doc: Document): void;
}
