import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { PandocAst } from '../abstraction/MarkdownParser';

export declare class HwpxConverter {
    private ast;
    private headerXmlContent;
    private headerDoc;
    private fs;
    private zipService;
    private normalStyleId;
    private normalParaPrId;
    private charPrCache;
    private maxCharPrId;
    private maxParaPrId;
    private images;
    constructor(ast: PandocAst, fileSystem: IFileSystem, zipService: IZipService);
    run(refPath: string, outputPath: string): Promise<void>;
    private parseStylesAndInitXml;
    private processBlocks;
    private createParaStart;
    private handlePara;
    private handlePlain;
    private handleHeader;
    private handleCodeBlock;
    private handleBulletList;
    private getCharPrId;
    private handleImage;
    private processInlines;
}
