// @ts-ignore - @xmldom/xmldom은 Node.js 환경에서 사용, 브라우저에서는 네이티브 DOMParser 사용
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { xmlEscape } from './utils';
import { HwpxTemplate } from './HwpxTemplate';
import { IFileSystem } from '../abstraction/FileSystem';
import { IZipService } from '../abstraction/ZipService';
import { PandocAst, PandocNode } from '../abstraction/MarkdownParser';

// Types for Pandoc AST
// (Already imported or defined in abstraction, but local interface matches)

export class HwpxConverter {
    private ast: PandocAst;
    private headerXmlContent: string | null = null;
    private headerDoc: Document | null = null; // XML DOM

    // Services
    private fs: IFileSystem;
    private zipService: IZipService;

    // Style Maps
    private normalStyleId: string = "0";
    private normalParaPrId: string = "1";

    // Caches
    private charPrCache: { [key: string]: string } = {};
    private maxCharPrId: number = 0;
    private maxParaPrId: number = 0;

    // Images
    private images: { id: string, path: string, ext: string }[] = [];

    constructor(ast: PandocAst, fileSystem: IFileSystem, zipService: IZipService) {
        this.ast = ast;
        this.fs = fileSystem;
        this.zipService = zipService;
    }

    public async run(refPath: string, outputPath: string) {
        // Read Reference
        // Ensure refPath exists (handled by caller or check here?)
        // In browser, refPath might be an object URL or key. 
        // We assume fileSystem.readFile handles 'refPath'.

        const refData = await this.fs.readFile(refPath);
        await this.zipService.load(refData);

        try {
            this.headerXmlContent = await this.zipService.readFile('Contents/header.xml');
        } catch (e) {
            throw new Error("Invalid reference HWPX: missing Contents/header.xml");
        }

        // Initialize Header DOM
        this.parseStylesAndInitXml();

        // Convert Blocks
        const xmlBody = await this.processBlocks(this.ast.blocks);

        // Serialize Header
        const serializer = new XMLSerializer();
        const newHeaderXml = serializer.serializeToString(this.headerDoc!);

        // Prepare Output Zip (We modify the loaded zip or create new? Implementation detail of ZipService)
        // Implementation: We can modify the loaded one.

        // 1. Mimetype (MUST be first, uncompressed)
        this.zipService.addFile('mimetype', HwpxTemplate.MIMETYPE, { compression: "STORE" });

        // 2. META-INF
        this.zipService.addFile('META-INF/container.xml', HwpxTemplate.CONTAINER_XML);
        this.zipService.addFile('META-INF/container.rdf', HwpxTemplate.CONTAINER_RDF);

        // Update Header
        this.zipService.addFile('Contents/header.xml', newHeaderXml);

        // Update Section0
        const section0Path = 'Contents/section0.xml';
        let originalSection = await this.zipService.readFile(section0Path);

        // Parse Section0
        const doc = new DOMParser().parseFromString(originalSection, 'text/xml');
        const sec = doc.getElementsByTagName('hs:sec')[0];

        if (!sec) {
            throw new Error("Invalid section0.xml: missing hs:sec");
        }

        // Ensure namespaces
        if (!doc.documentElement.getAttribute('xmlns:hc')) {
            doc.documentElement.setAttribute('xmlns:hc', 'http://www.hancom.co.kr/hwpml/2011/core');
        }
        if (!doc.documentElement.getAttribute('xmlns:hp')) {
            doc.documentElement.setAttribute('xmlns:hp', 'http://www.hancom.co.kr/hwpml/2011/paragraph');
        }

        // Remove existing children of sec (except definitions if any? usually sec contains p)
        while (sec.firstChild) {
            sec.removeChild(sec.firstChild);
        }

        // Parse new body content
        // Note: xmlBody is a string of <hp:p>...</hp:p>
        // We need to parse this string into nodes and append them.
        // We wrap it in a dummy tag to parse.
        const bodyDoc = new DOMParser().parseFromString(`<root xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">${xmlBody}</root>`, 'text/xml');

        // Import nodes
        // Note: xmldom might not support importNode correctly across docs in some versions, but usually does.
        // Alternatively, manual traversal.
        // Let's try importNode.

        const childNodes = bodyDoc.documentElement.childNodes;
        for (let i = 0; i < childNodes.length; i++) {
            const imported = doc.importNode(childNodes[i], true);
            sec.appendChild(imported);
        }

        const newContent = serializer.serializeToString(doc);
        // Ensure XML declaration
        const finalSectionXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' + newContent;
        this.zipService.addFile(section0Path, finalSectionXml);

        // Update Manifest
        const manifestPath = 'Contents/content.hpf';
        let hpf = await this.zipService.readFile(manifestPath);

        if (this.images.length > 0) {
            const manifestClose = hpf.indexOf('</opf:manifest>');
            if (manifestClose !== -1) {
                let items = "";
                for (const img of this.images) {
                    let mime = "image/png";
                    if (img.ext === 'jpg' || img.ext === 'jpeg') mime = "image/jpeg";
                    else if (img.ext === 'gif') mime = "image/gif";
                    else if (img.ext === 'bmp') mime = "image/bmp";

                    items += `<opf:item id="${img.id}" href="BinData/${img.id}.${img.ext}" media-type="${mime}" isEmbeded="1"/>\n`;
                }
                hpf = hpf.substring(0, manifestClose) + items + hpf.substring(manifestClose);
            }
            this.zipService.addFile(manifestPath, hpf);
        }

        // Handle Images (Bindata)
        for (const img of this.images) {
            try {
                if (await this.fs.exists(img.path)) {
                    const imgData = await this.fs.readFile(img.path);
                    this.zipService.addFile(`Contents/BinData/${img.id}.${img.ext}`, imgData);
                } else {
                    console.warn(`Image file not found: ${img.path}`);
                }
            } catch (e) {
                console.warn(`Failed to embed image: ${img.path}`);
            }
        }

        // Generate Output
        const outBuffer = await this.zipService.generate();
        await this.fs.writeFile(outputPath, outBuffer);
        console.log(`Created HWPX: ${outputPath}`);
    }

    private parseStylesAndInitXml() {
        if (!this.headerXmlContent) return;
        this.headerDoc = new DOMParser().parseFromString(this.headerXmlContent, 'text/xml');
        const doc = this.headerDoc;

        // 1. Find Max IDs
        const charPrs = doc!.getElementsByTagName('hh:charPr');
        for (let i = 0; i < charPrs.length; i++) {
            const id = parseInt(charPrs[i].getAttribute('id') || "0");
            if (id > this.maxCharPrId) this.maxCharPrId = id;
        }

        const paraPrs = doc!.getElementsByTagName('hh:paraPr');
        for (let i = 0; i < paraPrs.length; i++) {
            const id = parseInt(paraPrs[i].getAttribute('id') || "0");
            if (id > this.maxParaPrId) this.maxParaPrId = id;
        }

        // 2. Normal Style
        const styles = doc!.getElementsByTagName('hh:style');
        let normalStyle: Element | null = null;

        for (let i = 0; i < styles.length; i++) {
            if (styles[i].getAttribute('id') === "0") {
                normalStyle = styles[i];
                break;
            }
        }
        if (!normalStyle && styles.length > 0) normalStyle = styles[0];

        if (normalStyle) {
            this.normalStyleId = normalStyle.getAttribute('id') || "0";
            this.normalParaPrId = normalStyle.getAttribute('paraPrIDRef') || "1";
        }
    }

    private async processBlocks(blocks: PandocNode[]): Promise<string> {
        let res: string[] = [];
        for (const block of blocks) {
            if (block.t === 'Para') {
                res.push(this.handlePara(block.c));
            } else if (block.t === 'Header') {
                res.push(this.handleHeader(block.c));
            } else if (block.t === 'Plain') {
                res.push(this.handlePlain(block.c));
            } else if (block.t === 'CodeBlock') {
                res.push(this.handleCodeBlock(block.c));
            } else if (block.t === 'BulletList') {
                res.push(this.handleBulletList(block.c));
            }
        }
        return res.join('\n');
    }

    // --- Handlers ---
    private createParaStart(styleId = this.normalStyleId, paraPrId = this.normalParaPrId): string {
        return `<hp:p paraPrIDRef="${paraPrId}" styleIDRef="${styleId}" pageBreak="0" columnBreak="0" merged="0">`;
    }

    private handlePara(inlines: PandocNode[]): string {
        let xml = this.createParaStart();
        xml += this.processInlines(inlines);
        xml += '</hp:p>';
        return xml;
    }

    private handlePlain(inlines: PandocNode[]): string {
        return this.handlePara(inlines);
    }

    private handleHeader(content: any[]): string {
        const inlines = content[2];
        let xml = this.createParaStart();
        const activeFormats = new Set<string>();
        activeFormats.add('BOLD'); // Stub for Header style
        xml += this.processInlines(inlines, "0", activeFormats);
        xml += '</hp:p>';
        return xml;
    }

    private handleCodeBlock(content: any[]): string {
        const code = content[1];
        let xml = this.createParaStart();
        xml += `<hp:run charPrIDRef="0"><hp:t>${xmlEscape(code)}</hp:t></hp:run>`;
        xml += '</hp:p>';
        return xml;
    }

    private handleBulletList(items: PandocNode[][]): string {
        let res = "";
        for (const itemBlocks of items) {
            for (const b of itemBlocks) {
                if (b.t === 'Para') {
                    let xml = this.createParaStart(); // Should use List style
                    xml += `<hp:run charPrIDRef="0"><hp:t>• </hp:t></hp:run>`; // Fake bullet
                    xml += this.processInlines(b.c);
                    xml += '</hp:p>';
                    res += xml;
                }
            }
        }
        return res;
    }

    // --- Style Management ---

    private getCharPrId(baseId: string, formats: Set<string>): string {
        if (!formats || formats.size === 0) return baseId;

        const fmtKey = Array.from(formats).sort().join(',');
        const cacheKey = `${baseId}|${fmtKey}`;
        if (this.charPrCache[cacheKey]) return this.charPrCache[cacheKey];

        if (!this.headerDoc) return baseId;
        const doc = this.headerDoc;

        let baseNode: Element | null = null;
        const charPrs = doc.getElementsByTagName('hh:charPr');
        for (let i = 0; i < charPrs.length; i++) {
            if (charPrs[i].getAttribute('id') === baseId) {
                baseNode = charPrs[i];
                break;
            }
        }

        if (!baseNode) {
            for (let i = 0; i < charPrs.length; i++) {
                if (charPrs[i].getAttribute('id') === "0") {
                    baseNode = charPrs[i];
                    break;
                }
            }
        }
        if (!baseNode) return baseId;

        const newNode = baseNode.cloneNode(true) as Element;
        this.maxCharPrId++;
        const newId = this.maxCharPrId.toString();
        newNode.setAttribute('id', newId);

        if (formats.has('BOLD')) {
            if (newNode.getElementsByTagName('hh:bold').length === 0) {
                newNode.appendChild(doc.createElement('hh:bold'));
            }
        }
        if (formats.has('ITALIC')) {
            if (newNode.getElementsByTagName('hh:italic').length === 0) {
                newNode.appendChild(doc.createElement('hh:italic'));
            }
        }
        if (formats.has('UNDERLINE')) {
            let ul = newNode.getElementsByTagName('hh:underline')[0];
            if (!ul) {
                ul = doc.createElement('hh:underline');
                newNode.appendChild(ul);
            }
            ul.setAttribute('type', 'BOTTOM');
            ul.setAttribute('shape', 'SOLID');
            ul.setAttribute('color', '#000000');
        }

        const charProps = doc.getElementsByTagName('hh:charProperties')[0];
        if (charProps) {
            charProps.appendChild(newNode);
            const cnt = parseInt(charProps.getAttribute('itemCnt') || "0");
            charProps.setAttribute('itemCnt', (cnt + 1).toString());
        }

        this.charPrCache[cacheKey] = newId;
        return newId;
    }

    private handleImage(content: any[], charPrId: string): string {
        const target = content[2];
        const url_raw = target[0];

        let imgPath = url_raw;
        // In browser/abstract FS, paths are keys. We don't resolve against CWD.
        // But for Node usage ensuring absolute path might be needed OUTSIDE HwpxConverter
        // or FileSystem implementation handles it.
        // We'll leave it as is, expecting valid keys for FileSystem.

        // Get extension from path
        let ext = 'png';
        const parts = imgPath.split('.');
        if (parts.length > 1) {
            ext = parts[parts.length - 1].toLowerCase();
        }

        const timestamp = new Date().getTime();
        const rand = Math.floor(Math.random() * 1000);
        const imgId = `img_${timestamp}_${rand}`;

        this.images.push({ id: imgId, path: imgPath, ext: ext });

        const width = 28346;
        const height = 28346;

        const instId = Math.floor(Math.random() * 100000000);
        const picId = Math.floor(Math.random() * 100000000);

        return `<hp:run charPrIDRef="${charPrId}"><hp:pic id="${picId}" zOrder="0" numberingType="NONE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${instId}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${width}" height="${height}"/><hp:curSz width="${width}" height="${height}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="0" centerY="0" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hc:img binaryItemIDRef="${imgId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${width}" y="0"/><hc:pt2 x="${width}" y="${height}"/><hc:pt3 x="0" y="${height}"/></hp:imgRect><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="0" dimheight="0"/><hp:effects/><hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment/></hp:pic></hp:run>`;
    }

    private processInlines(inlines: PandocNode[], baseCharPrId = "0", activeFormats: Set<string> = new Set()): string {
        let res = "";
        if (!Array.isArray(inlines)) return "";

        for (const item of inlines) {
            const currentPrId = this.getCharPrId(baseCharPrId, activeFormats);

            if (item.t === 'Str') {
                res += `<hp:run charPrIDRef="${currentPrId}"><hp:t>${xmlEscape(item.c)}</hp:t></hp:run>`;
            } else if (item.t === 'Space') {
                res += `<hp:run charPrIDRef="${currentPrId}"><hp:t> </hp:t></hp:run>`;
            } else if (item.t === 'SoftBreak') {
                res += `<hp:run charPrIDRef="${currentPrId}"><hp:t> </hp:t></hp:run>`;
            } else if (item.t === 'Strong') {
                const newFmt = new Set(activeFormats);
                newFmt.add('BOLD');
                res += this.processInlines(item.c, baseCharPrId, newFmt);
            } else if (item.t === 'Emph') {
                const newFmt = new Set(activeFormats);
                newFmt.add('ITALIC');
                res += this.processInlines(item.c, baseCharPrId, newFmt);
            } else if (item.t === 'Link') {
                const newFmt = new Set(activeFormats);
                newFmt.add('UNDERLINE');
                res += this.processInlines(item.c[1], baseCharPrId, newFmt);
            } else if (item.t === 'Image') {
                res += this.handleImage(item.c, currentPrId);
            } else {
                if (Array.isArray(item.c)) {
                    res += this.processInlines(item.c, baseCharPrId, activeFormats);
                }
            }
        }
        return res;
    }
}

