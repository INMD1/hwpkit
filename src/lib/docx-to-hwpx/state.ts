import { CharShape, ParaShape, BorderFill, StyleData, MetaData, ImageInfo } from './types';

export class ConversionState {
    public ns = {
        w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
        pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    };

    public hwpxNs = {
        ha: 'http://www.hancom.co.kr/hwpml/2011/app',
        hp: 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        hp10: 'http://www.hancom.co.kr/hwpml/2016/paragraph',
        hs: 'http://www.hancom.co.kr/hwpml/2011/section',
        hc: 'http://www.hancom.co.kr/hwpml/2011/core',
        hh: 'http://www.hancom.co.kr/hwpml/2011/head',
        hhs: 'http://www.hancom.co.kr/hwpml/2011/history',
        hm: 'http://www.hancom.co.kr/hwpml/2011/master-page',
        hpf: 'http://www.hancom.co.kr/schema/2011/hpf',
        dc: 'http://purl.org/dc/elements/1.1/',
        opf: 'http://www.idpf.org/2007/opf/',
        ooxmlchart: 'http://www.hancom.co.kr/hwpml/2016/ooxmlchart',
        epub: 'http://www.idpf.org/2007/ops',
        config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
        hv: 'http://www.hancom.co.kr/hwpml/2011/version',
        ocf: 'urn:oasis:names:tc:opendocument:xmlns:container',
        odf: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
        rdf: 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
        odfPkg: 'http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document'
    };

    public HWPUNIT_PER_INCH = 7200;
    public EMU_PER_INCH = 914400;

    // 페이지 설정 (기본값 A4)
    public PAGE_WIDTH = 59530;
    public PAGE_HEIGHT = 84190;
    public MARGIN_TOP = 5668;
    public MARGIN_LEFT = 8505;
    public MARGIN_RIGHT = 8505;
    public MARGIN_BOTTOM = 4252;
    public HEADER_MARGIN = 4252;
    public FOOTER_MARGIN = 4252;
    public GUTTER_MARGIN = 0;

    public COL_COUNT = 1;
    public COL_GAP = 0;
    public COL_TYPE = 'ONE';
    public TEXT_WIDTH: number;

    public charShapes: Map<string, CharShape> = new Map();
    public paraShapes: Map<string, ParaShape> = new Map();
    public borderFills: Map<string, BorderFill> = new Map();

    public idCounter = 3121190098;
    public tableIdCounter = 2085132891;
    public borderFillIdCounter = 7; // IDs 1-6 are pre-initialized
    public imageCounter = 1;
    public currentVertPos = 0;

    // Footnote/Endnote
    public footnoteNumber = 0;
    public endnoteNumber = 0;

    public metadata: MetaData = {
        title: '', creator: '', subject: '', description: ''
    };

    public relationships = new Map<string, { target: string, type: string }>();
    public images = new Map<string, ImageInfo>();

    // Fonts
    public langFontFaces: { [key: string]: Map<string, string> } = {
        HANGUL: new Map(),
        LATIN: new Map(),
        HANJA: new Map(),
        JAPANESE: new Map(),
        OTHER: new Map(),
        SYMBOL: new Map(),
        USER: new Map()
    };

    // Numbering
    public numberingMap = new Map<string, string>(); // numId -> abstractNumId
    public abstractNumMap = new Map<string, Map<string, any>>(); // abstractNumId -> levels
    public listCounters = new Map<string, number>(); // "numId:ilvl" -> counter

    public docStyles = new Map<string, StyleData>();
    public docxStyleToHwpxId: { [key: string]: number } = {};

    public footnoteMap = new Map<string, Element>();
    public endnoteMap = new Map<string, Element>();

    public isFirstParagraph = true;

    constructor() {
        this.TEXT_WIDTH = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT - this.GUTTER_MARGIN;

        this.registerFontForLang('HANGUL', '함초롬바탕');
        this.registerFontForLang('LATIN', '함초롬바탕');
        this.registerFontForLang('HANJA', '함초롬바탕');
        this.registerFontForLang('JAPANESE', '함초롬바탕');
        this.registerFontForLang('OTHER', '함초롬바탕');
        this.registerFontForLang('SYMBOL', '함초롬바탕');
        this.registerFontForLang('USER', '함초롬바탕');

        this.initDocxStyleToHwpxId();
        this.initDefaultStyles();
    }

    public registerFontForLang(lang: string, face: string): number {
        const map = this.langFontFaces[lang];
        if (!map) return 0;

        if (!map.has(face)) {
            map.set(face, String(map.size));
        }
        return parseInt(map.get(face)!);
    }

    private initDocxStyleToHwpxId() {
        this.docxStyleToHwpxId = {
            'Normal': 0,
            'Heading1': 12, 'Heading2': 14, 'Heading3': 16,
            'Heading4': 18, 'Heading5': 20, 'Heading6': 22,
            'Heading7': 24, 'Heading8': 26, 'Heading9': 28,
            'Title': 46, 'Subtitle': 41,
            'ListParagraph': 35, 'NoSpacing': 36,
            'Quote': 38, 'IntenseQuote': 32,
            'Header': 10, 'Footer': 7,
            'FootnoteText': 51, 'EndnoteText': 49,
            'TOCHeading': 45,
            'TOC1': 53, 'TOC2': 54, 'TOC3': 55,
            'TOC4': 56, 'TOC5': 57, 'TOC6': 58,
            'TOC7': 59, 'TOC8': 60, 'TOC9': 61,
            'Caption': 2,
            'TableofFigures': 52,
        };
    }

    private initDefaultStyles() {
        // BorderFills 1-6 (JS 원본과 동일하게 사전 초기화)
        this.borderFills.set('1', {
            id: '1',
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#FFFFFF'
        });
        this.borderFills.set('2', {
            id: '2',
            leftBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#FFFFFF'
        });
        this.borderFills.set('3', {
            id: '3',
            leftBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#FFFFFF'
        });
        this.borderFills.set('4', {
            id: '4',
            leftBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            rightBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            topBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            bottomBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#FFFFFF'
        });
        this.borderFills.set('5', {
            id: '5',
            leftBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#FFFFFF'
        });
        this.borderFills.set('6', {
            id: '6',
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#000000'
        });

        const defaultFontIds = {
            hangulId: 0, latinId: 0, hanjaId: 0,
            japaneseId: 0, otherId: 0, symbolId: 0, userId: 0
        };

        const makeCharShape = (id: string, height: string, bold = false, italic = false): CharShape => ({
            id, height, textColor: '#000000', shadeColor: 'none', borderFillIDRef: '1',
            bold, italic, underline: false, strike: false, supscript: false, subscript: false,
            fontIds: { ...defaultFontIds }, fontId: 0
        });
        this.charShapes.set('0', makeCharShape('0', '1000'));
        this.charShapes.set('1', makeCharShape('1', '1000'));
        this.charShapes.set('2', makeCharShape('2', '1000'));
        this.charShapes.set('19', makeCharShape('19', '1000', true));
        this.charShapes.set('24', makeCharShape('24', '1400', true));
        this.charShapes.set('28', makeCharShape('28', '900'));

        this.paraShapes.set('0', {
            id: '0', align: 'LEFT', heading: 'NONE', level: '0', headingIdRef: '0',
            leftMargin: 0, rightMargin: 0, indent: 0,
            prevSpacing: 0, nextSpacing: 0,
            lineSpacingType: 'PERCENT', lineSpacingVal: 115,
            borderFillIDRef: '1', keepWithNext: '0', keepLines: '0',
            pageBreakBefore: '0', tabPrIDRef: '1', tabStops: []
        });

        this.paraShapes.set('1', {
            id: '1', align: 'LEFT', heading: 'NONE', level: '0', headingIdRef: '0',
            leftMargin: 0, rightMargin: 0, indent: 0,
            prevSpacing: 0, nextSpacing: 0,
            lineSpacingType: 'PERCENT', lineSpacingVal: 115,
            borderFillIDRef: '1', keepWithNext: '0', keepLines: '0',
            pageBreakBefore: '0', tabPrIDRef: '1', tabStops: []
        });
    }
}
