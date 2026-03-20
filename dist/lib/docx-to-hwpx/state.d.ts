import { CharShape, ParaShape, BorderFill, StyleData, MetaData, ImageInfo } from './types';

export declare class ConversionState {
    ns: {
        w: string;
        r: string;
        a: string;
        pic: string;
        wp: string;
    };
    hwpxNs: {
        ha: string;
        hp: string;
        hp10: string;
        hs: string;
        hc: string;
        hh: string;
        hhs: string;
        hm: string;
        hpf: string;
        dc: string;
        opf: string;
        ooxmlchart: string;
        epub: string;
        config: string;
        hv: string;
        ocf: string;
        odf: string;
        rdf: string;
        odfPkg: string;
    };
    HWPUNIT_PER_INCH: number;
    EMU_PER_INCH: number;
    PAGE_WIDTH: number;
    PAGE_HEIGHT: number;
    MARGIN_TOP: number;
    MARGIN_LEFT: number;
    MARGIN_RIGHT: number;
    MARGIN_BOTTOM: number;
    HEADER_MARGIN: number;
    FOOTER_MARGIN: number;
    GUTTER_MARGIN: number;
    COL_COUNT: number;
    COL_GAP: number;
    COL_TYPE: string;
    TEXT_WIDTH: number;
    charShapes: Map<string, CharShape>;
    paraShapes: Map<string, ParaShape>;
    borderFills: Map<string, BorderFill>;
    idCounter: number;
    tableIdCounter: number;
    borderFillIdCounter: number;
    imageCounter: number;
    currentVertPos: number;
    footnoteNumber: number;
    endnoteNumber: number;
    metadata: MetaData;
    relationships: Map<string, {
        target: string;
        type: string;
    }>;
    images: Map<string, ImageInfo>;
    langFontFaces: {
        [key: string]: Map<string, string>;
    };
    numberingMap: Map<string, string>;
    abstractNumMap: Map<string, Map<string, any>>;
    listCounters: Map<string, number>;
    docStyles: Map<string, StyleData>;
    docxStyleToHwpxId: {
        [key: string]: number;
    };
    footnoteMap: Map<string, Element>;
    endnoteMap: Map<string, Element>;
    isFirstParagraph: boolean;
    constructor();
    registerFontForLang(lang: string, face: string): number;
    private initDocxStyleToHwpxId;
    private initDefaultStyles;
}
