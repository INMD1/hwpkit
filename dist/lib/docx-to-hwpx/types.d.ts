export interface CharShape {
    id: string;
    height: string;
    textColor: string;
    shadeColor: string;
    borderFillIDRef: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strike: boolean;
    supscript: boolean;
    subscript: boolean;
    fontIds: {
        hangulId: number;
        latinId: number;
        hanjaId: number;
        japaneseId: number;
        otherId: number;
        symbolId: number;
        userId: number;
    };
    fontId: number;
    _key?: string;
}
export interface ParaShape {
    id: string;
    align: string;
    heading: string;
    level: string;
    headingIdRef: string;
    leftMargin: number;
    rightMargin: number;
    indent: number;
    prevSpacing: number;
    nextSpacing: number;
    lineSpacingType: string;
    lineSpacingVal: number;
    tabPrIDRef?: string;
    borderFillIDRef: string;
    keepWithNext: string;
    keepLines: string;
    pageBreakBefore: string;
    tabStops?: any[];
    _key?: string;
}
export interface StyleData {
    align?: string;
    ind?: {
        left?: string;
        right?: string;
        firstLine?: string;
        hanging?: string;
    };
    spacing?: {
        before?: string;
        after?: string;
        line?: string;
        lineRule?: string;
    };
    basedOn?: string;
    rPr?: {
        sz?: string;
        color?: string;
    };
    tblBorders?: any;
    tcBorders?: any;
    shd?: any;
}
export interface BorderDef {
    type: string;
    width: string;
    color: string;
}
export interface BorderFill {
    id: string;
    leftBorder: BorderDef;
    rightBorder: BorderDef;
    topBorder: BorderDef;
    bottomBorder: BorderDef;
    diagonal: BorderDef | null;
    backColor: string;
    slash: BorderDef;
    backSlash: BorderDef;
    _key?: string;
}
export interface MetaData {
    title: string;
    creator: string;
    subject: string;
    description: string;
    keywords?: string;
    lastModifiedBy?: string;
    revision?: string;
    created?: string;
    modified?: string;
}
export interface ImageInfo {
    data: Uint8Array;
    ext: string;
    manifestId: string;
    path: string;
    width: number;
    height: number;
    binDataId: number;
}
