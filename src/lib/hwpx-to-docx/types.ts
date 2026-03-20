export interface Border {
    type: string;
    width: string;
    color: string;
}

export interface BorderFill {
    id: string;
    leftBorder: Border;
    rightBorder: Border;
    topBorder: Border;
    bottomBorder: Border;
    faceColor?: string;
}

export interface CharProperty {
    id: string;
    height: number;
    textColor: string;
    shadeColor: string;
    borderFillIDRef: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikeout: boolean;
    supscript: boolean;
    subscript: boolean;
    hangulFont: string;
    latinFont: string;
}

export interface ParaProperty {
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
    keepWithNext: string;
    keepLines: string;
    pageBreakBefore: string;
    borderFillIDRef: string;
}

export interface ImageInfo {
    data: Uint8Array;
    mediaType: string;
    ext: string;
}

export interface ImageRel {
    id: string;
    target: string;
    data: Uint8Array;
    mediaType: string;
    ext: string;
}
