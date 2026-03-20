export interface CharProperty {
    id: string;
    height: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikeout: boolean;
    supscript: boolean;
    subscript: boolean;
}

export interface ParaProperty {
    id: string;
    align: string;
}

export interface StyleEntry {
    id: string;
    type: string;
    level: number;
    paraPrIDRef: string;
    charPrIDRef: string;
}

export interface ImageInfo {
    data: Uint8Array;
    ext: string;
}

export type BorderLine = {
    width: number;
    color: string;
};

export type BorderFill = {
    left?: BorderLine;
    right?: BorderLine;
    top?: BorderLine;
    bottom?: BorderLine;
    backgroundColor?: string;
};
