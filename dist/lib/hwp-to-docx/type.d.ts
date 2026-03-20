export interface CharShape {
    fontIds: number[];
    isBold: boolean;
    isItalic: boolean;
    isUnderline: boolean;
    isStrike: boolean;
    baseSize: number;
    charColor: string;
    letterSpacing: number;
}
export interface ParaShape {
    alignment: number;
    leftMargin: number;
    rightMargin: number;
    indent: number;
    lineHeight: number;
    lineHeightType: number;
    spaceAbove: number;
    spaceBelow: number;
    borderFillId: number;
    borderDistLeft: number;
    borderDistRight: number;
    borderDistTop: number;
    borderDistBottom: number;
}
export interface PageDef {
    width: number;
    height: number;
    leftMargin: number;
    rightMargin: number;
    topMargin: number;
    bottomMargin: number;
    headerMargin: number;
    footerMargin: number;
}
export interface TableCell {
    text: string;
    colSpan: number;
    rowSpan: number;
    borderFillId: number;
    width: number;
    height: number;
    isVMergeContinue?: boolean;
}
export interface BorderFill {
    left: any;
    right: any;
    top: any;
    bottom: any;
    faceColor?: string;
}
export interface TextSegment {
    textOffset: number;
    lineY: number;
    lineHeight: number;
    textHeight: number;
    baselineOffset: number;
    lineSpacing: number;
    columnStart: number;
    segmentWidth: number;
    flags: number;
}
