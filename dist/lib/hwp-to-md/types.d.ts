export interface CharShape {
    isBold: boolean;
    isItalic?: boolean;
    isUnderline?: boolean;
    isStrike?: boolean;
    baseSize?: number;
}
export interface ParaShape {
    alignment: number;
}
export interface Border {
    type: number;
    thickness: number;
    color: string;
}
export interface BorderFill {
    left: Border;
    right: Border;
    top: Border;
    bottom: Border;
    faceColor?: string;
}
