export interface CharShape {
  isBold: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  isStrike?: boolean;
  baseSize?: number; // 폰트 크기 (pt)
}

export interface ParaShape {
  alignment: number; // 0: justify, 1: left, 2: right, 3: center
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
