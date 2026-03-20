export interface CharShape {
    fontIds: number[];    // 언어별 글꼴 ID [한글, 영문, 한자, 일본어, 기타, 기호, 사용자]
    isBold: boolean;
    isItalic: boolean;
    isUnderline: boolean;
    isStrike: boolean;
    baseSize: number;     // pt 단위
    charColor: string;    // #RRGGBB
    letterSpacing: number;// 자간 (-50% ~ 50%)
}

export interface ParaShape {
    alignment: number;
    leftMargin: number;      // HWPUNIT
    rightMargin: number;     // HWPUNIT
    indent: number;          // HWPUNIT (들여쓰기/내어쓰기)
    lineHeight: number;      // 줄 간격
    lineHeightType: number;  // 줄 간격 종류 (0=글자에 따라, 1=고정값, 2=여백만)
    spaceAbove: number;      // 문단 간격 위 (HWPUNIT)
    spaceBelow: number;      // 문단 간격 아래 (HWPUNIT)
    borderFillId: number;    // 테두리/배경 ID (1-based, 0=없음)
    borderDistLeft: number;  // 문단 테두리 좌 간격 (HWPUNIT)
    borderDistRight: number; // 문단 테두리 우 간격 (HWPUNIT)
    borderDistTop: number;   // 문단 테두리 상 간격 (HWPUNIT)
    borderDistBottom: number;// 문단 테두리 하 간격 (HWPUNIT)
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
    isVMergeContinue?: boolean; // 세로 병합되어 내용이 없는 셀인지 여부
}

export interface BorderFill {
    left: any; right: any; top: any; bottom: any; faceColor?: string;
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