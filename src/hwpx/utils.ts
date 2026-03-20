export function xmlEscape(str: string): string {
    if (!str) return "";
    return str.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * HWPUNIT 변환 유틸리티
 * 10 pt = 1000 HWPUNIT -> 1 pt = 100 HWPUNIT
 * 1 mm = 283.456 HWPUNIT
 * 1 cm = 2834.56 HWPUNIT
 * 1 inch = 7200 HWPUNIT
 * 1 pixel = 75 HWPUNIT
 * 1 char = 500 HWPUNIT
 * 1 twips = 5 HWPUNIT
 */

export function ptToHwpUnit(pt: number): number {
    return Math.round(pt * 100);
}

export function hwpUnitToPt(hwpUnit: number): number {
    return hwpUnit / 100;
}

export function mmToHwpUnit(mm: number): number {
    return Math.round(mm * 283.456);
}

export function hwpUnitToMm(hwpUnit: number): number {
    return hwpUnit / 283.456;
}

export function cmToHwpUnit(cm: number): number {
    return Math.round(cm * 2834.56);
}

export function hwpUnitToCm(hwpUnit: number): number {
    return hwpUnit / 2834.56;
}

export function inchToHwpUnit(inch: number): number {
    return Math.round(inch * 7200);
}

export function hwpUnitToInch(hwpUnit: number): number {
    return hwpUnit / 7200;
}

export function pixelToHwpUnit(pixel: number): number {
    return Math.round(pixel * 75);
}

export function hwpUnitToPixel(hwpUnit: number): number {
    return hwpUnit / 75;
}

export function charToHwpUnit(char: number): number {
    return Math.round(char * 500);
}

export function hwpUnitToChar(hwpUnit: number): number {
    return hwpUnit / 500;
}

export function twipsToHwpUnit(twips: number): number {
    return Math.round(twips * 5);
}

export function hwpUnitToTwips(hwpUnit: number): number {
    return hwpUnit / 5;
}
