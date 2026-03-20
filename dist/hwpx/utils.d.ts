export declare function xmlEscape(str: string): string;
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
export declare function ptToHwpUnit(pt: number): number;
export declare function hwpUnitToPt(hwpUnit: number): number;
export declare function mmToHwpUnit(mm: number): number;
export declare function hwpUnitToMm(hwpUnit: number): number;
export declare function cmToHwpUnit(cm: number): number;
export declare function hwpUnitToCm(hwpUnit: number): number;
export declare function inchToHwpUnit(inch: number): number;
export declare function hwpUnitToInch(hwpUnit: number): number;
export declare function pixelToHwpUnit(pixel: number): number;
export declare function hwpUnitToPixel(hwpUnit: number): number;
export declare function charToHwpUnit(char: number): number;
export declare function hwpUnitToChar(hwpUnit: number): number;
export declare function twipsToHwpUnit(twips: number): number;
export declare function hwpUnitToTwips(hwpUnit: number): number;
