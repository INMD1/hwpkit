import { CharShape } from './types';

/**
 * Parses HWPTAG_PARA_TEXT data, handling inline/extended controls properly.
 * HWP WCHAR is 2 bytes (UTF-16LE).
 */
export declare function parseTextData(data: Uint8Array): string;
export declare function applyCharShapes(textData: Uint8Array, paraCharShapes: {
    pos: number;
    shapeId: number;
}[], charShapes: CharShape[]): string;
