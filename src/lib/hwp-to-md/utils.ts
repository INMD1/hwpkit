import { CharShape } from './types';

/**
 * Parses HWPTAG_PARA_TEXT data, handling inline/extended controls properly.
 * HWP WCHAR is 2 bytes (UTF-16LE).
 */
export function parseTextData(data: Uint8Array): string {
    let text = "";
    let i = 0;
    const words = Math.floor(data.length / 2);

    // Convert Uint8Array to DataView for easier 16-bit reading
    const dataView = new DataView(data.buffer, data.byteOffset, data.byteLength);

    while (i < words) {
        const chCode = dataView.getUint16(i * 2, true); // Little endian

        if (chCode === 10 || chCode === 13) {
            text += '\n';
            i += 1;
        } else if (chCode >= 0 && chCode <= 31) { // Control characters
            if ([0, 10, 13, ...[24, 25, 26, 27, 28, 29, 30, 31]].includes(chCode)) {
                i += 1; // char controls
            } else if ([1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23].includes(chCode)) {
                i += 8; // inline/extended controls are 16 bytes (8 WCHARs)
            } else {
                i += 1;
            }
        } else {
            text += String.fromCharCode(chCode);
            i += 1;
        }
    }
    return text.trim();
}

export function applyCharShapes(textData: Uint8Array, paraCharShapes: { pos: number, shapeId: number }[], charShapes: CharShape[]): string {
    let result = "";

    const dataView = new DataView(textData.buffer, textData.byteOffset, textData.byteLength);
    const words = Math.floor(textData.length / 2);

    let shapeIdx = 0;
    let currentStyle: CharShape = { isBold: false, isItalic: false, isUnderline: false, isStrike: false };

    let i = 0;

    const openTags = (style: CharShape) => {
        let tags = "";
        if (style.baseSize && style.baseSize !== 10) tags += `<span style="font-size: ${style.baseSize}pt;">`;
        if (style.isBold) tags += "<b>";
        if (style.isItalic) tags += "<i>";
        if (style.isUnderline) tags += "<u>";
        if (style.isStrike) tags += "<del>";
        return tags;
    };

    const closeTags = (style: CharShape) => {
        let tags = "";
        if (style.isStrike) tags += "</del>";
        if (style.isUnderline) tags += "</u>";
        if (style.isItalic) tags += "</i>";
        if (style.isBold) tags += "</b>";
        if (style.baseSize && style.baseSize !== 10) tags += "</span>";
        return tags;
    };

    while (i < words) {
        if (shapeIdx < paraCharShapes.length && paraCharShapes[shapeIdx].pos === i) {
            const shapeId = paraCharShapes[shapeIdx].shapeId;
            const newStyle = shapeId < charShapes.length ? charShapes[shapeId] : { isBold: false };

            if (JSON.stringify(currentStyle) !== JSON.stringify(newStyle)) {
                result += closeTags(currentStyle);
                result += openTags(newStyle);
                currentStyle = newStyle;
            }
            shapeIdx++;
        }

        const chCode = dataView.getUint16(i * 2, true);

        let charsToAdd = "";
        let step = 1;

        if (chCode === 10 || chCode === 13) {
            charsToAdd = '\n';
        } else if (chCode >= 0 && chCode <= 31) {
            if ([0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31].includes(chCode)) {
                step = 1;
            } else if ([1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23].includes(chCode)) {
                step = 8;
            } else {
                step = 1;
            }
        } else {
            charsToAdd = String.fromCharCode(chCode);
        }

        if (charsToAdd) {
            result += charsToAdd;
        }

        i += step;
    }

    result += closeTags(currentStyle);

    result = result.replace(/<del><\/del>/g, '');
    result = result.replace(/<u><\/u>/g, '');
    result = result.replace(/<i><\/i>/g, '');
    result = result.replace(/<b><\/b>/g, '');
    result = result.replace(/<span[^>]*><\/span>/g, '');

    return result.trim();
}
