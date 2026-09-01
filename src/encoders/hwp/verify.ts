const HWPTAG_BEGIN = 0x10;

const TAG_ID_MAPPINGS = HWPTAG_BEGIN + 1;
const TAG_BIN_DATA = HWPTAG_BEGIN + 2;
const TAG_FACE_NAME = HWPTAG_BEGIN + 3;
const TAG_BORDER_FILL = HWPTAG_BEGIN + 4;
const TAG_CHAR_SHAPE = HWPTAG_BEGIN + 5;
const TAG_TAB_DEF = HWPTAG_BEGIN + 6;
const TAG_NUMBERING = HWPTAG_BEGIN + 7;
const TAG_BULLET = HWPTAG_BEGIN + 8;
const TAG_PARA_SHAPE = HWPTAG_BEGIN + 9;
const TAG_STYLE = HWPTAG_BEGIN + 10;
const TAG_MEMO_SHAPE = HWPTAG_BEGIN + 76;
const TAG_TRACK_CHANGE = HWPTAG_BEGIN + 80;
const TAG_TRACK_CHANGE_AUTHOR = HWPTAG_BEGIN + 81;

const TAG_PARA_HEADER = HWPTAG_BEGIN + 50;
const TAG_PARA_TEXT = HWPTAG_BEGIN + 51;
const TAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52;
const TAG_PARA_LINE_SEG = HWPTAG_BEGIN + 53;
const TAG_LIST_HEADER = HWPTAG_BEGIN + 56;
const TAG_PAGE_BORDER_FILL = HWPTAG_BEGIN + 59;
const TAG_TABLE = HWPTAG_BEGIN + 61;

const ID_MAPPING_COUNT = 18;
const PARA_HEADER_SIZE = 24;
const PARA_SHAPE_BORDER_FILL_OFFSET = 32;
const CHAR_SHAPE_BORDER_FILL_OFFSET = 68;
const CHAR_SHAPE_RANGE_SIZE = 8;
const LINE_SEG_SIZE = 36;
const CELL_LIST_HEADER_SIZE = 47;
const BULLET_RECORD_MIN_SIZE = 23;

export interface HwpRecord {
  tag: number;
  level: number;
  offset: number;
  data: Uint8Array;
}

export interface HwpVerificationStats {
  charShapeCount: number;
  maxCharShapeId: number;
  paraShapeCount: number;
  styleCount: number;
  borderFillCount: number;
  paragraphCount: number;
}

export interface HwpVerificationResult {
  ok: boolean;
  errors: string[];
  stats: HwpVerificationStats;
}

function readU16(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint16(
    offset,
    true,
  );
}

function readU32(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(
    offset,
    true,
  );
}

function readI32(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getInt32(
    offset,
    true,
  );
}

/** Parse an uncompressed HWP record stream and reject every truncated byte. */
export function parseHwpRecords(data: Uint8Array): HwpRecord[] {
  const records: HwpRecord[] = [];
  let offset = 0;

  while (offset < data.length) {
    const recordOffset = offset;
    if (data.length - offset < 4) {
      throw new Error(
        `offset ${offset}: ${data.length - offset} trailing record byte(s)`,
      );
    }

    const header = readU32(data, offset);
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    offset += 4;

    if (size === 0xfff) {
      if (data.length - offset < 4) {
        throw new Error(`offset ${recordOffset}: missing extended record size`);
      }
      size = readU32(data, offset);
      offset += 4;
    }

    if (size > data.length - offset) {
      throw new Error(
        `offset ${recordOffset}: record size ${size} exceeds ${data.length - offset} remaining byte(s)`,
      );
    }

    records.push({
      tag,
      level,
      offset: recordOffset,
      data: data.subarray(offset, offset + size),
    });
    offset += size;
  }

  return records;
}

function countTag(records: HwpRecord[], tag: number): number {
  return records.reduce(
    (count, record) => count + (record.tag === tag ? 1 : 0),
    0,
  );
}

/** BorderFill references are 1-based; unlike the other HWP IDs, zero is invalid. */
function verifyBorderFillId(
  borderFillId: number,
  borderFillCount: number,
  label: string,
  errors: string[],
): void {
  if (borderFillId < 1 || borderFillId > borderFillCount) {
    errors.push(
      `${label} references borderFillId ${borderFillId}, but valid BorderFill IDs are 1..${borderFillCount}`,
    );
  }
}

function verifyIdMappings(
  records: HwpRecord[],
  errors: string[],
): void {
  const mappings = records.filter((record) => record.tag === TAG_ID_MAPPINGS);
  if (mappings.length !== 1) {
    errors.push(`ID_MAPPINGS record count is ${mappings.length}, expected 1`);
    return;
  }
  if (mappings[0].data.length < ID_MAPPING_COUNT * 4) {
    errors.push(
      `ID_MAPPINGS payload is ${mappings[0].data.length} bytes, expected at least ${ID_MAPPING_COUNT * 4}`,
    );
    return;
  }

  const values = Array.from({ length: ID_MAPPING_COUNT }, (_, index) =>
    readI32(mappings[0].data, index * 4),
  );
  for (let index = 0; index < values.length; index++) {
    if (values[index] < 0) {
      errors.push(`ID_MAPPINGS[${index}] is negative (${values[index]})`);
    }
  }

  // All seven language groups use the same FACE_NAME tag. The record stream
  // cannot identify a group's boundary, so the independently verifiable value
  // is the sum of ID_MAPPINGS[1..7], as in tools/hwp-dump.py.
  const faceNameDeclared = values
    .slice(1, 8)
    .reduce((sum, value) => sum + value, 0);
  const faceNameActual = countTag(records, TAG_FACE_NAME);
  if (faceNameDeclared !== faceNameActual) {
    errors.push(
      `ID_MAPPINGS FACE_NAME total is ${faceNameDeclared}, actual record count is ${faceNameActual}`,
    );
  }

  const mappingIndexToTag: ReadonlyArray<readonly [number, number]> = [
    [0, TAG_BIN_DATA],
    [8, TAG_BORDER_FILL],
    [9, TAG_CHAR_SHAPE],
    [10, TAG_TAB_DEF],
    [11, TAG_NUMBERING],
    [12, TAG_BULLET],
    [13, TAG_PARA_SHAPE],
    [14, TAG_STYLE],
    [15, TAG_MEMO_SHAPE],
    [16, TAG_TRACK_CHANGE],
    [17, TAG_TRACK_CHANGE_AUTHOR],
  ];
  for (const [index, tag] of mappingIndexToTag) {
    const actual = countTag(records, tag);
    if (values[index] !== actual) {
      errors.push(
        `ID_MAPPINGS[${index}] is ${values[index]}, tag ${tag} record count is ${actual}`,
      );
    }
  }
}

function directChildren(records: HwpRecord[], parentIndex: number): HwpRecord[] {
  const parentLevel = records[parentIndex].level;
  const children: HwpRecord[] = [];
  for (let index = parentIndex + 1; index < records.length; index++) {
    const record = records[index];
    if (record.level <= parentLevel) break;
    if (record.level === parentLevel + 1) children.push(record);
  }
  return children;
}

/**
 * Validate every cross-record reference in raw, uncompressed DocInfo/Section0.
 * The encoder calls this before deflate so it never emits a silently broken HWP.
 */
export function verifyHwpRecordStreams(
  docInfoRaw: Uint8Array,
  sectionRaw: Uint8Array,
): HwpVerificationResult {
  const errors: string[] = [];
  let docInfoRecords: HwpRecord[] = [];
  let bodyRecords: HwpRecord[] = [];

  try {
    docInfoRecords = parseHwpRecords(docInfoRaw);
  } catch (error) {
    errors.push(
      `DocInfo record framing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  try {
    bodyRecords = parseHwpRecords(sectionRaw);
  } catch (error) {
    errors.push(
      `Section0 record framing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const charShapeCount = countTag(docInfoRecords, TAG_CHAR_SHAPE);
  const binDataCount = countTag(docInfoRecords, TAG_BIN_DATA);
  const paraShapeCount = countTag(docInfoRecords, TAG_PARA_SHAPE);
  const styleCount = countTag(docInfoRecords, TAG_STYLE);
  const borderFillCount = countTag(docInfoRecords, TAG_BORDER_FILL);
  let maxCharShapeId = -1;
  let paragraphCount = 0;

  if (docInfoRecords.length > 0) verifyIdMappings(docInfoRecords, errors);

  for (const record of docInfoRecords) {
    if (record.tag === TAG_PARA_SHAPE) {
      if (record.data.length < PARA_SHAPE_BORDER_FILL_OFFSET + 2) {
        errors.push(
          `PARA_SHAPE at offset ${record.offset} is ${record.data.length} bytes, expected at least ${PARA_SHAPE_BORDER_FILL_OFFSET + 2}`,
        );
        continue;
      }
      verifyBorderFillId(
        readU16(record.data, PARA_SHAPE_BORDER_FILL_OFFSET),
        borderFillCount,
        `PARA_SHAPE at offset ${record.offset}`,
        errors,
      );
    } else if (record.tag === TAG_CHAR_SHAPE) {
      if (record.data.length < CHAR_SHAPE_BORDER_FILL_OFFSET + 2) {
        errors.push(
          `CHAR_SHAPE at offset ${record.offset} is ${record.data.length} bytes, expected at least ${CHAR_SHAPE_BORDER_FILL_OFFSET + 2}`,
        );
        continue;
      }
      verifyBorderFillId(
        readU16(record.data, CHAR_SHAPE_BORDER_FILL_OFFSET),
        borderFillCount,
        `CHAR_SHAPE at offset ${record.offset}`,
        errors,
      );
    } else if (record.tag === TAG_BULLET) {
      if (record.data.length < BULLET_RECORD_MIN_SIZE) {
        errors.push(
          `BULLET at offset ${record.offset} is ${record.data.length} bytes, expected at least ${BULLET_RECORD_MIN_SIZE}`,
        );
        continue;
      }
      const charShapeId = readU32(record.data, 8);
      if (charShapeId !== 0xffffffff && charShapeId >= charShapeCount) {
        errors.push(
          `BULLET at offset ${record.offset} references charShapeId ${charShapeId}, but only ${charShapeCount} CHAR_SHAPE record(s) exist`,
        );
      }
      const imageBinDataId = readU16(record.data, 21);
      if (imageBinDataId > binDataCount) {
        errors.push(
          `BULLET at offset ${record.offset} references BinData ID ${imageBinDataId}, but only ${binDataCount} BIN_DATA record(s) exist`,
        );
      }
    }
  }

  for (let index = 0; index < bodyRecords.length; index++) {
    const header = bodyRecords[index];
    if (header.tag !== TAG_PARA_HEADER) continue;
    paragraphCount++;

    if (header.data.length < PARA_HEADER_SIZE) {
      errors.push(
        `PARA_HEADER at offset ${header.offset} is ${header.data.length} bytes, expected at least ${PARA_HEADER_SIZE}`,
      );
      continue;
    }

    const rawCharCount = readU32(header.data, 0);
    const nChars = rawCharCount & 0x7fffffff;
    const paraShapeId = readU16(header.data, 8);
    const styleId = header.data[10];
    const csCount = readU16(header.data, 12);
    const lineAlignCount = readU16(header.data, 16);
    const label = `PARA_HEADER at offset ${header.offset}`;

    if ((rawCharCount & 0x80000000) === 0) {
      errors.push(`${label} has no 0x80000000 character-count bit`);
    }
    if (paraShapeId >= paraShapeCount) {
      errors.push(
        `${label} references paraShapeId ${paraShapeId}, but only ${paraShapeCount} PARA_SHAPE record(s) exist`,
      );
    }
    if (styleId >= styleCount) {
      errors.push(
        `${label} references styleId ${styleId}, but only ${styleCount} STYLE record(s) exist`,
      );
    }

    const children = directChildren(bodyRecords, index);
    const paraTexts = children.filter((record) => record.tag === TAG_PARA_TEXT);
    const charShapeRanges = children.filter(
      (record) => record.tag === TAG_PARA_CHAR_SHAPE,
    );
    const lineSegs = children.filter(
      (record) => record.tag === TAG_PARA_LINE_SEG,
    );

    let textBytes = 0;
    for (const record of paraTexts) {
      if (record.data.length % 2 !== 0) {
        errors.push(
          `${label} has an odd-sized PARA_TEXT payload (${record.data.length} bytes)`,
        );
      }
      textBytes += record.data.length;
    }
    if (nChars !== textBytes / 2) {
      errors.push(
        `${label} declares nChars ${nChars}, but direct PARA_TEXT payloads contain ${textBytes / 2} UTF-16 code unit(s)`,
      );
    }

    let actualCsCount = 0;
    for (const record of charShapeRanges) {
      if (record.data.length % CHAR_SHAPE_RANGE_SIZE !== 0) {
        errors.push(
          `${label} has a PARA_CHAR_SHAPE payload of ${record.data.length} bytes, not a multiple of ${CHAR_SHAPE_RANGE_SIZE}`,
        );
        continue;
      }
      actualCsCount += record.data.length / CHAR_SHAPE_RANGE_SIZE;
      for (
        let offset = 0;
        offset < record.data.length;
        offset += CHAR_SHAPE_RANGE_SIZE
      ) {
        const charShapeId = readU32(record.data, offset + 4);
        maxCharShapeId = Math.max(maxCharShapeId, charShapeId);
        if (charShapeId >= charShapeCount) {
          errors.push(
            `${label} references charShapeId ${charShapeId}, but only ${charShapeCount} CHAR_SHAPE record(s) exist`,
          );
        }
      }
    }
    if (csCount !== actualCsCount) {
      errors.push(
        `${label} declares csCount ${csCount}, but has ${actualCsCount} PARA_CHAR_SHAPE range(s)`,
      );
    }

    let actualLineAlignCount = 0;
    for (const record of lineSegs) {
      if (record.data.length % LINE_SEG_SIZE !== 0) {
        errors.push(
          `${label} has a PARA_LINE_SEG payload of ${record.data.length} bytes, not a multiple of ${LINE_SEG_SIZE}`,
        );
        continue;
      }
      actualLineAlignCount += record.data.length / LINE_SEG_SIZE;
    }
    if (lineAlignCount !== actualLineAlignCount) {
      errors.push(
        `${label} declares lineAlignCount ${lineAlignCount}, but has ${actualLineAlignCount} LINE_SEG item(s)`,
      );
    }
  }

  for (const record of bodyRecords) {
    let borderFillId: number | undefined;
    if (record.tag === TAG_PAGE_BORDER_FILL) {
      if (record.data.length < 14) {
        errors.push(
          `PAGE_BORDER_FILL at offset ${record.offset} is ${record.data.length} bytes, expected at least 14`,
        );
        continue;
      }
      borderFillId = readU16(record.data, 12);
    } else if (
      record.tag === TAG_LIST_HEADER &&
      record.data.length === CELL_LIST_HEADER_SIZE
    ) {
      borderFillId = readU16(record.data, 32);
    } else if (record.tag === TAG_TABLE) {
      if (record.data.length < 6) {
        errors.push(
          `TABLE at offset ${record.offset} is ${record.data.length} bytes, expected at least 6`,
        );
        continue;
      }
      const rowCount = readU16(record.data, 4);
      const borderFillOffset = 18 + rowCount * 2;
      if (record.data.length < borderFillOffset + 2) {
        errors.push(
          `TABLE at offset ${record.offset} is ${record.data.length} bytes, expected borderFillId at offset ${borderFillOffset}`,
        );
        continue;
      }
      borderFillId = readU16(record.data, borderFillOffset);
    }

    if (borderFillId !== undefined) {
      verifyBorderFillId(
        borderFillId,
        borderFillCount,
        `record tag ${record.tag} at offset ${record.offset}`,
        errors,
      );
    }
  }

  return {
    ok: errors.length === 0,
    errors,
    stats: {
      charShapeCount,
      maxCharShapeId,
      paraShapeCount,
      styleCount,
      borderFillCount,
      paragraphCount,
    },
  };
}
