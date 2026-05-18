// src/contract/result.ts
function succeed(data, warns = []) {
  return { ok: true, data, warns };
}
function fail(error, warns = []) {
  return { ok: false, error, warns };
}

// src/pipeline/registry.ts
var FormatRegistry = class {
  constructor() {
    this.decoders = /* @__PURE__ */ new Map();
    this.encoders = /* @__PURE__ */ new Map();
  }
  registerDecoder(d) {
    this.decoders.set(d.format, d);
    if (d.aliases) {
      for (const alias of d.aliases) this.decoders.set(alias, d);
    }
  }
  registerEncoder(e) {
    this.encoders.set(e.format, e);
    if (e.aliases) {
      for (const alias of e.aliases) this.encoders.set(alias, e);
    }
  }
  getDecoder(fmt) {
    return this.decoders.get(fmt);
  }
  getEncoder(fmt) {
    return this.encoders.get(fmt);
  }
  supportedInputs() {
    return [...this.decoders.keys()];
  }
  supportedOutputs() {
    return [...this.encoders.keys()];
  }
};
var registry = new FormatRegistry();

// src/model/doc-props.ts
var A4 = {
  wPt: 595.28,
  hPt: 841.89,
  mt: 56.69,
  mb: 56.69,
  ml: 70.87,
  mr: 70.87,
  orient: "portrait"
};
var A4_LANDSCAPE = {
  wPt: 841.89,
  hPt: 595.28,
  mt: 56.69,
  mb: 56.69,
  ml: 70.87,
  mr: 70.87,
  orient: "landscape"
};
function normalizeDims(dims) {
  const orient = dims.orient ?? "portrait";
  if (orient === "landscape" && dims.wPt < dims.hPt) {
    return { ...dims, wPt: dims.hPt, hPt: dims.wPt };
  }
  if (orient === "portrait" && dims.wPt > dims.hPt) {
    return { ...dims, wPt: dims.hPt, hPt: dims.wPt };
  }
  return dims;
}
var DEFAULT_STROKE = { kind: "solid", pt: 0.5, color: "000000" };

// src/model/builders.ts
function buildRoot(meta = {}, kids = []) {
  return { tag: "root", meta, kids };
}
function buildSheet(kids = [], dims = A4, opts) {
  const node = { tag: "sheet", dims, kids };
  if (opts?.headers) node.headers = opts.headers;
  if (opts?.footers) node.footers = opts.footers;
  return node;
}
function buildPageNum(format) {
  return { tag: "pagenum", format };
}
function buildBr() {
  return { tag: "br" };
}
function buildPb() {
  return { tag: "pb" };
}
function buildPara(kids = [], props = {}) {
  return { tag: "para", props, kids };
}
function buildSpan(content, props = {}) {
  const txt = { tag: "txt", content };
  return { tag: "span", props, kids: [txt] };
}
function buildImg(b64, mime, w, h, alt, layout) {
  const node = { tag: "img", b64, mime, w, h };
  if (alt) node.alt = alt;
  if (layout) node.layout = layout;
  return node;
}
function buildGrid(kids, props = {}) {
  return { tag: "grid", props, kids };
}
function buildRow(kids, heightPt) {
  const node = { tag: "row", kids };
  if (heightPt != null) node.heightPt = heightPt;
  return node;
}
function buildCell(kids, opts = {}) {
  return { tag: "cell", cs: opts.cs ?? 1, rs: opts.rs ?? 1, props: opts.props ?? {}, kids };
}

// src/safety/ShieldedParser.ts
var ShieldedParser = class {
  constructor() {
    this.log = [];
  }
  /** 단일 요소 안전 파싱 */
  guard(fn, fallback, label) {
    try {
      const v = fn();
      if (v == null) {
        this.warn(label, "returned null/undefined");
        return fallback;
      }
      return v;
    } catch (e) {
      this.warn(label, e?.message ?? String(e));
      return fallback;
    }
  }
  /** 배열 각 요소 독립 파싱 (하나 실패해도 나머지 계속) */
  guardAll(items, fn, fb, label) {
    return items.map(
      (x, i) => this.guard(() => fn(x, i), fb(x, i), `${label}[${i}]`)
    );
  }
  /**
   * 표 전용 4단계 폴백
   *   Lv1: Full → Lv2: Grid → Lv3: Flat → Lv4: Text
   */
  guardGrid(node, lv1Full, lv2Grid, lv3Flat, lv4Text, label) {
    const levels = [
      [lv1Full, 1],
      [lv2Grid, 2],
      [lv3Flat, 3],
      [lv4Text, 4]
    ];
    for (const [fn, lv] of levels) {
      try {
        const v = fn(node);
        if (v != null) {
          if (lv > 1) this.warn(label, `degraded to level ${lv}`);
          return { value: v, level: lv };
        }
      } catch (e) {
        this.warn(label, `Lv${lv} failed: ${e?.message ?? String(e)}`);
      }
    }
    this.warn(label, "ALL LEVELS FAILED \u2014 returning lv4Text forced");
    return { value: lv4Text(null), level: 4 };
  }
  /** 이미지 안전 파싱 */
  guardImg(node, fn, placeholder, label) {
    try {
      const v = fn(node);
      if (v != null) return v;
    } catch (e) {
      this.warn(label, e?.message ?? String(e));
    }
    this.warn(label, "using placeholder image");
    return placeholder(`[\uC774\uBBF8\uC9C0 \uB85C\uB4DC \uC2E4\uD328: ${label}]`);
  }
  warn(label, msg) {
    const w = `[SHIELD] ${label}: ${msg}`;
    console.warn(w);
    this.log.push(w);
  }
  flush() {
    const r = [...this.log];
    this.log = [];
    return r;
  }
};

// src/safety/StyleBridge.ts
var Metric = {
  // HWP 세계 (1 inch = 7200 HWPUNIT)
  hwpToPt: (v) => v / 100,
  ptToHwp: (v) => Math.round(v * 100),
  hwpToDxa: (v) => Math.round(v / 5),
  dxaToHwp: (v) => Math.round(v * 5),
  hwpToEmu: (v) => Math.round(v * 127),
  emuToHwp: (v) => Math.round(v / 127),
  // DOCX 세계 (1 inch = 1440 dxa, 1 pt = 20 dxa)
  dxaToPt: (v) => v / 20,
  ptToDxa: (v) => Math.round(v * 20),
  dxaToEmu: (v) => Math.round(v * 635),
  emuToDxa: (v) => Math.round(v / 635),
  emuToPt: (v) => v / 12700,
  ptToEmu: (v) => Math.round(v * 12700),
  // HWPX charPr height: 1000 = 10pt
  hHeightToPt: (v) => v / 100,
  ptToHHeight: (v) => Math.round(v * 100),
  // DOCX half-point: 24 = 12pt
  halfPtToPt: (v) => v / 2,
  ptToHalfPt: (v) => Math.round(v * 2)
};
function safeHex(raw) {
  if (raw == null) return void 0;
  if (typeof raw === "number") {
    if (raw <= 0) return "000000";
    if (raw >= 16777215) return void 0;
    return raw.toString(16).padStart(6, "0").toUpperCase();
  }
  let s = String(raw).replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{3}$/.test(s)) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
  if (/^[0-9A-F]{6}$/.test(s)) return s;
  if (s === "AUTO" || s === "NONE" || s === "TRANSPARENT") return void 0;
  return void 0;
}
var ALIGN_MAP = {
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
  JUSTIFY: "justify",
  BOTH: "justify",
  DISTRIBUTE: "justify",
  left: "left",
  center: "center",
  right: "right",
  both: "justify",
  start: "left",
  end: "right"
};
function safeAlign(raw) {
  return ALIGN_MAP[raw ?? ""] ?? "left";
}
var HWPX_STROKE = {
  SOLID: "solid",
  NONE: "none",
  DASH: "dash",
  DOT: "dot",
  DOUBLE: "double",
  LONG_DASH: "dash",
  DASH_DOT: "dashDot",
  DASH_DOT_DOT: "dashDotDot",
  THICK_THIN: "double",
  THIN_THICK: "double",
  TRIPLE: "double"
};
var DOCX_STROKE = {
  single: "solid",
  none: "none",
  nil: "none",
  dashed: "dash",
  dotted: "dot",
  double: "double",
  dotDash: "dashDot",
  dotDotDash: "dashDotDot",
  thickThin: "double",
  thinThick: "double",
  triple: "double",
  wave: "wave",
  dashDotStroked: "dashDot",
  threeDEmboss: "solid",
  threeDEngrave: "solid"
};
function safeStrokeHwpx(type, w, c) {
  return {
    kind: HWPX_STROKE[type ?? ""] ?? "solid",
    pt: w != null ? Metric.hwpToPt(w) : 0.5,
    color: safeHex(c) ?? "000000"
  };
}
function safeStrokeDocx(val, sz, c) {
  return {
    kind: DOCX_STROKE[val ?? ""] ?? "solid",
    pt: sz != null ? sz / 8 : 0.5,
    color: safeHex(c) ?? "000000"
  };
}
var FONT_MAP = {
  // 맑은 고딕 계열
  "\uB9D1\uC740 \uACE0\uB515": "Malgun Gothic",
  "\uB9D1\uC740\uACE0\uB515": "Malgun Gothic",
  // 바탕 계열 (serif)
  "\uBC14\uD0D5": "Batang",
  "\uBC14\uD0D5\uCCB4": "BatangChe",
  "\uD55C\uCEF4\uBC14\uD0D5": "Batang",
  "\uD568\uCD08\uB86C\uBC14\uD0D5": "Batang",
  "HY\uC2E0\uBA85\uC870": "Batang",
  "HY\uACAC\uBA85\uC870": "Batang",
  "HY\uADF8\uB798\uD53D": "Batang",
  "\uAD81\uC11C": "Gungsuh",
  "\uAD81\uC11C\uCCB4": "GungsuhChe",
  // 고딕 계열 (sans-serif)
  "\uB3CB\uC6C0": "Dotum",
  "\uB3CB\uC6C0\uCCB4": "DotumChe",
  "\uAD74\uB9BC": "Gulim",
  "\uAD74\uB9BC\uCCB4": "GulimChe",
  "\uD55C\uCEF4\uB3CB\uC6C0": "Malgun Gothic",
  "\uD568\uCD08\uB86C\uB3CB\uC6C0": "Malgun Gothic",
  "HY\uACAC\uACE0\uB515": "Malgun Gothic",
  "HY\uC911\uACE0\uB515": "Malgun Gothic",
  "HY\uD5E4\uB4DC\uB77C\uC778M": "Malgun Gothic",
  "HY\uAC15B": "Malgun Gothic",
  "HY\uB098\uBB34M": "Malgun Gothic",
  "HY\uBAA9\uAC01\uD30C\uC784B": "Malgun Gothic",
  "HY\uC5FD\uC11CM": "Malgun Gothic",
  "HY\uC5FD\uC11CL": "Malgun Gothic",
  // 나눔 계열
  "\uB098\uB214\uACE0\uB515": "Malgun Gothic",
  "\uB098\uB214\uBA85\uC870": "Batang"
};
function safeFont(raw) {
  return FONT_MAP[raw ?? ""] ?? raw ?? "Malgun Gothic";
}
var FONT_MAP_KR = {
  "Malgun Gothic": "\uB9D1\uC740 \uACE0\uB515",
  "Batang": "\uBC14\uD0D5",
  "Dotum": "\uB3CB\uC6C0",
  "Gulim": "\uAD74\uB9BC"
};
function safeFontToKr(raw) {
  return FONT_MAP_KR[raw ?? ""] ?? raw ?? "\uB9D1\uC740 \uACE0\uB515";
}

// src/toolkit/ArchiveKit.ts
import pako from "pako";
var ArchiveKit = {
  async inflate(compressed) {
    return pako.inflate(compressed);
  },
  async deflate(data) {
    return pako.deflate(data, { level: 6 });
  },
  async unzip(zipData) {
    const files = /* @__PURE__ */ new Map();
    const view = new DataView(zipData.buffer, zipData.byteOffset, zipData.byteLength);
    let eocdOffset = -1;
    const searchStart = Math.max(0, zipData.length - 65558);
    for (let i = zipData.length - 22; i >= searchStart; i--) {
      if (view.getUint32(i, true) === 101010256) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset !== -1) {
      const entryCount = view.getUint16(eocdOffset + 10, true);
      const centralDirOffset = view.getUint32(eocdOffset + 16, true);
      let cdOffset = centralDirOffset;
      for (let i = 0; i < entryCount; i++) {
        if (cdOffset + 46 > zipData.length) break;
        if (view.getUint32(cdOffset, true) !== 33639248) break;
        const compressionMethod = view.getUint16(cdOffset + 10, true);
        const compressedSize = view.getUint32(cdOffset + 20, true);
        const uncompressedSize = view.getUint32(cdOffset + 24, true);
        const fileNameLength = view.getUint16(cdOffset + 28, true);
        const extraLength = view.getUint16(cdOffset + 30, true);
        const commentLength = view.getUint16(cdOffset + 32, true);
        const localHeaderOffset = view.getUint32(cdOffset + 42, true);
        const nameBytes = zipData.subarray(cdOffset + 46, cdOffset + 46 + fileNameLength);
        const name = new TextDecoder("utf-8").decode(nameBytes);
        cdOffset += 46 + fileNameLength + extraLength + commentLength;
        if (name.endsWith("/")) continue;
        const localFnLen = view.getUint16(localHeaderOffset + 26, true);
        const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
        const dataOffset = localHeaderOffset + 30 + localFnLen + localExtraLen;
        let fileData;
        if (compressionMethod === 0) {
          fileData = zipData.subarray(dataOffset, dataOffset + uncompressedSize);
        } else if (compressionMethod === 8) {
          const compressed = zipData.subarray(dataOffset, dataOffset + compressedSize);
          fileData = pako.inflateRaw(compressed);
        } else {
          throw new Error(`Unsupported ZIP compression method: ${compressionMethod}`);
        }
        files.set(name, new Uint8Array(fileData));
      }
      return files;
    }
    let offset = 0;
    while (offset < zipData.length - 4) {
      const sig = view.getUint32(offset, true);
      if (sig === 67324752) {
        const compressionMethod = view.getUint16(offset + 8, true);
        const compressedSize = view.getUint32(offset + 18, true);
        const uncompressedSize = view.getUint32(offset + 22, true);
        const fileNameLength = view.getUint16(offset + 26, true);
        const extraLength = view.getUint16(offset + 28, true);
        const nameBytes = zipData.subarray(offset + 30, offset + 30 + fileNameLength);
        const name = new TextDecoder("utf-8").decode(nameBytes);
        const dataOffset = offset + 30 + fileNameLength + extraLength;
        let fileData;
        if (compressionMethod === 0) {
          fileData = zipData.subarray(dataOffset, dataOffset + uncompressedSize);
        } else if (compressionMethod === 8) {
          const compressed = zipData.subarray(dataOffset, dataOffset + compressedSize);
          fileData = pako.inflateRaw(compressed);
        } else {
          throw new Error(`Unsupported ZIP compression method: ${compressionMethod}`);
        }
        files.set(name, new Uint8Array(fileData));
        offset = dataOffset + compressedSize;
      } else if (sig === 33639248 || sig === 101010256) {
        break;
      } else {
        offset++;
      }
    }
    return files;
  },
  async zip(entries) {
    const localHeaders = [];
    const centralHeaders = [];
    let localOffset = 0;
    for (const entry of entries) {
      const nameBytes = new TextEncoder().encode(entry.name);
      const crc = crc32(entry.data);
      const store = entry.name === "mimetype" || entry.name === "version.xml";
      const method = store ? 0 : 8;
      const payload = store ? entry.data : pako.deflateRaw(entry.data, { level: 6 });
      const local = new Uint8Array(30 + nameBytes.length + payload.length);
      const lv = new DataView(local.buffer);
      lv.setUint32(0, 67324752, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 33, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, payload.length, true);
      lv.setUint32(22, entry.data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      local.set(nameBytes, 30);
      local.set(payload, 30 + nameBytes.length);
      const central = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(central.buffer);
      cv.setUint32(0, 33639248, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, method, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 33, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, payload.length, true);
      cv.setUint32(24, entry.data.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, localOffset, true);
      central.set(nameBytes, 46);
      localHeaders.push(local);
      centralHeaders.push(central);
      localOffset += local.length;
    }
    const centralDir = concat(centralHeaders);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 101010256, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralDir.length, true);
    ev.setUint32(16, localOffset, true);
    ev.setUint16(20, 0, true);
    return concat([...localHeaders, centralDir, eocd]);
  }
};
function concat(arrays) {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    out.set(a, offset);
    offset += a.length;
  }
  return out;
}
function crc32(data) {
  let crc = 4294967295;
  for (const byte of data) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) {
      crc = crc & 1 ? crc >>> 1 ^ 3988292384 : crc >>> 1;
    }
  }
  return (crc ^ 4294967295) >>> 0;
}

// src/toolkit/XmlKit.ts
import { SaxesParser } from "saxes";
function parseXmlStrict(xml) {
  return new Promise((resolve, reject) => {
    const parser = new SaxesParser({ xmlns: false });
    const stack = [];
    let result = null;
    parser.on("error", (err) => reject(err));
    parser.on("opentag", (node) => {
      const obj = {};
      const attrs = node.attributes;
      if (attrs && Object.keys(attrs).length > 0) {
        obj["_attr"] = { ...attrs };
      }
      stack.push({ tag: node.name, obj });
    });
    const appendText = (text) => {
      if (stack.length > 0 && text) {
        const frame = stack[stack.length - 1];
        const cur = frame.obj["_text"];
        frame.obj["_text"] = typeof cur === "string" ? cur + text : text;
      }
    };
    parser.on("text", (text) => appendText(text));
    parser.on("cdata", (cdata) => appendText(cdata));
    parser.on("closetag", () => {
      const frame = stack.pop();
      if (!frame) return;
      const { tag, obj } = frame;
      if (stack.length === 0) {
        result = { [tag]: [obj] };
      } else {
        const parent = stack[stack.length - 1].obj;
        const existing = parent[tag];
        if (Array.isArray(existing)) {
          existing.push(obj);
        } else {
          parent[tag] = [obj];
        }
        if (!parent["_childOrder"]) parent["_childOrder"] = [];
        parent["_childOrder"].push(tag);
      }
    });
    try {
      parser.write(xml).close();
      resolve(result);
    } catch (e) {
      reject(e);
    }
  });
}
var XmlKit = {
  /** @deprecated Use parseStrict instead */
  async parse(xml) {
    return parseXmlStrict(xml);
  },
  async parseStrict(xml) {
    return parseXmlStrict(xml);
  },
  attr(node, key) {
    const a = node["_attr"];
    return a?.[key];
  },
  text(node) {
    if (node == null) return "";
    if (typeof node === "string") return node;
    const t = node["_text"];
    return typeof t === "string" ? t : "";
  }
};

// src/toolkit/TextKit.ts
var TextKit = {
  decode(data, encoding = "utf-8") {
    try {
      return new TextDecoder(encoding, { fatal: true }).decode(data);
    } catch {
      return new TextDecoder("utf-8", { fatal: false }).decode(data);
    }
  },
  encode(text) {
    return new TextEncoder().encode(text);
  },
  escapeXml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  },
  unescapeXml(s) {
    return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  },
  normalizeWhitespace(s) {
    return s.replace(/\s+/g, " ").trim();
  },
  stripControl(s) {
    return s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  },
  base64Encode(data) {
    let binary = "";
    for (let i = 0; i < data.length; i++) {
      binary += String.fromCharCode(data[i]);
    }
    return btoa(binary);
  },
  base64Decode(b64) {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }
};

// src/core/BaseDecoder.ts
var BaseDecoder = class {
  constructor() {
    this.format = this.getFormat();
    this.aliases = this.getAliases();
  }
  /** 별칭 목록 반환 (하위 클래스에서 필요 시 오버라이드) */
  getAliases() {
    return [];
  }
  // ─── 공통 유틸리티 메서드 ──────────────────────────
  /** 바이트를 UTF-8 문자열로 변환 */
  bytesToString(data) {
    return TextKit.decode(data);
  }
  /** 문자열을 UTF-8 바이트로 변환 */
  stringToBytes(s) {
    return TextKit.encode(s);
  }
  /** XML 이스케이프 */
  escapeXml(s) {
    return TextKit.escapeXml(s);
  }
  /** XML 언이스케이프 */
  unescapeXml(s) {
    return TextKit.unescapeXml(s);
  }
  /** base64 문자열을 Uint8Array 로 변환 */
  base64ToBytes(b64) {
    return TextKit.base64Decode(b64);
  }
  /** Uint8Array 를 base64 문자열로 변환 */
  bytesToBase64(data) {
    return TextKit.base64Encode(data);
  }
  /** ZIP 해제 */
  async unzip(data) {
    return ArchiveKit.unzip(data);
  }
  /** ZIP 압축 */
  async zip(entries) {
    return ArchiveKit.zip(entries);
  }
  /** inflate 해제 */
  async inflate(data) {
    return ArchiveKit.inflate(data);
  }
  /** deflate 압축 */
  async deflate(data) {
    return ArchiveKit.deflate(data);
  }
  /** 제어 문자 제거 */
  stripControl(s) {
    return TextKit.stripControl(s);
  }
  /** 공백 정규화 */
  normalizeWhitespace(s) {
    return TextKit.normalizeWhitespace(s);
  }
  /** XML 파싱 (DOMParser 사용) */
  parseXml(xmlString) {
    const parser = new DOMParser();
    return parser.parseFromString(xmlString, "text/xml");
  }
  /** XML 요소에서 텍스트 내용 추출 */
  getTextContent(element) {
    if (!element) return "";
    return element.textContent ?? "";
  }
  /** XML 요소의 속성 값 추출 */
  getAttr(element, name) {
    return element?.getAttribute(name) ?? null;
  }
  /** XML 요소의 자식 요소 찾기 */
  getChild(element, tagName) {
    if (!element) return null;
    return element.querySelector(`>${tagName}`) ?? null;
  }
  /** XML 요소의 모든 자식 요소 찾기 */
  getChildren(element, tagName) {
    if (!element) return [];
    return Array.from(element.querySelectorAll(`>${tagName}`));
  }
};

// src/encoders/hwpx/constants.ts
var HWPX_MIME_TYPE = "application/hwp+zip";
var NAMESPACES = {
  /** Hancom 문서 네임스페이스 */
  HANCOM: "http://www.hancom.co.kr/hwp/xml",
  /** Hancom 공통 네임스페이스 */
  HANCOM_COMMON: "http://www.hancom.co.kr/hwp/xml/common",
  /** Hancom 버전 네임스페이스 */
  HANCOM_VERSION: "http://www.hancom.co.kr/hwp/xml/version",
  /** Hancom 속성 네임스페이스 */
  HANCOM_PROP: "http://www.hancom.co.kr/hwp/xml/property"
};
var NAMESPACE_DECLARATIONS = {
  HEAD: `xmlns:hh="${NAMESPACES.HANCOM}" xmlns:hc="${NAMESPACES.HANCOM_COMMON}" xmlns:hv="${NAMESPACES.HANCOM_VERSION}" xmlns:hp="${NAMESPACES.HANCOM_PROP}"`,
  SECTION: `xmlns:hs="${NAMESPACES.HANCOM}" xmlns:hp="${NAMESPACES.HANCOM_PROP}"`
};
var PT_PER_INCH = 72;
var PIXELS_PER_INCH = 96;
var PT_PER_PIXEL = PT_PER_INCH / PIXELS_PER_INCH;

// src/decoders/hwpx/HwpxDecoder.ts
var HwpxDecoder = class extends BaseDecoder {
  getFormat() {
    return "hwpx";
  }
  getAliases() {
    return [HWPX_MIME_TYPE, "application/hwp+zip"];
  }
  async decode(data) {
    const shield = new ShieldedParser();
    const warns = [];
    try {
      const files = await ArchiveKit.unzip(data);
      const sectionFiles = [];
      for (let i = 0; ; i++) {
        const sec = files.get(`Contents/section${i}.xml`) ?? files.get(`section${i}.xml`);
        if (!sec) break;
        sectionFiles.push(sec);
      }
      if (sectionFiles.length === 0) {
        const fallback = findSectionFile(files);
        if (fallback) sectionFiles.push(fallback);
      }
      if (sectionFiles.length === 0)
        return fail("HWPX: No section files found");
      const headXml = files.get("Contents/header.xml") ?? files.get("header.xml");
      let meta = {};
      let dims = { ...A4 };
      let borderFills = /* @__PURE__ */ new Map();
      let charPrs = /* @__PURE__ */ new Map();
      let paraPrs = /* @__PURE__ */ new Map();
      if (headXml) {
        try {
          const headStr = TextKit.decode(headXml);
          const headObj = await XmlKit.parseStrict(headStr);
          if (headObj) {
            meta = extractMeta(headObj);
            dims = extractDims(headObj) ?? dims;
            borderFills = extractBorderFills(headObj);
            charPrs = extractCharPrs(headObj);
            paraPrs = extractParaPrs(headObj);
          }
        } catch {
        }
      }
      const ctx = {
        files,
        shield,
        borderFills,
        charPrs,
        paraPrs,
        warns
      };
      const allSections = [];
      for (const secFile of sectionFiles) {
        const bodyStr = TextKit.decode(secFile);
        const bodyObj = await XmlKit.parseStrict(bodyStr);
        allSections.push(...normalizeSections(bodyObj));
      }
      const kids = shield.guardAll(
        allSections,
        (sec) => decodeSection(sec, dims, ctx),
        () => buildSheet([buildPara([buildSpan("[\uC139\uC158 \uD30C\uC2F1 \uC2E4\uD328]")])], dims),
        "hwpx:section"
      );
      warns.push(...shield.flush());
      return succeed(buildRoot(meta, kids), warns);
    } catch (e) {
      warns.push(...shield.flush());
      return fail(`HWPX decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
};
function findSectionFile(files) {
  for (const [key, val] of files) {
    if (key.toLowerCase().includes("section") && key.endsWith(".xml"))
      return val;
  }
  return void 0;
}
function normalizeSections(bodyObj) {
  if (bodyObj?.["hs:sec"]) return toArr(bodyObj["hs:sec"]);
  if (bodyObj?.["hp:SEC"]) return toArr(bodyObj["hp:SEC"]);
  const root = bodyObj?.["hp:HWPML"] ?? bodyObj?.HWPML ?? bodyObj;
  const body = root?.["hp:BODY"]?.[0] ?? root?.BODY?.[0] ?? root?.["hp:BODY"] ?? root?.BODY;
  if (!body) return [bodyObj];
  const sections = body?.["hp:SECTION"] ?? body?.SECTION ?? [];
  return Array.isArray(sections) ? sections : [sections];
}
function getTag(obj, ...names) {
  for (const n of names) {
    const v = obj?.[n];
    if (v != null) return toArr(v);
  }
  return [];
}
function extractMeta(headObj) {
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const info = root?.["hh:DOCSUMMARY"]?.[0] ?? root?.DOCSUMMARY?.[0];
    if (!info) return {};
    const a = (k) => info?.[`hh:${k}`]?.[0]?._text ?? info?.[k]?.[0]?._text ?? "";
    return {
      title: a("TITLE") || void 0,
      author: a("AUTHOR") || void 0,
      subject: a("SUBJECT") || void 0
    };
  } catch {
    return {};
  }
}
function extractDims(headObj) {
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return null;
    const secPrList = refList?.["hh:SECPRLST"]?.[0]?.["hh:SECPR"] ?? refList?.SECPRLST?.[0]?.SECPR;
    const sec = Array.isArray(secPrList) ? secPrList[0] : secPrList;
    if (!sec) return null;
    const pa = sec?.["hh:PAGEPROPERTY"]?.[0]?._attr ?? sec?.PAGEPROPERTY?.[0]?._attr;
    if (!pa) return null;
    const ew = Number(pa.Width ?? 59528);
    const eh = Number(pa.Height ?? 84188);
    return {
      wPt: Metric.hwpToPt(ew),
      hPt: Metric.hwpToPt(eh),
      mt: Metric.hwpToPt(Number(pa.TopMargin ?? 5670)),
      mb: Metric.hwpToPt(Number(pa.BottomMargin ?? 4252)),
      ml: Metric.hwpToPt(Number(pa.LeftMargin ?? 8504)),
      mr: Metric.hwpToPt(Number(pa.RightMargin ?? 8504)),
      orient: ew > eh ? "landscape" : "portrait"
    };
  } catch {
    return null;
  }
}
function extractBorderFills(headObj) {
  const map = /* @__PURE__ */ new Map();
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;
    const bfList = refList?.["hh:borderFills"]?.[0] ?? refList?.["hh:BORDERFILLLIST"]?.[0] ?? refList?.BORDERFILLLIST?.[0];
    if (!bfList) return map;
    const bfs = getTag(bfList, "hh:borderFill", "hh:BORDERFILL");
    for (const bf of bfs) {
      const attr = bf?._attr ?? {};
      const id = Number(attr.id ?? 0);
      if (id === 0) continue;
      const info = {};
      const parseBorderEl = (el) => {
        if (!el) return void 0;
        const a = el?._attr ?? {};
        const mmVal = parseFloat(a.width) || void 0;
        const hwpVal = mmVal != null ? mmVal * 2.835 * 100 : void 0;
        return safeStrokeHwpx(a.type, hwpVal, a.color);
      };
      const topEl = bf?.["hh:topBorder"]?.[0] ?? bf?.["hh:top"]?.[0] ?? bf?.top?.[0];
      const rightEl = bf?.["hh:rightBorder"]?.[0] ?? bf?.["hh:right"]?.[0] ?? bf?.right?.[0];
      const bottomEl = bf?.["hh:bottomBorder"]?.[0] ?? bf?.["hh:bottom"]?.[0] ?? bf?.bottom?.[0];
      const leftEl = bf?.["hh:leftBorder"]?.[0] ?? bf?.["hh:left"]?.[0] ?? bf?.left?.[0];
      info.top = parseBorderEl(topEl);
      info.right = parseBorderEl(rightEl);
      info.bottom = parseBorderEl(bottomEl);
      info.left = parseBorderEl(leftEl);
      info.stroke = info.top ?? info.left ?? info.right ?? info.bottom;
      const fillBrush = bf?.["hc:fillBrush"]?.[0] ?? bf?.["hh:fillBrush"]?.[0] ?? bf?.["hh:fill"]?.[0] ?? bf?.fill?.[0] ?? bf?.fillBrush?.[0];
      if (fillBrush) {
        const winBrush = fillBrush?.["hc:winBrush"]?.[0]?._attr ?? fillBrush?.["hh:winBrush"]?.[0]?._attr ?? fillBrush?.winBrush?.[0]?._attr;
        if (winBrush?.faceColor && winBrush.faceColor !== "none") {
          info.bgColor = safeHex(winBrush.faceColor);
        }
      }
      map.set(id, info);
    }
  } catch {
  }
  return map;
}
function buildFontIdMap(headObj) {
  const fontMap = /* @__PURE__ */ new Map();
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return fontMap;
    const fontfaces = refList?.["hh:fontfaces"]?.[0] ?? refList?.["hh:FONTFACES"]?.[0];
    if (!fontfaces) return fontMap;
    const ffGroups = getTag(fontfaces, "hh:fontface", "hh:FONTFACE");
    for (const ff of ffGroups) {
      const fonts = getTag(ff, "hh:font", "hh:FONT");
      for (const font of fonts) {
        const fa = font?._attr ?? {};
        const fid = Number(fa.id ?? -1);
        const name = fa.face ?? fa.name ?? fa.Face ?? "";
        if (fid >= 0 && name && !fontMap.has(fid)) fontMap.set(fid, name);
      }
      if (fontMap.size > 0) break;
    }
  } catch {
  }
  return fontMap;
}
function extractCharPrs(headObj) {
  const map = /* @__PURE__ */ new Map();
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;
    const fontIdMap = buildFontIdMap(headObj);
    const cpList = refList?.["hh:charProperties"]?.[0] ?? refList?.["hh:CHARPROPERTIES"]?.[0];
    if (!cpList) return map;
    const cps = getTag(cpList, "hh:charPr", "hh:CHARPR");
    for (const cp of cps) {
      const attr = cp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;
      const info = {};
      if (attr.height) info.pt = Metric.hHeightToPt(Number(attr.height));
      if (attr.textColor) info.color = safeHex(attr.textColor);
      if (cp?.["hh:bold"]?.[0] != null) info.b = true;
      if (cp?.["hh:italic"]?.[0] != null) info.i = true;
      const ulAttr = cp?.["hh:underline"]?.[0]?._attr;
      if (ulAttr?.type && ulAttr.type !== "NONE") info.u = true;
      const stAttr = cp?.["hh:strikeout"]?.[0]?._attr;
      if (stAttr?.shape && stAttr.shape !== "NONE" && stAttr.shape !== "3D")
        info.s = true;
      const fontRefAttr = cp?.["hh:fontRef"]?.[0]?._attr ?? cp?.["hh:FONTREF"]?.[0]?._attr;
      if (fontRefAttr) {
        const fid = Number(
          fontRefAttr.hangul ?? fontRefAttr.latin ?? fontRefAttr.Hangul ?? 0
        );
        const name = fontIdMap.get(fid);
        if (name) info.font = safeFont(name);
      }
      map.set(id, info);
    }
  } catch {
  }
  return map;
}
function extractParaPrs(headObj) {
  const map = /* @__PURE__ */ new Map();
  try {
    const root = headObj?.["hh:head"]?.[0] ?? headObj?.["hh:HEAD"]?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;
    const ppList = refList?.["hh:paraProperties"]?.[0] ?? refList?.["hh:PARAPROPERTIES"]?.[0];
    if (!ppList) return map;
    const pps = getTag(ppList, "hh:paraPr", "hh:PARAPR");
    for (const pp of pps) {
      const attr = pp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;
      const alignNode = pp?.["hh:align"]?.[0]?._attr ?? pp?.["hh:ALIGN"]?.[0]?._attr;
      const align = alignNode?.horizontal ?? alignNode?.Horizontal;
      let marginEl = pp?.["hh:margin"]?.[0] ?? null;
      let lineSpEl = pp?.["hh:lineSpacing"]?.[0] ?? null;
      if (!marginEl) {
        const sw = pp?.["hp:switch"]?.[0];
        const container = sw?.["hp:default"]?.[0] ?? sw?.["hp:case"]?.[0];
        marginEl = container?.["hh:margin"]?.[0] ?? null;
        lineSpEl = lineSpEl ?? container?.["hh:lineSpacing"]?.[0] ?? null;
      }
      let indentPt;
      let indentRightPt;
      let firstLineIndentPt;
      let spaceBefore;
      let spaceAfter;
      let lineHeight;
      let lineHeightFixed;
      if (marginEl) {
        const leftEl = marginEl?.["hc:left"]?.[0];
        const rightEl = marginEl?.["hc:right"]?.[0];
        const indentEl = marginEl?.["hc:intent"]?.[0] ?? marginEl?.["hc:indent"]?.[0];
        const prevEl = marginEl?.["hc:prev"]?.[0];
        const nextEl = marginEl?.["hc:next"]?.[0];
        const leftVal = Number(leftEl?._attr?.value ?? 0);
        const rightVal = Number(rightEl?._attr?.value ?? 0);
        const indentVal = Number(indentEl?._attr?.value ?? 0);
        const prevVal = Number(prevEl?._attr?.value ?? 0);
        const nextVal = Number(nextEl?._attr?.value ?? 0);
        if (leftVal !== 0) indentPt = Metric.hwpToPt(leftVal);
        if (rightVal !== 0) indentRightPt = Metric.hwpToPt(rightVal);
        if (indentVal !== 0) firstLineIndentPt = Metric.hwpToPt(indentVal);
        if (prevVal > 0) spaceBefore = Metric.hwpToPt(prevVal);
        if (nextVal > 0) spaceAfter = Metric.hwpToPt(nextVal);
      }
      if (lineSpEl) {
        const lsAttr = lineSpEl._attr ?? {};
        const lsType = lsAttr.type ?? "PERCENT";
        const lsVal = Number(lsAttr.value ?? 160);
        if (lsType === "PERCENT" && lsVal > 0 && lsVal !== 160) {
          lineHeight = lsVal / 100;
        } else if (lsType === "FIXED" && lsVal > 0) {
          lineHeightFixed = Metric.hwpToPt(lsVal);
        }
      }
      map.set(id, {
        align,
        indentPt,
        indentRightPt,
        firstLineIndentPt,
        spaceBefore,
        spaceAfter,
        lineHeight,
        lineHeightFixed
      });
    }
  } catch {
  }
  return map;
}
function addParaItems(p, items) {
  const runs = getTag(p, "hp:run", "hp:RUN");
  let hasTable = false;
  for (const run of runs) {
    const tbls = getTag(run, "hp:tbl", "hp:TABLE");
    if (tbls.length > 0) {
      for (const tbl of tbls) {
        items.push({ type: "table", node: tbl });
      }
      hasTable = true;
    }
  }
  if (!hasTable) {
    items.push({ type: "para", node: p });
  }
}
function decodeSection(sec, dims, ctx) {
  const firstParas = getTag(sec, "hp:p", "hp:P");
  const pageDims = extractSecPrDims(firstParas[0]) ?? dims;
  const items = [];
  const paras = getTag(sec, "hp:p", "hp:P");
  const tbls = getTag(sec, "hp:tbl", "hp:TABLE");
  const childOrder = sec?.["_childOrder"];
  if (Array.isArray(childOrder)) {
    let pi = 0;
    let ti = 0;
    for (const tag of childOrder) {
      if ((tag === "hp:p" || tag === "hp:P") && pi < paras.length) {
        addParaItems(paras[pi++], items);
      } else if ((tag === "hp:tbl" || tag === "hp:TABLE") && ti < tbls.length) {
        items.push({ type: "table", node: tbls[ti++] });
      }
    }
    while (pi < paras.length) addParaItems(paras[pi++], items);
    while (ti < tbls.length) items.push({ type: "table", node: tbls[ti++] });
  } else {
    for (const p of paras) addParaItems(p, items);
    for (const t of tbls) items.push({ type: "table", node: t });
  }
  const kids = ctx.shield.guardAll(
    items,
    (item) => {
      if (item.type === "table") {
        try {
          const { value } = ctx.shield.guardGrid(
            item.node,
            (n) => decodeGrid(n, ctx),
            (n) => decodeGridSimple(n, ctx),
            (n) => decodeGridFlat(n),
            (n) => decodeGridText(n),
            "hwpx:table"
          );
          return value;
        } catch {
          return buildPara([buildSpan("[\uD45C \uD30C\uC2F1 \uC2E4\uD328]")]);
        }
      }
      return decodePara(item.node, ctx);
    },
    () => buildPara([buildSpan("[\uD30C\uC2F1 \uC2E4\uD328]")]),
    "hwpx:content"
  );
  const headerParas = decodeHeaderFooter(sec, "header", ctx);
  const footerParas = decodeHeaderFooter(sec, "footer", ctx);
  return buildSheet(kids.filter(Boolean), pageDims, {
    headers: { default: headerParas },
    footers: { default: footerParas }
  });
}
function parseSecPrDims(secPr) {
  const pagePr = secPr?.["hp:pagePr"]?.[0]?._attr ?? secPr?.["hp:PAGEPR"]?.[0]?._attr;
  if (!pagePr) return null;
  const margin = secPr?.["hp:pagePr"]?.[0]?.["hp:margin"]?.[0]?._attr ?? secPr?.["hp:PAGEPR"]?.[0]?.["hp:MARGIN"]?.[0]?._attr ?? {};
  const pw = Number(pagePr.width ?? 59528);
  const ph = Number(pagePr.height ?? 84188);
  return {
    wPt: Metric.hwpToPt(pw),
    hPt: Metric.hwpToPt(ph),
    mt: Metric.hwpToPt(Number(margin.top ?? 5670)),
    mb: Metric.hwpToPt(Number(margin.bottom ?? 4252)),
    ml: Metric.hwpToPt(Number(margin.left ?? 8504)),
    mr: Metric.hwpToPt(Number(margin.right ?? 8504)),
    orient: pw > ph ? "landscape" : "portrait"
  };
}
function extractSecPrDims(p) {
  if (!p) return null;
  try {
    const secPrDirect = p?.["hp:secPr"]?.[0] ?? p?.["hp:SECPR"]?.[0];
    if (secPrDirect) {
      const dims = parseSecPrDims(secPrDirect);
      if (dims) return dims;
    }
    const runs = getTag(p, "hp:run", "hp:RUN");
    for (const run of runs) {
      const secPr = run?.["hp:secPr"]?.[0] ?? run?.["hp:SECPR"]?.[0];
      if (!secPr) continue;
      const dims = parseSecPrDims(secPr);
      if (dims) return dims;
    }
  } catch {
  }
  return null;
}
function decodeHeaderFooter(sec, kind, ctx) {
  try {
    const hf = sec?.["hp:headerFooter"]?.[0] ?? sec?.["hp:HEADERFOOTER"]?.[0] ?? sec?.headerFooter?.[0] ?? sec?.HEADERFOOTER?.[0];
    if (!hf) return void 0;
    const part = hf?.["hp:" + kind]?.[0] ?? hf?.["hp:" + kind.toUpperCase()]?.[0] ?? hf?.[kind]?.[0] ?? hf?.[kind.toUpperCase()]?.[0];
    if (!part) return void 0;
    const paras = getTag(part, "hp:p", "hp:P");
    if (paras.length === 0) return void 0;
    return paras.map((p) => decodePara(p, ctx));
  } catch {
    return void 0;
  }
}
function decodePara(p, ctx) {
  const pAttr = p?._attr ?? {};
  const paraPrIdRef = Number(pAttr.paraPrIDRef ?? -1);
  let align;
  const paraPrDef = ctx.paraPrs.get(paraPrIdRef);
  if (paraPrDef?.align) align = paraPrDef.align;
  const inlineParaPr = p?.["hp:PARAPR"]?.[0] ?? p?.["hp:paraPr"]?.[0] ?? p?.PARAPR?.[0];
  if (inlineParaPr) {
    const alignNode = inlineParaPr?.["hp:ALIGN"]?.[0]?._attr ?? inlineParaPr?.["hp:align"]?.[0]?._attr ?? inlineParaPr?.ALIGN?.[0]?._attr;
    if (alignNode?.Type) align = alignNode.Type;
    if (alignNode?.horizontal) align = alignNode.horizontal;
  }
  const inlineAttr = inlineParaPr?._attr ?? {};
  const props = { align: safeAlign(align) };
  if (paraPrDef) {
    if (paraPrDef.indentPt !== void 0) props.indentPt = paraPrDef.indentPt;
    if (paraPrDef.indentRightPt !== void 0)
      props.indentRightPt = paraPrDef.indentRightPt;
    if (paraPrDef.firstLineIndentPt !== void 0)
      props.firstLineIndentPt = paraPrDef.firstLineIndentPt;
    if (paraPrDef.spaceBefore !== void 0)
      props.spaceBefore = paraPrDef.spaceBefore;
    if (paraPrDef.spaceAfter !== void 0)
      props.spaceAfter = paraPrDef.spaceAfter;
    if (paraPrDef.lineHeight !== void 0)
      props.lineHeight = paraPrDef.lineHeight;
    if (paraPrDef.lineHeightFixed !== void 0)
      props.lineHeightFixed = paraPrDef.lineHeightFixed;
  }
  if (inlineAttr.listType) {
    props.listOrd = inlineAttr.listType === "DIGIT" || inlineAttr.listType === "DECIMAL";
    props.listLv = Number(inlineAttr.listLevel ?? 0);
  }
  const runs = getTag(p, "hp:run", "hp:RUN");
  const kids = [];
  const collectPics = (container) => {
    const direct = getTag(container, "hp:pic", "hp:PIC");
    const ctrls = getTag(container, "hp:ctrl", "hp:CTRL");
    const nested = ctrls.flatMap((c) => getTag(c, "hp:pic", "hp:PIC"));
    return [...direct, ...nested];
  };
  for (const pic of collectPics(p)) {
    const img = decodePic(pic, ctx);
    if (img) kids.push(img);
  }
  for (const run of runs) {
    for (const pic of collectPics(run)) {
      const img = decodePic(pic, ctx);
      if (img) kids.push(img);
    }
    const pageNums = getTag(run, "hp:pageNum", "hp:PAGENUM");
    if (pageNums.length > 0) {
      const pn = pageNums[0]?._attr ?? {};
      const fmt = pn.formatType === "ROMAN_LOWER" ? "roman" : pn.formatType === "ROMAN_UPPER" ? "romanCaps" : "decimal";
      const pageNumNode = { tag: "pagenum", format: fmt };
      const spanProps = resolveCharPr(run, ctx);
      kids.push({ tag: "span", props: spanProps, kids: [pageNumNode] });
      continue;
    }
    const runPics = collectPics(run);
    const textNodes = getTag(run, "hp:t", "hp:T", "hp:CHAR");
    const content = textNodes.map((t) => {
      const val = typeof t === "string" ? t : t?._text ?? t?._ ?? t?.["#text"] ?? "";
      return val.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
    }).join("");
    if (content === "" && (run?.["hp:secPr"]?.[0] || run?.["hp:SECPR"]?.[0]) && runPics.length === 0 && pageNums.length === 0)
      continue;
    if (content !== "" || runPics.length === 0 && pageNums.length === 0) {
      const spanProps = resolveCharPr(run, ctx);
      kids.push(buildSpan(content, spanProps));
    }
  }
  if (pAttr.pageBreak === "1") {
    kids.unshift({ tag: "span", props: {}, kids: [buildPb()] });
  }
  return buildPara(kids.filter(Boolean), props);
}
function resolveCharPr(run, ctx) {
  const runAttr = run?._attr ?? {};
  const charPrIdRef = Number(runAttr.charPrIDRef ?? runAttr.CharPrIDRef ?? -1);
  const def = ctx.charPrs.get(charPrIdRef);
  if (def) {
    return {
      b: def.b,
      i: def.i,
      u: def.u,
      s: def.s,
      pt: def.pt,
      color: def.color,
      font: def.font,
      bg: def.bg
    };
  }
  const inlinePr = run?.["hp:CHARPR"]?.[0] ?? run?.["hp:charPr"]?.[0] ?? run?.CHARPR?.[0] ?? run?.charPr?.[0];
  const ca = inlinePr?._attr ?? {};
  const bVal = ca.Bold ?? ca.bold ?? ca.B ?? "";
  const iVal = ca.Italic ?? ca.italic ?? ca.I ?? "";
  const uVal = ca.Underline ?? ca.underline ?? "";
  const sVal = ca.Strikeout ?? ca.strikeout ?? "";
  const fontName = ca.FontName ?? ca.fontName ?? ca.FaceNameHangul ?? ca.faceNameHangul ?? "";
  const heightVal = ca.Height ?? ca.height ?? "";
  return {
    b: bVal === "1" || bVal === "true" || bVal === "True" || void 0,
    i: iVal === "1" || iVal === "true" || iVal === "True" || void 0,
    u: uVal && uVal !== "NONE" ? true : void 0,
    s: sVal && sVal !== "NONE" && sVal !== "3D" ? true : void 0,
    font: fontName ? safeFont(fontName) : void 0,
    pt: heightVal ? Metric.hHeightToPt(Number(heightVal)) : void 0,
    color: safeHex(ca.TextColor ?? ca.textColor),
    bg: safeHex(ca.BgColor ?? ca.bgColor)
  };
}
function decodePic(pic, ctx) {
  try {
    const szAttr = pic?.["hp:sz"]?.[0]?._attr ?? pic?.sz?.[0]?._attr ?? {};
    const w = Metric.hwpToPt(Number(szAttr.width ?? 0));
    const h = Metric.hwpToPt(Number(szAttr.height ?? 0));
    const imgNode = pic?.["hp:img"]?.[0]?._attr ?? pic?.["hc:img"]?.[0]?._attr ?? pic?.img?.[0]?._attr ?? {};
    const binRef = imgNode.binaryItemIDRef ?? imgNode.BinaryItemIDRef;
    if (!binRef) return null;
    let imgData;
    for (const [key, val] of ctx.files) {
      if (key.includes(binRef) || key.toLowerCase().includes(binRef.toLowerCase())) {
        imgData = val;
        break;
      }
    }
    if (!imgData) return null;
    const ext = binRef.split(".").pop()?.toLowerCase() ?? "png";
    const mimeMap = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp"
    };
    const posAttr = pic?.["hp:pos"]?.[0]?._attr ?? pic?.pos?.[0]?._attr ?? {};
    const layout = extractHwpxLayout(posAttr, pic);
    return buildImg(
      TextKit.base64Encode(imgData),
      mimeMap[ext] ?? "image/png",
      w,
      h,
      void 0,
      layout
    );
  } catch {
    return null;
  }
}
function extractHwpxLayout(posAttr, pic) {
  const treatAsChar = posAttr.treatAsChar === "1" || posAttr.treatAsChar === "true";
  if (treatAsChar) return { wrap: "inline" };
  const textWrap = pic?._attr?.textWrap ?? pic?.pic?.[0]?._attr?.textWrap ?? "TOP_AND_BOTTOM";
  const wrapMap = {
    TOP_AND_BOTTOM: "topAndBottom",
    // float, 위아래 텍스트 흐름
    SQUARE: "square",
    BOTH_SIDES: "tight",
    LEFT: "tight",
    RIGHT: "tight",
    LARGER_ONLY: "tight",
    SMALLER_ONLY: "tight",
    LARGEST_ONLY: "tight",
    BEHIND_TEXT: "behind",
    FRONT_TEXT: "front"
  };
  const wrap = wrapMap[textWrap] ?? "square";
  const horzRelToMap = {
    PARA: "para",
    MARGIN: "margin",
    PAGE: "page",
    COLUMN: "column"
  };
  const vertRelToMap = {
    PARA: "para",
    MARGIN: "margin",
    PAGE: "page",
    PAPER: "page",
    LINE: "line"
  };
  const horzRelTo = horzRelToMap[posAttr.horzRelTo ?? ""] ?? "para";
  const vertRelTo = vertRelToMap[posAttr.vertRelTo ?? ""] ?? "para";
  const horzAlignMap = {
    LEFT: "left",
    CENTER: "center",
    RIGHT: "right"
  };
  const vertAlignMap = {
    TOP: "top",
    CENTER: "center",
    BOTTOM: "bottom"
  };
  const horzAlign = horzAlignMap[posAttr.horzAlign ?? ""];
  const vertAlign = vertAlignMap[posAttr.vertAlign ?? ""];
  const horzOffset = Number(posAttr.horzOffset ?? 0);
  const vertOffset = Number(posAttr.vertOffset ?? 0);
  const xPt = horzOffset !== 0 ? Metric.hwpToPt(horzOffset) : void 0;
  const yPt = vertOffset !== 0 ? Metric.hwpToPt(vertOffset) : void 0;
  return { wrap, horzAlign, vertAlign, horzRelTo, vertRelTo, xPt, yPt };
}
function decodeGrid(tbl, ctx) {
  const tblAttr = tbl?._attr ?? {};
  const borderFillId = Number(tblAttr.borderFillIDRef ?? 0);
  const borderFill = ctx.borderFills.get(borderFillId);
  const headerRow = tblAttr.repeatHeader === "1";
  const gridProps = { headerRow: headerRow || void 0 };
  if (borderFill?.stroke) gridProps.defaultStroke = borderFill.stroke;
  const rowArr = getTag(tbl, "hp:tr", "hp:ROW");
  for (const row of rowArr) {
    const cells = getTag(row, "hp:tc", "hp:CELL");
    const rowWidths = [];
    let allSingle = true;
    for (const cell of cells) {
      const cellSpanAttr = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cs = Number(cellSpanAttr.colSpan ?? cell?._attr?.ColSpan ?? 1);
      if (cs > 1) {
        allSingle = false;
        break;
      }
      const szAttr = cell?.["hp:cellSz"]?.[0]?._attr ?? {};
      const w = Number(szAttr.width ?? 0);
      rowWidths.push(Metric.hwpToPt(w));
    }
    if (allSingle && rowWidths.length > 0 && rowWidths.some((w) => w > 0)) {
      gridProps.colWidths = rowWidths;
      break;
    }
  }
  if (!gridProps.colWidths) {
    let detectedCols = 0;
    for (const row of rowArr) {
      let ci = 0;
      for (const cell of getTag(row, "hp:tc", "hp:CELL")) {
        const csEl = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
        ci += Number(csEl.colSpan ?? cell?._attr?.ColSpan ?? 1);
      }
      if (ci > detectedCols) detectedCols = ci;
    }
    if (detectedCols > 0) {
      const sums = new Float64Array(detectedCols);
      const counts = new Int32Array(detectedCols);
      for (const row of rowArr) {
        let ci = 0;
        for (const cell of getTag(row, "hp:tc", "hp:CELL")) {
          const csEl = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
          const cs = Number(csEl.colSpan ?? cell?._attr?.ColSpan ?? 1);
          const szAttr = cell?.["hp:cellSz"]?.[0]?._attr ?? {};
          const w = Number(szAttr.width ?? 0);
          if (w > 0 && cs > 0) {
            const perCol = w / cs;
            for (let k = 0; k < cs && ci + k < detectedCols; k++) {
              sums[ci + k] += perCol;
              counts[ci + k]++;
            }
          }
          ci += cs;
        }
      }
      const estimated = Array.from(sums).map(
        (s, i) => counts[i] > 0 ? Metric.hwpToPt(s / counts[i]) : 0
      );
      if (estimated.some((w) => w > 0)) gridProps.colWidths = estimated;
    }
  }
  const rowNodes = rowArr.map((row) => {
    const cellArr = getTag(row, "hp:tc", "hp:CELL");
    const cellNodes = cellArr.map((cell) => {
      const ca = cell?._attr ?? {};
      const cellBfId = Number(ca.borderFillIDRef ?? 0);
      const cellBf = ctx.borderFills.get(cellBfId);
      const cellProps = {
        bg: cellBf?.bgColor ?? safeHex(ca.BgColor)
      };
      if (cellBf) {
        cellProps.top = cellBf.top ?? cellBf.stroke;
        cellProps.bot = cellBf.bottom ?? cellBf.stroke;
        cellProps.left = cellBf.left ?? cellBf.stroke;
        cellProps.right = cellBf.right ?? cellBf.stroke;
      }
      const subList = cell?.["hp:subList"]?.[0] ?? cell?.subList?.[0];
      const subAttr = subList?._attr ?? {};
      if (subAttr.vertAlign) {
        const vaMap = {
          TOP: "top",
          CENTER: "mid",
          BOTTOM: "bot"
        };
        cellProps.va = vaMap[subAttr.vertAlign];
      }
      const HWPX_DEFAULT_MARGIN_LR = 360;
      const HWPX_DEFAULT_MARGIN_TB = 141;
      const mL = Number(subAttr.marginLeft ?? HWPX_DEFAULT_MARGIN_LR);
      const mR = Number(subAttr.marginRight ?? HWPX_DEFAULT_MARGIN_LR);
      const mT = Number(subAttr.marginTop ?? HWPX_DEFAULT_MARGIN_TB);
      const mB = Number(subAttr.marginBottom ?? HWPX_DEFAULT_MARGIN_TB);
      if (mL !== HWPX_DEFAULT_MARGIN_LR) cellProps.padL = Metric.hwpToPt(mL);
      if (mR !== HWPX_DEFAULT_MARGIN_LR) cellProps.padR = Metric.hwpToPt(mR);
      if (mT !== HWPX_DEFAULT_MARGIN_TB) cellProps.padT = Metric.hwpToPt(mT);
      if (mB !== HWPX_DEFAULT_MARGIN_TB) cellProps.padB = Metric.hwpToPt(mB);
      const cellSpan = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cs = Number(cellSpan.colSpan ?? ca.ColSpan ?? 1);
      const rs = Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1);
      const cellKids = [];
      const source = subList ?? cell;
      const sourcePSource = getTag(source, "hp:p", "hp:P");
      for (const sp of sourcePSource) {
        try {
          const runs = getTag(sp, "hp:run", "hp:RUN");
          let hasNestedTable = false;
          for (const run of runs) {
            const nestedTbls = getTag(run, "hp:tbl", "hp:TABLE");
            for (const nestedTbl of nestedTbls) {
              try {
                cellKids.push(decodeGrid(nestedTbl, ctx));
              } catch {
              }
              hasNestedTable = true;
            }
          }
          if (!hasNestedTable) {
            cellKids.push(decodePara(sp, ctx));
          }
        } catch {
        }
      }
      return buildCell(
        cellKids.length > 0 ? cellKids : [buildPara([buildSpan("")])],
        { cs, rs, props: cellProps }
      );
    });
    let rowHeightPt;
    for (const cell of cellArr) {
      const ca = cell?._attr ?? {};
      const cellSpan = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cellRs = Math.max(1, Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1));
      const hSz = cell?.["hp:cellSz"]?.[0]?._attr ?? {};
      const hVal = Number(hSz.height ?? 0);
      if (hVal > 0) {
        rowHeightPt = Metric.hwpToPt(hVal) / cellRs;
        if (cellRs === 1) break;
      }
    }
    return buildRow(cellNodes, rowHeightPt);
  });
  return buildGrid(rowNodes, gridProps);
}
function decodeGridSimple(tbl, ctx) {
  const rowArr = getTag(tbl, "hp:tr", "hp:ROW");
  const rowNodes = rowArr.map((row) => {
    const cellArr = getTag(row, "hp:tc", "hp:CELL");
    return buildRow(
      cellArr.map(
        (cell) => buildCell([buildPara([buildSpan(cellText(cell))])])
      )
    );
  });
  return buildGrid(rowNodes);
}
function decodeGridFlat(tbl) {
  return buildGrid([
    buildRow([buildCell([buildPara([buildSpan(tableText(tbl))])])])
  ]);
}
function decodeGridText(tbl) {
  return buildPara([buildSpan(tableText(tbl))]);
}
function cellText(cell) {
  const subList = cell?.["hp:subList"]?.[0] ?? cell?.subList?.[0];
  const source = subList ?? cell;
  return getTag(source, "hp:p", "hp:P").map(
    (p) => getTag(p, "hp:run", "hp:RUN").map(
      (r) => getTag(r, "hp:t", "hp:T").map((t) => {
        const val = typeof t === "string" ? t : t?._text ?? t?._ ?? t?.["#text"] ?? "";
        return val.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      }).join("")
    ).join("")
  ).join(" ");
}
function tableText(tbl) {
  return getTag(tbl, "hp:tr", "hp:ROW").map(
    (row) => getTag(row, "hp:tc", "hp:CELL").map((c) => cellText(c)).join("	")
  ).join("\n");
}
function toArr(v) {
  return v == null ? [] : Array.isArray(v) ? v : [v];
}
registry.registerDecoder(new HwpxDecoder());

// src/toolkit/BinaryKit.ts
var BinaryKit = {
  readU16LE(buf, offset) {
    return buf[offset] | buf[offset + 1] << 8;
  },
  readU32LE(buf, offset) {
    return ((buf[offset] | buf[offset + 1] << 8 | buf[offset + 2] << 16) >>> 0) + buf[offset + 3] * 16777216;
  },
  isOle2(data) {
    return data.length >= 8 && data[0] === 208 && data[1] === 207 && data[2] === 17 && data[3] === 224 && data[4] === 161 && data[5] === 177 && data[6] === 26 && data[7] === 225;
  },
  parseCfb(data) {
    const streams = /* @__PURE__ */ new Map();
    if (!this.isOle2(data)) {
      throw new Error("Not a valid OLE2 file");
    }
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const sectorSize = 1 << view.getUint16(30, true);
    const miniSectorSz = 1 << view.getUint16(32, true);
    const dirFirstSec = view.getUint32(48, true);
    const miniStreamCutoff = view.getUint32(56, true);
    const miniFatFirst = view.getUint32(60, true);
    const miniFatCnt = view.getUint32(64, true);
    const difatFirst = view.getUint32(68, true);
    const ENDOFCHAIN = 4294967294;
    const FREESECT = 4294967295;
    const sectorAt = (sec) => data.subarray(512 + sec * sectorSize, 512 + (sec + 1) * sectorSize);
    const fatSecNums = [];
    for (let i = 0; i < 109; i++) {
      const s = view.getUint32(76 + i * 4, true);
      if (s === FREESECT || s === ENDOFCHAIN) break;
      fatSecNums.push(s);
    }
    if (difatFirst !== ENDOFCHAIN && difatFirst !== FREESECT) {
      let difSec = difatFirst;
      while (difSec !== ENDOFCHAIN && difSec !== FREESECT) {
        const sec = sectorAt(difSec);
        const sv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
        for (let i = 0; i < sectorSize / 4 - 1; i++) {
          const s = sv.getUint32(i * 4, true);
          if (s === FREESECT || s === ENDOFCHAIN) break;
          fatSecNums.push(s);
        }
        difSec = sv.getUint32(sectorSize - 4, true);
      }
    }
    const fat = [];
    for (const sec of fatSecNums) {
      const s = sectorAt(sec);
      const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
      for (let i = 0; i < sectorSize / 4; i++) {
        fat.push(sv.getUint32(i * 4, true));
      }
    }
    const readChain = (startSec) => {
      const chunks = [];
      let sec = startSec;
      while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < fat.length) {
        chunks.push(sectorAt(sec));
        sec = fat[sec];
      }
      return concatUint8(chunks);
    };
    const dirData = readChain(dirFirstSec);
    const dirView = new DataView(dirData.buffer, dirData.byteOffset, dirData.byteLength);
    const dirCount = dirData.length / 128;
    const dirEntries = [];
    for (let i = 0; i < dirCount; i++) {
      const base = i * 128;
      const nameLen = dirView.getUint16(base + 64, true);
      const nameBytes = dirData.subarray(base, base + Math.max(0, nameLen - 2));
      const name = new TextDecoder("utf-16le").decode(nameBytes);
      const type = dirData[base + 66];
      const childId = dirView.getInt32(base + 76, true);
      const sibLeft = dirView.getInt32(base + 68, true);
      const sibRight = dirView.getInt32(base + 72, true);
      const startSec = dirView.getUint32(base + 116, true);
      const size = dirView.getUint32(base + 120, true);
      dirEntries.push({ name, type, startSec, size, childId, siblingLeftId: sibLeft, siblingRightId: sibRight });
    }
    const rootEntry = dirEntries[0];
    let miniStreamData = null;
    let miniFat = [];
    if (rootEntry && rootEntry.startSec !== ENDOFCHAIN && rootEntry.startSec !== FREESECT) {
      miniStreamData = readChain(rootEntry.startSec);
    }
    if (miniFatCnt > 0 && miniFatFirst !== ENDOFCHAIN && miniFatFirst !== FREESECT) {
      const mfData = readChain(miniFatFirst);
      const mfv = new DataView(mfData.buffer, mfData.byteOffset, mfData.byteLength);
      for (let i = 0; i < mfData.length / 4; i++) {
        miniFat.push(mfv.getUint32(i * 4, true));
      }
    }
    const readMiniChain = (startSec, size) => {
      if (!miniStreamData) return new Uint8Array(0);
      const chunks = [];
      let sec = startSec;
      let remaining = size;
      while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < miniFat.length && remaining > 0) {
        const off = sec * miniSectorSz;
        const chunk = miniStreamData.subarray(off, off + Math.min(miniSectorSz, remaining));
        chunks.push(chunk);
        remaining -= chunk.length;
        sec = miniFat[sec];
      }
      return concatUint8(chunks).subarray(0, size);
    };
    const visit = (id, path) => {
      if (id < 0 || id >= dirEntries.length) return;
      const entry = dirEntries[id];
      const fullPath = path ? `${path}/${entry.name}` : entry.name;
      if (entry.type === 2) {
        let streamData;
        if (entry.size < miniStreamCutoff && miniStreamData) {
          streamData = readMiniChain(entry.startSec, entry.size);
        } else {
          streamData = readChain(entry.startSec).subarray(0, entry.size);
        }
        streams.set(fullPath, streamData);
        streams.set(entry.name, streamData);
      }
      if (entry.childId >= 0) visit(entry.childId, fullPath);
      if (entry.siblingLeftId >= 0) visit(entry.siblingLeftId, path);
      if (entry.siblingRightId >= 0) visit(entry.siblingRightId, path);
    };
    if (dirEntries.length > 0 && dirEntries[0].childId >= 0) {
      visit(dirEntries[0].childId, "");
    }
    return streams;
  }
};
function concatUint8(arrays) {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) {
    out.set(a, off);
    off += a.length;
  }
  return out;
}

// src/decoders/hwp/HwpScanner.ts
import pako2 from "pako";
var HWPTAG_BEGIN = 16;
var TAG_FACE_NAME = HWPTAG_BEGIN + 3;
var TAG_BORDER_FILL = HWPTAG_BEGIN + 4;
var TAG_CHAR_SHAPE = HWPTAG_BEGIN + 5;
var TAG_PARA_SHAPE = HWPTAG_BEGIN + 9;
var TAG_PARA_HEADER = HWPTAG_BEGIN + 50;
var TAG_PARA_TEXT = HWPTAG_BEGIN + 51;
var TAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52;
var TAG_CTRL_HEADER = HWPTAG_BEGIN + 55;
var TAG_PAGE_DEF = HWPTAG_BEGIN + 57;
var TAG_LIST_HEADER = HWPTAG_BEGIN + 56;
var TAG_TABLE_A = HWPTAG_BEGIN + 61;
var TAG_CELL_A = HWPTAG_BEGIN + 62;
var TAG_TABLE_B = HWPTAG_BEGIN + 64;
var TAG_CELL_B = HWPTAG_BEGIN + 65;
function isTableTag(t) {
  return t === TAG_TABLE_A || t === TAG_TABLE_B;
}
function isCellTag(t) {
  return t === TAG_CELL_A || t === TAG_CELL_B || t === TAG_LIST_HEADER;
}
var CTRL_TABLE = 1952607264;
var CTRL_IMAGE = 1768777504;
var CTRL_OBJ = 1868720672;
var CTRL_FIG = 1718183712;
var CTRL_GSO = 1735618336;
function parseRecords(data) {
  const out = [];
  let off = 0;
  while (off + 4 <= data.length) {
    const hdr = BinaryKit.readU32LE(data, off);
    const tag = hdr & 1023;
    const level = hdr >> 10 & 1023;
    let size = hdr >> 20 & 4095;
    off += 4;
    if (size === 4095) {
      if (off + 4 > data.length) break;
      size = BinaryKit.readU32LE(data, off);
      off += 4;
    }
    if (off + size > data.length) break;
    out.push({ tag, level, data: data.subarray(off, off + size) });
    off += size;
  }
  return out;
}
function tryInflate(data) {
  try {
    return pako2.inflate(data);
  } catch {
    try {
      return pako2.inflateRaw(data);
    } catch {
      return data;
    }
  }
}
function parseFileHeader(buf) {
  if (buf.length < 40) return { compressed: true, encrypted: false };
  const props = BinaryKit.readU32LE(buf, 36);
  return { compressed: (props & 1) !== 0, encrypted: (props & 2) !== 0 };
}
function parseDocInfo(data, compressed) {
  const raw = compressed ? tryInflate(data) : data;
  const recs = parseRecords(raw);
  const info = { faceNames: [], charShapes: [], paraShapes: [], borderFills: [] };
  for (const r of recs) {
    try {
      if (r.tag === TAG_FACE_NAME) info.faceNames.push(parseFaceName(r.data));
      if (r.tag === TAG_CHAR_SHAPE) info.charShapes.push(parseCharShape(r.data));
      if (r.tag === TAG_PARA_SHAPE) info.paraShapes.push(parseParaShape(r.data));
      if (r.tag === TAG_BORDER_FILL) info.borderFills.push(parseBorderFill(r.data));
    } catch {
    }
  }
  return info;
}
function parseFaceName(d) {
  if (d.length < 3) return "";
  const len = BinaryKit.readU16LE(d, 1);
  if (d.length < 3 + len * 2) return "";
  return new TextDecoder("utf-16le").decode(d.subarray(3, 3 + len * 2));
}
function parseCharShape(d) {
  const faceIds = [];
  for (let i = 0; i < 7; i++) faceIds.push(d.length >= (i + 1) * 2 ? BinaryKit.readU16LE(d, i * 2) : 0);
  const height = d.length >= 46 ? BinaryKit.readU32LE(d, 42) : 1e3;
  const attr = d.length >= 50 ? BinaryKit.readU32LE(d, 46) : 0;
  const ulType = attr >> 2 & 7;
  const skType = attr >> 18 & 7;
  const suType = attr >> 16 & 3;
  return {
    faceIds,
    height: height > 0 && height < 1e5 ? height : 1e3,
    italic: (attr & 1) !== 0,
    bold: (attr >> 1 & 1) !== 0,
    underline: ulType !== 0,
    strikeout: skType !== 0,
    superscript: suType === 1,
    subscript: suType === 2,
    textColor: d.length >= 56 ? colorRef(d, 52) : "000000"
  };
}
var ALIGN_TBL = { 0: "justify", 1: "left", 2: "right", 3: "center", 4: "justify" };
function parseParaShape(d) {
  if (d.length < 4) return { align: "left", spaceBefore: 0, spaceAfter: 0, lineSpacing: 160, leftMargin: 0, indent: 0 };
  const attr = BinaryKit.readU32LE(d, 0);
  return {
    align: ALIGN_TBL[attr >> 2 & 7] ?? "left",
    leftMargin: d.length >= 8 ? i32(d, 4) : 0,
    // offset 4: leftMargin (들여쓰기)
    indent: d.length >= 16 ? i32(d, 12) : 0,
    // offset 12: first-line indent
    spaceBefore: d.length >= 20 ? i32(d, 16) : 0,
    spaceAfter: d.length >= 24 ? i32(d, 20) : 0,
    lineSpacing: d.length >= 28 ? i32(d, 24) : 160
  };
}
var BORDER_W_PT = [0.28, 0.34, 0.43, 0.57, 0.71, 0.85, 1.13, 1.42, 1.7, 1.98, 2.84, 4.25, 5.67, 8.5, 11.34, 14.17];
var BORDER_KIND = { 0: "solid", 1: "dash", 2: "dash", 3: "dot", 4: "dash", 5: "dash", 6: "dash", 7: "double", 8: "double", 9: "double", 10: "none" };
function parseBorderFill(d) {
  const borders = [];
  const BASE_TYPE = 2;
  const BASE_WIDTH = 6;
  const BASE_COLOR = 10;
  for (let i = 0; i < 4; i++) {
    const type = BASE_TYPE + i < d.length ? d[BASE_TYPE + i] : 0;
    const widthPt = BASE_WIDTH + i < d.length ? BORDER_W_PT[d[BASE_WIDTH + i]] ?? 0.5 : 0.5;
    const color = BASE_COLOR + i * 4 + 4 <= d.length ? colorRef(d, BASE_COLOR + i * 4) : "000000";
    borders.push({ type, widthPt, color });
  }
  let bgColor;
  const fOff = 32;
  if (d.length >= fOff + 8) {
    const ft = BinaryKit.readU32LE(d, fOff);
    if (ft & 1) bgColor = colorRef(d, fOff + 4);
  }
  return { borders, bgColor };
}
function parseBody(raw, compressed, di, shield, gsoCtx) {
  const recs = parseRecords(compressed ? tryInflate(raw) : raw);
  const content = [];
  let pageDims;
  for (const r of recs) {
    if (r.tag === TAG_PAGE_DEF) {
      pageDims = shield.guard(() => parsePageDef(r.data), A4, "hwp:pageDef");
      break;
    }
  }
  let i = 0;
  while (i < recs.length) {
    if (recs[i].tag === TAG_PAGE_DEF) {
      i++;
    } else if (recs[i].tag === TAG_PARA_HEADER) {
      const r = shield.guard(
        () => parseParagraphGroup(recs, i, di, shield, gsoCtx),
        { nodes: [], next: i + 1 },
        `hwp:para@${i}`
      );
      content.push(...r.nodes);
      i = r.next;
    } else {
      i++;
    }
  }
  return { content, pageDims };
}
function parseParagraphGroup(recs, start, di, shield, gsoCtx) {
  const hdr = recs[start];
  const lv = hdr.level;
  const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
  const ps = di.paraShapes[psId];
  let text = null;
  let csPairs = [];
  const grids = [];
  const ctrlHeaders = [];
  let i = start + 1;
  while (i < recs.length && recs[i].level > lv) {
    const r = recs[i];
    if (r.tag === TAG_PARA_TEXT && r.level === lv + 1) {
      text = decodeParaText(r.data);
      i++;
    } else if (r.tag === TAG_PARA_CHAR_SHAPE && r.level === lv + 1) {
      csPairs = parseCharShapePairs(r.data);
      i++;
    } else if (r.tag === TAG_CTRL_HEADER && r.level === lv + 1) {
      if (r.data.length >= 4) {
        const ctrlId = BinaryKit.readU32LE(r.data, 0);
        const MAX_HWP = 1e6;
        const rawW = r.data.length >= 24 ? BinaryKit.readU32LE(r.data, 16) : 0;
        const rawH = r.data.length >= 28 ? BinaryKit.readU32LE(r.data, 20) : 0;
        const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
        const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
        const imgId = ctrlId === CTRL_GSO ? gsoCtx.count++ : r.data.length >= 6 ? BinaryKit.readU16LE(r.data, 4) : 0;
        ctrlHeaders.push({ ctrlId, imgId, wPt, hPt });
        if (ctrlId === CTRL_TABLE) {
          const tr = shield.guard(
            () => parseTableCtrl(recs, i, di, shield, gsoCtx),
            { grid: null, next: skipKids(recs, i) },
            `hwp:tbl@${i}`
          );
          if (tr.grid) grids.push(tr.grid);
          i = tr.next;
        } else {
          i = skipKids(recs, i);
        }
      } else {
        i = skipKids(recs, i);
      }
    } else {
      i++;
    }
  }
  const nodes = [];
  if (text && (text.chars.length > 0 || text.controls.length > 0)) {
    const paraContent = [];
    if (text.chars.length > 0) {
      const spans = resolveCharShapes(text.chars, csPairs, di);
      paraContent.push(...spans);
    }
    if (text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        const ch = ctrlHeaders[ci];
        if (!ch) continue;
        const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
        if (!isImg) continue;
        const dimStr = ch.wPt > 0 && ch.hPt > 0 ? `_W${Math.round(ch.wPt)}_H${Math.round(ch.hPt)}` : "";
        paraContent.push(buildSpan(`__EXT_${ch.imgId}${dimStr}__`));
      }
    }
    if (paraContent.length > 0) {
      nodes.push(buildPara(paraContent, buildParaProps(ps)));
    }
  }
  nodes.push(...grids);
  return { nodes, next: i };
}
function skipKids(recs, idx) {
  const lv = recs[idx].level;
  let i = idx + 1;
  while (i < recs.length && recs[i].level > lv) i++;
  return i;
}
var EXT_CTRL = /* @__PURE__ */ new Set([2, 3, 11, 12, 14, 15]);
var INL_CTRL = /* @__PURE__ */ new Set([4, 5, 6, 7, 8]);
function decodeParaText(d) {
  const chars = [];
  const controls = [];
  let i = 0, pos = 0;
  while (i + 1 < d.length) {
    const c = d[i] | d[i + 1] << 8;
    if (c === 0) {
      i += 2;
      pos++;
      continue;
    }
    if (c === 13) {
      break;
    }
    if (c === 10) {
      chars.push({ pos, ch: "\n" });
      i += 2;
      pos++;
      continue;
    }
    if (EXT_CTRL.has(c)) {
      let objId = 0;
      if (i + 16 <= d.length) {
        objId = BinaryKit.readU16LE(d, i + 8);
      }
      controls.push({ pos, ctrlId: 0, objId, matched: false });
      i += 16;
      pos += 8;
      continue;
    }
    if (INL_CTRL.has(c)) {
      i += 16;
      pos += 8;
      continue;
    }
    if (c === 9) {
      chars.push({ pos, ch: "	" });
      i += 16;
      pos += 8;
      continue;
    }
    if (c >= 1 && c <= 31) {
      i += 2;
      pos++;
      continue;
    }
    chars.push({ pos, ch: String.fromCharCode(c) });
    i += 2;
    pos++;
  }
  return { chars, controls };
}
function parseCharShapePairs(d) {
  const out = [];
  for (let i = 0; i + 7 < d.length; i += 8)
    out.push([BinaryKit.readU32LE(d, i), BinaryKit.readU32LE(d, i + 4)]);
  return out;
}
function resolveCharShapes(chars, pairs, di) {
  if (chars.length === 0) return [buildSpan("")];
  const defaultId = pairs.length > 0 ? pairs[0][1] : 0;
  function idFor(pos) {
    let id = defaultId;
    for (const [p, sid] of pairs) {
      if (p <= pos) id = sid;
      else break;
    }
    return id;
  }
  const spans = [];
  let curId = idFor(chars[0].pos);
  let buf = chars[0].ch;
  for (let k = 1; k < chars.length; k++) {
    const sid = idFor(chars[k].pos);
    if (sid !== curId) {
      spans.push(styledSpan(buf, curId, di));
      buf = "";
      curId = sid;
    }
    buf += chars[k].ch;
  }
  if (buf) spans.push(styledSpan(buf, curId, di));
  return spans;
}
function styledSpan(text, shapeId, di) {
  const cs = di.charShapes[shapeId];
  if (!cs) return buildSpan(text);
  const props = {};
  const fid = cs.faceIds[0] ?? 0;
  if (fid < di.faceNames.length && di.faceNames[fid]) props.font = safeFont(di.faceNames[fid]);
  if (cs.height > 0) props.pt = Metric.hwpToPt(cs.height);
  if (cs.bold) props.b = true;
  if (cs.italic) props.i = true;
  if (cs.underline) props.u = true;
  if (cs.strikeout) props.s = true;
  if (cs.superscript) props.sup = true;
  if (cs.subscript) props.sub = true;
  const hex = safeHex(cs.textColor);
  if (hex && hex !== "000000") props.color = hex;
  return buildSpan(text, props);
}
function parseTableCtrl(recs, ctrlIdx, di, shield, gsoCtx) {
  const ctrlLv = recs[ctrlIdx].level;
  let i = ctrlIdx + 1;
  let tblData = null;
  const cells = [];
  const tblLevel = ctrlLv + 1;
  while (i < recs.length && recs[i].level > ctrlLv) {
    const r = recs[i];
    if (isTableTag(r.tag) && r.level === tblLevel) {
      tblData = r.data;
      i++;
    } else if (r.tag === TAG_LIST_HEADER && r.level === tblLevel) {
      const cellData = r.data;
      const paraCount = cellData.length >= 2 ? BinaryKit.readU16LE(cellData, 0) : 0;
      i++;
      const cStart = i;
      let consumed = 0;
      while (i < recs.length && consumed < paraCount) {
        if (recs[i].tag === TAG_PARA_HEADER && recs[i].level === tblLevel) {
          consumed++;
          i++;
          while (i < recs.length && recs[i].level > tblLevel) i++;
        } else if (recs[i].level > tblLevel) {
          i++;
        } else {
          break;
        }
      }
      cells.push({ data: cellData, tag: TAG_LIST_HEADER, cStart, cEnd: i });
    } else if (isCellTag(r.tag) && r.level === tblLevel) {
      const cellData = r.data;
      const cellTag = r.tag;
      i++;
      const cStart = i;
      while (i < recs.length && recs[i].level > tblLevel) i++;
      cells.push({ data: cellData, tag: cellTag, cStart, cEnd: i });
    } else {
      i++;
    }
  }
  if (!tblData || cells.length === 0) return { grid: null, next: i };
  const rowCnt = tblData.length >= 6 ? BinaryKit.readU16LE(tblData, 4) : 1;
  const colCnt = tblData.length >= 8 ? BinaryKit.readU16LE(tblData, 6) : 1;
  const parsed = [];
  for (let ci = 0; ci < cells.length; ci++) {
    const c = cells[ci];
    const seqIdx = ci;
    const pc = shield.guard(
      () => parseCellRec(c.data, c.tag, recs, c.cStart, c.cEnd, di, shield, seqIdx, colCnt, gsoCtx),
      { row: Math.floor(ci / (colCnt || 1)), col: ci % (colCnt || 1), cs: 1, rs: 1, widthHwp: 0, heightHwp: void 0, props: {}, cellChildren: [buildPara([buildSpan("")])] },
      `hwp:cell@${c.cStart}`
    );
    parsed.push(pc);
  }
  const maxRow = parsed.reduce((m, c) => Math.max(m, c.row + c.rs), 0);
  const actualRowCnt = Math.max(rowCnt, maxRow);
  const posValid = parsed.every((c) => c.row >= 0 && c.col >= 0 && c.col < colCnt);
  if (!posValid) {
    let idx = 0;
    for (const c of parsed) {
      c.row = Math.floor(idx / colCnt);
      c.col = idx % colCnt;
      idx++;
    }
  }
  const colWidthsPt = new Array(colCnt).fill(0);
  for (const c of parsed) {
    if (c.cs === 1 && c.widthHwp > 0) {
      const wPt = Metric.hwpToPt(c.widthHwp);
      if (wPt > colWidthsPt[c.col]) colWidthsPt[c.col] = wPt;
    }
  }
  const zeroColumns = colWidthsPt.filter((w) => w === 0).length;
  if (zeroColumns > 0) {
    const spanCells = parsed.filter((c) => c.cs > 1 && c.widthHwp > 0).sort((a, b) => a.cs - b.cs);
    for (const c of spanCells) {
      if (c.cs > 1 && c.widthHwp > 0) {
        let known = 0;
        let unknownCols = 0;
        for (let ci = c.col; ci < c.col + c.cs && ci < colCnt; ci++) {
          if (colWidthsPt[ci] > 0) known += colWidthsPt[ci];
          else unknownCols++;
        }
        if (unknownCols > 0) {
          const remaining = Metric.hwpToPt(c.widthHwp) - known;
          const each = remaining > 0 ? remaining / unknownCols : 0;
          for (let ci = c.col; ci < c.col + c.cs && ci < colCnt; ci++) {
            if (colWidthsPt[ci] === 0 && each > 0) colWidthsPt[ci] = each;
          }
        }
      }
    }
  }
  for (let i2 = 0; i2 < colWidthsPt.length; i2++) {
    if (colWidthsPt[i2] > 0 && colWidthsPt[i2] < 1) colWidthsPt[i2] = 1;
  }
  const rows = [];
  for (let r = 0; r < actualRowCnt; r++) {
    const rc = parsed.filter((c) => c.row === r).sort((a, b) => a.col - b.col);
    if (rc.length === 0) continue;
    let rowHeightPt = void 0;
    for (const c of rc) {
      if (c.heightHwp && c.heightHwp > 0 && c.rs === 1) {
        const hPt = Metric.hwpToPt(c.heightHwp);
        if (rowHeightPt == null || hPt > rowHeightPt) rowHeightPt = hPt;
      }
    }
    if (rowHeightPt == null) {
      for (const c of rc) {
        if (c.heightHwp && c.heightHwp > 0) {
          const hPt = Metric.hwpToPt(c.heightHwp) / c.rs;
          if (rowHeightPt == null || hPt > rowHeightPt) rowHeightPt = hPt;
        }
      }
    }
    rows.push(buildRow(rc.map((c) => {
      return buildCell(c.cellChildren, { cs: c.cs, rs: c.rs, props: c.props });
    }), rowHeightPt));
  }
  if (rows.length === 0) return { grid: null, next: i };
  let defStroke;
  const bfOff = 18 + rowCnt * 2;
  if (tblData.length >= bfOff + 2) {
    const bfId = BinaryKit.readU16LE(tblData, bfOff);
    defStroke = strokeFromBF(bfId, di);
  }
  const gp = {};
  if (defStroke) gp.defaultStroke = defStroke;
  const hasWidths = colWidthsPt.some((w) => w > 0);
  if (hasWidths) gp.colWidths = colWidthsPt;
  return { grid: buildGrid(rows, gp), next: i };
}
function parseCellRec(d, tag, recs, cStart, cEnd, di, shield, seqIdx, colCnt, gsoCtx) {
  let col, row, cs = 1, rs = 1;
  let widthHwp = 0;
  let heightHwp = 0;
  const props = {};
  const attr = d.length >= 6 ? BinaryKit.readU32LE(d, 2) : 0;
  const va = attr >> 6 & 3;
  if (va === 1) props.va = "mid";
  else if (va === 2) props.va = "bot";
  const HWP_PAD_LR_DEFAULT = 360;
  const HWP_PAD_TB_DEFAULT = 141;
  if (tag === TAG_LIST_HEADER && d.length >= 22) {
    col = BinaryKit.readU16LE(d, 8);
    row = BinaryKit.readU16LE(d, 10);
    cs = Math.max(1, BinaryKit.readU16LE(d, 12));
    rs = Math.max(1, BinaryKit.readU16LE(d, 14));
    widthHwp = BinaryKit.readU32LE(d, 16);
    heightHwp = d.length >= 24 ? BinaryKit.readU32LE(d, 20) : 0;
    if (d.length >= 32) {
      const pL = BinaryKit.readU16LE(d, 24);
      const pR = BinaryKit.readU16LE(d, 26);
      const pT = BinaryKit.readU16LE(d, 28);
      const pB = BinaryKit.readU16LE(d, 30);
      if (pL !== HWP_PAD_LR_DEFAULT) props.padL = Metric.hwpToPt(pL);
      if (pR !== HWP_PAD_LR_DEFAULT) props.padR = Metric.hwpToPt(pR);
      if (pT !== HWP_PAD_TB_DEFAULT) props.padT = Metric.hwpToPt(pT);
      if (pB !== HWP_PAD_TB_DEFAULT) props.padB = Metric.hwpToPt(pB);
    }
    const bfId = d.length >= 34 ? BinaryKit.readU16LE(d, 32) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) applyCellBorderFill(di.borderFills[bfId - 1], props);
  } else if (tag !== TAG_LIST_HEADER) {
    col = d.length >= 8 ? BinaryKit.readU16LE(d, 6) : seqIdx % (colCnt || 1);
    row = d.length >= 10 ? BinaryKit.readU16LE(d, 8) : Math.floor(seqIdx / (colCnt || 1));
    cs = d.length >= 12 ? Math.max(1, BinaryKit.readU16LE(d, 10)) : 1;
    rs = d.length >= 14 ? Math.max(1, BinaryKit.readU16LE(d, 12)) : 1;
    widthHwp = d.length >= 18 ? BinaryKit.readU32LE(d, 14) : 0;
    heightHwp = d.length >= 22 ? BinaryKit.readU32LE(d, 18) : 0;
    if (d.length >= 30) {
      const pL = BinaryKit.readU16LE(d, 22);
      const pR = BinaryKit.readU16LE(d, 24);
      const pT = BinaryKit.readU16LE(d, 26);
      const pB = BinaryKit.readU16LE(d, 28);
      if (pL !== HWP_PAD_LR_DEFAULT) props.padL = Metric.hwpToPt(pL);
      if (pR !== HWP_PAD_LR_DEFAULT) props.padR = Metric.hwpToPt(pR);
      if (pT !== HWP_PAD_TB_DEFAULT) props.padT = Metric.hwpToPt(pT);
      if (pB !== HWP_PAD_TB_DEFAULT) props.padB = Metric.hwpToPt(pB);
    }
    const bfId = d.length >= 32 ? BinaryKit.readU16LE(d, 30) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) applyCellBorderFill(di.borderFills[bfId - 1], props);
  } else {
    row = Math.floor(seqIdx / (colCnt || 1));
    col = seqIdx % (colCnt || 1);
  }
  const cellChildren = [];
  const MAX_HWP = 1e6;
  let k = cStart;
  while (k < cEnd) {
    if (recs[k].tag === TAG_PARA_HEADER) {
      const r = shield.guard(
        () => {
          const hdr = recs[k];
          const lv = hdr.level;
          const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
          const ps = di.paraShapes[psId];
          let txt = null;
          let csp = [];
          const ctrlHdrs = [];
          const innerGrids = [];
          let j = k + 1;
          while (j < cEnd && recs[j].level > lv) {
            if (recs[j].tag === TAG_PARA_TEXT) {
              txt = decodeParaText(recs[j].data);
              j++;
            } else if (recs[j].tag === TAG_PARA_CHAR_SHAPE) {
              csp = parseCharShapePairs(recs[j].data);
              j++;
            } else if (recs[j].tag === TAG_CTRL_HEADER && recs[j].level === lv + 1) {
              if (recs[j].data.length >= 4) {
                const ctrlId = BinaryKit.readU32LE(recs[j].data, 0);
                if (ctrlId === CTRL_TABLE) {
                  const nestedTr = shield.guard(
                    () => parseTableCtrl(recs, j, di, shield, gsoCtx),
                    { grid: null, next: skipKids(recs, j) },
                    `hwp:innerNestedTbl@${j}`
                  );
                  if (nestedTr.grid) innerGrids.push(nestedTr.grid);
                  j = nestedTr.next;
                } else {
                  const rawW = recs[j].data.length >= 24 ? BinaryKit.readU32LE(recs[j].data, 16) : 0;
                  const rawH = recs[j].data.length >= 28 ? BinaryKit.readU32LE(recs[j].data, 20) : 0;
                  const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
                  const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
                  const imgId = ctrlId === CTRL_GSO ? gsoCtx.count++ : recs[j].data.length >= 6 ? BinaryKit.readU16LE(recs[j].data, 4) : 0;
                  ctrlHdrs.push({ ctrlId, imgId, wPt, hPt });
                  j = skipKids(recs, j);
                }
              } else {
                j = skipKids(recs, j);
              }
            } else j++;
          }
          const paraContent = [];
          if (txt && txt.chars.length > 0) paraContent.push(...resolveCharShapes(txt.chars, csp, di));
          if (txt && txt.controls.length > 0) {
            for (let ci = 0; ci < txt.controls.length; ci++) {
              const ch = ctrlHdrs[ci];
              if (!ch) continue;
              const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
              if (!isImg) continue;
              const dimStr = ch.wPt > 0 && ch.hPt > 0 ? `_W${Math.round(ch.wPt)}_H${Math.round(ch.hPt)}` : "";
              paraContent.push(buildSpan(`__EXT_${ch.imgId}${dimStr}__`));
            }
          }
          const kids = paraContent.length > 0 ? paraContent : [buildSpan("")];
          const items = [buildPara(kids, buildParaProps(ps)), ...innerGrids];
          return { items, next: j };
        },
        { items: [buildPara([buildSpan("")])], next: k + 1 },
        `hwp:cellP@${k}`
      );
      cellChildren.push(...r.items);
      k = r.next;
    } else if (recs[k].tag === TAG_CTRL_HEADER && recs[k].data.length >= 4) {
      const cellCtrlId = BinaryKit.readU32LE(recs[k].data, 0);
      if (cellCtrlId === CTRL_GSO) {
        const gsoId = gsoCtx.count++;
        const rawW = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 16) : 0;
        const rawH = recs[k].data.length >= 28 ? BinaryKit.readU32LE(recs[k].data, 20) : 0;
        const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
        const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
        const dimStr = wPt > 0 && hPt > 0 ? `_W${Math.round(wPt)}_H${Math.round(hPt)}` : "";
        cellChildren.push(buildPara([buildSpan(`__EXT_${gsoId}${dimStr}__`)]));
        k = skipKids(recs, k);
      } else if (cellCtrlId === CTRL_TABLE) {
        const tr = shield.guard(
          () => parseTableCtrl(recs, k, di, shield, gsoCtx),
          { grid: null, next: skipKids(recs, k) },
          `hwp:nestedTbl@${k}`
        );
        if (tr.grid) cellChildren.push(tr.grid);
        k = tr.next;
      } else {
        k = skipKids(recs, k);
      }
    } else {
      k++;
    }
  }
  return {
    row,
    col,
    cs,
    rs,
    props,
    widthHwp,
    heightHwp: heightHwp || void 0,
    cellChildren: cellChildren.length ? cellChildren : [buildPara([buildSpan("")])]
  };
}
function parsePageDef(d) {
  if (d.length < 24) return A4;
  const w = BinaryKit.readU32LE(d, 0);
  const h = BinaryKit.readU32LE(d, 4);
  const ml = BinaryKit.readU32LE(d, 8);
  const mr = BinaryKit.readU32LE(d, 12);
  const mt = BinaryKit.readU32LE(d, 16);
  const mb = BinaryKit.readU32LE(d, 20);
  const at = d.length >= 40 ? BinaryKit.readU32LE(d, 36) : 0;
  return {
    wPt: Metric.hwpToPt(w),
    hPt: Metric.hwpToPt(h),
    ml: Metric.hwpToPt(ml),
    mr: Metric.hwpToPt(mr),
    mt: Metric.hwpToPt(mt),
    mb: Metric.hwpToPt(mb),
    orient: at & 1 ? "landscape" : "portrait"
  };
}
function i32(d, o) {
  const u = BinaryKit.readU32LE(d, o);
  return u > 2147483647 ? u - 4294967296 : u;
}
function colorRef(d, o) {
  if (o + 3 > d.length) return "000000";
  return (d[o] << 16 | d[o + 1] << 8 | d[o + 2]).toString(16).padStart(6, "0").toUpperCase();
}
function toStroke(b) {
  return { kind: BORDER_KIND[b.type] ?? "solid", pt: b.widthPt, color: b.color };
}
function applyCellBorderFill(bf, props) {
  if (bf.borders.length >= 4) {
    props.left = toStroke(bf.borders[0]);
    props.right = toStroke(bf.borders[1]);
    props.top = toStroke(bf.borders[2]);
    props.bot = toStroke(bf.borders[3]);
  }
  if (bf.bgColor && bf.bgColor !== "FFFFFF") props.bg = bf.bgColor;
}
function strokeFromBF(bfId, di) {
  if (bfId <= 0 || bfId > di.borderFills.length) return void 0;
  const bf = di.borderFills[bfId - 1];
  if (!bf.borders.length) return void 0;
  const b = bf.borders[0];
  return { kind: BORDER_KIND[b.type] ?? "solid", pt: b.widthPt, color: b.color };
}
function buildParaProps(ps) {
  if (!ps) return {};
  const p = {};
  if (ps.align && ps.align !== "left") p.align = ps.align;
  if (ps.spaceBefore > 0) p.spaceBefore = Metric.hwpToPt(ps.spaceBefore);
  if (ps.spaceAfter > 0) p.spaceAfter = Metric.hwpToPt(ps.spaceAfter);
  if (ps.lineSpacing > 0 && ps.lineSpacing !== 160) p.lineHeight = ps.lineSpacing / 100;
  const leftMarginPt = Math.max(0, Metric.hwpToPt(ps.leftMargin));
  if (leftMarginPt > 0) p.leftMargin = leftMarginPt;
  if (ps.indent !== 0) p.firstLineIndentPt = Metric.hwpToPt(ps.indent);
  return p;
}
var HwpScanner = class {
  constructor() {
    this.format = "hwp";
    this.aliases = ["application/vnd.hancom.hwp"];
  }
  async decode(data) {
    const shield = new ShieldedParser();
    const warns = [];
    try {
      if (!BinaryKit.isOle2(data)) return fail("HWP: Invalid OLE2 signature");
      const streams = BinaryKit.parseCfb(data);
      const fh = streams.get("FileHeader");
      const { compressed, encrypted } = fh ? parseFileHeader(fh) : { compressed: true, encrypted: false };
      if (encrypted) return fail("HWP: \uC554\uD638\uD654\uB41C \uD30C\uC77C\uC740 \uC9C0\uC6D0\uD558\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4");
      const diRaw = streams.get("DocInfo");
      let di = { faceNames: [], charShapes: [], paraShapes: [], borderFills: [] };
      if (diRaw) {
        di = shield.guard(() => parseDocInfo(diRaw, compressed), di, "hwp:docInfo");
      }
      const binEntries = [];
      for (const [path, streamData] of streams) {
        const m = path.match(/^BinData[/\\]BIN(\d+)\.\w+$/i);
        if (m) binEntries.push({ binNum: parseInt(m[1], 10), data: streamData });
      }
      binEntries.sort((a, b) => a.binNum - b.binNum);
      const objectMap = /* @__PURE__ */ new Map();
      for (let idx = 0; idx < binEntries.length; idx++) {
        const { data: imgData } = binEntries[idx];
        let mimeType = "image/jpeg";
        if (imgData[0] === 137 && imgData[1] === 80) mimeType = "image/png";
        else if (imgData[0] === 71 && imgData[1] === 73) mimeType = "image/gif";
        else if (imgData[0] === 66 && imgData[1] === 77) mimeType = "image/bmp";
        const base64 = TextKit.base64Encode(imgData);
        const { wPt, hPt } = getImageDimsPt(imgData, mimeType);
        objectMap.set(idx, buildImg(base64, mimeType, wPt, hPt));
      }
      const gsoCtx = { count: 0 };
      const allContent = [];
      let pageDims = A4;
      for (let s = 0; s < 100; s++) {
        const sec = streams.get(`BodyText/Section${s}`) ?? streams.get(`Section${s}`);
        if (!sec) {
          if (s === 0) {
            const fb = findBodySection(streams);
            if (fb) {
              const r2 = parseBody(fb, compressed, di, shield, gsoCtx);
              allContent.push(...r2.content);
              if (r2.pageDims) pageDims = r2.pageDims;
            }
          }
          break;
        }
        const r = shield.guard(
          () => parseBody(sec, compressed, di, shield, gsoCtx),
          { content: [], pageDims: void 0 },
          `hwp:sec${s}`
        );
        allContent.push(...r.content);
        if (r.pageDims) pageDims = r.pageDims;
      }
      if (objectMap.size > 0) {
        injectImagesIntoContent(allContent, objectMap);
      }
      warns.push(...shield.flush());
      const content = allContent.length > 0 ? allContent : [buildPara([buildSpan("")])];
      return succeed(buildRoot({}, [buildSheet(content, pageDims)]), warns);
    } catch (e) {
      warns.push(...shield.flush());
      return fail(`HWP decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
};
function findBodySection(streams) {
  for (const [k, v] of streams)
    if (k.includes("Section") && !k.includes("Header") && !k.includes("Info")) return v;
  return void 0;
}
function getImageDimsPt(data, mime) {
  const fallback = { wPt: 72, hPt: 72 };
  try {
    if (mime === "image/png" && data.length >= 24) {
      const w = (data[16] << 24 | data[17] << 16 | data[18] << 8 | data[19]) >>> 0;
      const h = (data[20] << 24 | data[21] << 16 | data[22] << 8 | data[23]) >>> 0;
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
    }
    if (mime === "image/jpeg") {
      let i = 2;
      while (i + 8 < data.length) {
        if (data[i] !== 255) {
          i++;
          continue;
        }
        const marker = data[i + 1];
        if (marker >= 192 && marker <= 195) {
          const h = (data[i + 5] << 8 | data[i + 6]) >>> 0;
          const w = (data[i + 7] << 8 | data[i + 8]) >>> 0;
          if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
        }
        const segLen = data[i + 2] << 8 | data[i + 3];
        i += 2 + (segLen > 0 ? segLen : 2);
      }
    }
    if (mime === "image/bmp" && data.length >= 26) {
      const w = BinaryKit.readU32LE(data, 18);
      const h = Math.abs(BinaryKit.readU32LE(data, 22) | 0);
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
    }
    if (mime === "image/gif" && data.length >= 10) {
      const w = data[6] | data[7] << 8;
      const h = data[8] | data[9] << 8;
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
    }
  } catch {
  }
  return fallback;
}
function injectImagesIntoContent(content, objectMap) {
  if (objectMap.size === 0) return;
  const processKids = (kids) => {
    for (let i = 0; i < kids.length; i++) {
      const kid = kids[i];
      if (kid.tag === "span" && kid.kids && kid.kids[0]?.tag === "txt") {
        const text = kid.kids[0].content;
        const match = text.match?.(/^__(?:IMG|EXT)_(\d+)(?:_W(\d+)_H(\d+))?__$/);
        if (match) {
          const objId = parseInt(match[1], 10);
          const base = objectMap.get(objId);
          if (base) {
            const wPt = match[2] ? parseInt(match[2], 10) : 0;
            const hPt = match[3] ? parseInt(match[3], 10) : 0;
            kids[i] = wPt > 0 && hPt > 0 ? { ...base, w: wPt, h: hPt } : base;
          }
        }
      }
    }
  };
  const processGridKids = (grid) => {
    if (!grid.kids || !Array.isArray(grid.kids)) return;
    for (const row of grid.kids) {
      if (!row.kids || !Array.isArray(row.kids)) continue;
      for (const cell of row.kids) {
        if (!cell.kids || !Array.isArray(cell.kids)) continue;
        for (const cellKid of cell.kids) {
          if (cellKid.tag === "grid") {
            processGridKids(cellKid);
          } else if (cellKid.tag === "para" && cellKid.kids) {
            processKids(cellKid.kids);
          }
        }
      }
    }
  };
  for (const node of content) {
    if (node.tag === "para" && node.kids) {
      processKids(node.kids);
      for (const kid of node.kids) {
        if (kid.tag === "grid") {
          processGridKids(kid);
        }
      }
    } else if (node.tag === "grid") {
      processGridKids(node);
    }
  }
}
registry.registerDecoder(new HwpScanner());

// src/decoders/docx/DocxDecoder.ts
var DocxDecoder = class extends BaseDecoder {
  getFormat() {
    return "docx";
  }
  async decode(data) {
    const shield = new ShieldedParser();
    const warns = [];
    try {
      const files = await ArchiveKit.unzip(data);
      const getFile = (path) => {
        const lower = path.toLowerCase();
        for (const [name, data2] of files.entries()) {
          if (name.toLowerCase() === lower) return data2;
        }
        return void 0;
      };
      const docXml = getFile("word/document.xml");
      if (!docXml) return fail("DOCX: word/document.xml not found");
      const relsXml = getFile("word/_rels/document.xml.rels");
      const relsMap = relsXml ? await parseRels(TextKit.decode(relsXml)) : /* @__PURE__ */ new Map();
      const coreXml2 = getFile("docProps/core.xml");
      let meta = {};
      if (coreXml2) {
        try {
          meta = await parseCoreProps(TextKit.decode(coreXml2));
        } catch {
        }
      }
      const numXml = getFile("word/numbering.xml");
      let numMap = /* @__PURE__ */ new Map();
      if (numXml) {
        try {
          numMap = await parseNumbering(TextKit.decode(numXml));
        } catch {
        }
      }
      let stylesMap = /* @__PURE__ */ new Map();
      let paraStyleMap = /* @__PURE__ */ new Map();
      const stylesXml2 = getFile("word/styles.xml");
      if (stylesXml2) {
        try {
          const stylesStr = TextKit.decode(stylesXml2);
          stylesMap = await parseStylesMap(stylesStr);
          paraStyleMap = await parseParaStyleMap(stylesStr);
        } catch {
        }
      }
      let docStr = TextKit.decode(docXml).trim();
      if (!docStr) {
        warns.push(
          "DOCX: word/document.xml is empty, using fallback empty document"
        );
        docStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>';
      }
      const docObj = await XmlKit.parseStrict(docStr);
      const body = getBody(docObj);
      const dims = extractDims2(body) ?? { ...A4 };
      const elements = getBodyElements(body);
      console.log(
        `[DocxDecoder] \uD30C\uC2F1\uB41C \uC804\uCCB4 \uBCF8\uBB38 \uC694\uC18C \uAC1C\uC218: ${elements.length}`
      );
      const decCtx = {
        relsMap,
        files,
        shield,
        numMap,
        warns,
        stylesMap,
        paraStyleMap
      };
      const kids = [];
      for (const el of elements) {
        const nodes = shield.guard(
          () => decodeElement(el, decCtx),
          [buildPara([buildSpan("[\uC694\uC18C \uD30C\uC2F1 \uC2E4\uD328]")])],
          "docx:bodyElement"
        );
        if (Array.isArray(nodes)) {
          kids.push(...nodes);
        } else {
          kids.push(nodes);
        }
        if (el.type === "para") {
          const pPr = el.node?.["w:pPr"]?.[0] ?? el.node?.pPr?.[0] ?? {};
          const inlineSectPr = pPr?.["w:sectPr"]?.[0] ?? pPr?.sectPr?.[0];
          if (inlineSectPr) {
            const typeAttr = inlineSectPr?.["w:type"]?.[0]?._attr;
            const sectType = typeAttr?.["w:val"] ?? typeAttr?.val ?? "nextPage";
            if (sectType !== "continuous") {
              kids.push(
                buildPara([{ tag: "span", props: {}, kids: [buildPb()] }])
              );
            }
          }
        }
      }
      const headersMap = await decodeHeaderFooter2(
        "header",
        body,
        relsMap,
        files,
        decCtx
      );
      const footersMap = await decodeHeaderFooter2(
        "footer",
        body,
        relsMap,
        files,
        decCtx
      );
      warns.push(...shield.flush());
      const sheet = buildSheet(kids.filter(Boolean), dims, {
        headers: headersMap,
        footers: footersMap
      });
      return succeed(buildRoot(meta, [sheet]), warns);
    } catch (e) {
      warns.push(...shield.flush());
      return fail(`DOCX decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
};
function toArr2(v) {
  return v == null ? [] : Array.isArray(v) ? v : [v];
}
function resolveDocxPath(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + "/" + target).split("/");
  const stack = [];
  for (const p of parts) {
    if (p === "..") {
      stack.pop();
    } else if (p !== ".") {
      stack.push(p);
    }
  }
  return stack.join("/");
}
async function parseRels(xml) {
  const map = /* @__PURE__ */ new Map();
  const trimmed = xml.trim();
  if (!trimmed) return map;
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    for (const rel of toArr2(obj?.Relationships?.[0]?.Relationship)) {
      const a = rel?._attr ?? {};
      if (a.Id && a.Target) map.set(a.Id, a.Target);
    }
  } catch {
  }
  return map;
}
async function parseCoreProps(xml) {
  const trimmed = xml.trim();
  if (!trimmed) return {};
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    const c = obj?.["cp:coreProperties"]?.[0] ?? obj?.coreProperties?.[0] ?? {};
    return {
      title: c?.["dc:title"]?.[0]?._text ?? void 0,
      author: c?.["dc:creator"]?.[0]?._text ?? void 0,
      subject: c?.["dc:subject"]?.[0]?._text ?? void 0,
      created: c?.["dcterms:created"]?.[0]?._text ?? void 0,
      modified: c?.["dcterms:modified"]?.[0]?._text ?? void 0
    };
  } catch {
    return {};
  }
}
async function parseNumbering(xml) {
  const map = /* @__PURE__ */ new Map();
  const trimmed = xml.trim();
  if (!trimmed) return map;
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    const root = obj?.["w:numbering"]?.[0] ?? obj?.numbering?.[0] ?? obj;
    const absMap = /* @__PURE__ */ new Map();
    for (const abs of toArr2(root?.["w:abstractNum"] ?? root?.abstractNum)) {
      const absId = Number(
        abs?._attr?.["w:abstractNumId"] ?? abs?._attr?.abstractNumId ?? 0
      );
      const levels = /* @__PURE__ */ new Map();
      for (const lvl of toArr2(abs?.["w:lvl"] ?? abs?.lvl)) {
        const ilvl = Number(lvl?._attr?.["w:ilvl"] ?? lvl?._attr?.ilvl ?? 0);
        const fmtNode = lvl?.["w:numFmt"]?.[0]?._attr ?? lvl?.numFmt?.[0]?._attr ?? {};
        const fmt = fmtNode?.["w:val"] ?? fmtNode?.val ?? "decimal";
        levels.set(ilvl, { fmt, isOrdered: fmt !== "bullet" });
      }
      absMap.set(absId, levels);
    }
    for (const num of toArr2(root?.["w:num"] ?? root?.num)) {
      const numId = Number(num?._attr?.["w:numId"] ?? num?._attr?.numId ?? 0);
      const absRef = num?.["w:abstractNumId"]?.[0]?._attr ?? num?.abstractNumId?.[0]?._attr ?? {};
      const absId = Number(absRef?.["w:val"] ?? absRef?.val ?? 0);
      const levels = absMap.get(absId) ?? /* @__PURE__ */ new Map();
      map.set(numId, { levels });
    }
  } catch {
  }
  return map;
}
function getBody(obj) {
  const doc = obj?.["w:document"]?.[0] ?? obj?.document?.[0] ?? obj;
  const body = doc?.["w:body"]?.[0] ?? doc?.body?.[0] ?? doc;
  if (!body) {
    console.error("[DocxDecoder] \uBCF8\uBB38(body)\uC744 \uCC3E\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.");
  }
  return body;
}
function extractDims2(body) {
  try {
    const sp = body?.["w:sectPr"]?.[0] ?? body?.sectPr?.[0];
    if (!sp) return null;
    const sz = sp?.["w:pgSz"]?.[0]?._attr ?? sp?.pgSz?.[0]?._attr;
    const mar = sp?.["w:pgMar"]?.[0]?._attr ?? sp?.pgMar?.[0]?._attr;
    if (!sz) return null;
    const headerDxa = Number(mar?.["w:header"] ?? mar?.header ?? 0);
    const footerDxa = Number(mar?.["w:footer"] ?? mar?.footer ?? 0);
    return {
      wPt: Metric.dxaToPt(Number(sz["w:w"] ?? sz.w ?? 11906)),
      hPt: Metric.dxaToPt(Number(sz["w:h"] ?? sz.h ?? 16838)),
      mt: Metric.dxaToPt(Number(mar?.["w:top"] ?? mar?.top ?? 1440)),
      mb: Metric.dxaToPt(Number(mar?.["w:bottom"] ?? mar?.bottom ?? 1440)),
      ml: Metric.dxaToPt(Number(mar?.["w:left"] ?? mar?.left ?? 1800)),
      mr: Metric.dxaToPt(Number(mar?.["w:right"] ?? mar?.right ?? 1800)),
      orient: (sz["w:orient"] ?? sz.orient) === "landscape" ? "landscape" : "portrait",
      headerPt: headerDxa > 0 ? Metric.dxaToPt(headerDxa) : void 0,
      footerPt: footerDxa > 0 ? Metric.dxaToPt(footerDxa) : void 0
    };
  } catch {
    return null;
  }
}
function getBodyElements(body) {
  const paras = toArr2(body?.["w:p"] ?? body?.p);
  const tables = toArr2(body?.["w:tbl"] ?? body?.tbl);
  const sdts = toArr2(body?.["w:sdt"] ?? body?.sdt);
  const childOrder = body?.["_childOrder"];
  if (Array.isArray(childOrder)) {
    const items = [];
    let pi = 0, ti = 0, si = 0;
    for (const tag of childOrder) {
      if ((tag === "w:p" || tag === "p") && pi < paras.length) {
        items.push({ type: "para", node: paras[pi++] });
      } else if ((tag === "w:tbl" || tag === "tbl") && ti < tables.length) {
        items.push({ type: "table", node: tables[ti++] });
      } else if ((tag === "w:sdt" || tag === "sdt") && si < sdts.length) {
        items.push({ type: "sdt", node: sdts[si++] });
      }
    }
    while (pi < paras.length) items.push({ type: "para", node: paras[pi++] });
    while (ti < tables.length)
      items.push({ type: "table", node: tables[ti++] });
    while (si < sdts.length) items.push({ type: "sdt", node: sdts[si++] });
    return items;
  }
  return [
    ...paras.map((n) => ({ type: "para", node: n })),
    ...tables.map((n) => ({ type: "table", node: n })),
    ...sdts.map((n) => ({ type: "sdt", node: n }))
  ];
}
async function decodeHeaderFooter2(kind, body, relsMap, files, ctx) {
  try {
    const sp = body?.["w:sectPr"]?.[0] ?? body?.sectPr?.[0];
    if (!sp) return void 0;
    const refTag = kind === "header" ? "w:headerReference" : "w:footerReference";
    const refs = toArr2(sp?.[refTag] ?? sp?.[refTag.replace("w:", "")]);
    if (refs.length === 0) return void 0;
    const result = {};
    for (const ref of refs) {
      const type = ref._attr?.["w:type"] ?? ref._attr?.type ?? "default";
      const rId = ref._attr?.["r:id"] ?? ref._attr?.["r:Id"] ?? ref._attr?.id;
      if (!rId) continue;
      const target = relsMap.get(rId);
      if (!target) continue;
      const filePath = resolveDocxPath("word", target);
      const fileData = files.get(filePath);
      if (!fileData) continue;
      const hfFileName = filePath.split("/").pop() ?? "";
      const hfRelsPath = `word/_rels/${hfFileName}.rels`;
      const hfRelsData = files.get(hfRelsPath);
      let hfRelsMap = relsMap;
      if (hfRelsData) {
        const hfRelsStr = TextKit.decode(hfRelsData).trim();
        const parsed = hfRelsStr ? await parseRels(hfRelsStr) : /* @__PURE__ */ new Map();
        hfRelsMap = new Map([...relsMap, ...parsed]);
      }
      const xmlStr = TextKit.decode(fileData).trim();
      if (!xmlStr) continue;
      const watermark = extractWatermark(xmlStr);
      if (watermark) {
        result[type] = [
          buildPara([
            buildSpan(watermark, { pt: 80, color: "CCCCCC", b: true })
          ])
        ];
        continue;
      }
      try {
        const obj = await XmlKit.parseStrict(xmlStr);
        const rootTag = kind === "header" ? "w:hdr" : "w:ftr";
        const root = obj?.[rootTag]?.[0] ?? obj?.[rootTag.replace("w:", "")]?.[0] ?? obj;
        const origRelsMap = ctx.relsMap;
        ctx.relsMap = hfRelsMap;
        const paras = toArr2(root?.["w:p"] ?? root?.p);
        result[type] = paras.map((p) => decodePara2(p, ctx));
        ctx.relsMap = origRelsMap;
      } catch (err) {
        console.warn(`[DocxDecoder] ${kind} (${type}) XML \uD30C\uC2F1 \uC2E4\uD328:`, err);
        continue;
      }
    }
    return Object.keys(result).length > 0 ? result : void 0;
  } catch {
    return void 0;
  }
}
function extractWatermark(xml) {
  if (!xml.includes("v:textpath")) return null;
  const m = xml.match(/string="([^"]+)"/);
  return m ? m[1] : null;
}
function hasDrawingDeep(node) {
  if (!node || typeof node !== "object") return false;
  if (node["w:drawing"] || node["w:pict"]) return true;
  return Object.values(node).some((v) => {
    if (Array.isArray(v)) return v.some(hasDrawingDeep);
    return hasDrawingDeep(v);
  });
}
function decodeElement(el, ctx) {
  if (el.type === "table") {
    const { value } = ctx.shield.guardGrid(
      el.node,
      (n) => decodeGrid2(n, ctx),
      (n) => decodeGridSimple2(n),
      (n) => decodeGridFlat2(n),
      (n) => decodeGridText2(n),
      "docx:table"
    );
    return value;
  } else if (el.type === "sdt") {
    return decodeSdt(el.node, ctx);
  }
  return decodePara2(el.node, ctx);
}
function decodeSdt(sdt, ctx) {
  const content = sdt?.["w:sdtContent"]?.[0] ?? sdt?.sdtContent?.[0];
  if (!content) return [];
  const elements = getBodyElements(content);
  const kids = [];
  for (const el of elements) {
    const res = decodeElement(el, ctx);
    if (Array.isArray(res)) kids.push(...res);
    else kids.push(res);
  }
  return kids;
}
function decodePara2(p, ctx) {
  const pPr = p?.["w:pPr"]?.[0] ?? {};
  const alignVal = pPr?.["w:jc"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:jc"]?.[0]?._attr?.val;
  const headStyle = pPr?.["w:pStyle"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:pStyle"]?.[0]?._attr?.val ?? "";
  const styleInherited = resolveParaStyle(
    headStyle || void 0,
    ctx.paraStyleMap
  );
  const props = {
    align: safeAlign(alignVal),
    heading: parseHeading(headStyle),
    styleId: headStyle || void 0
  };
  const spacingAttr = pPr?.["w:spacing"]?.[0]?._attr ?? pPr?.spacing?.[0]?._attr ?? {};
  const beforeVal = Number(
    spacingAttr?.["w:before"] ?? spacingAttr?.before ?? 0
  );
  const afterVal = Number(spacingAttr?.["w:after"] ?? spacingAttr?.after ?? 0);
  const lineVal = Number(spacingAttr?.["w:line"] ?? spacingAttr?.line ?? 0);
  const lineRule = spacingAttr?.["w:lineRule"] ?? spacingAttr?.lineRule ?? "auto";
  if (beforeVal > 0) props.spaceBefore = Metric.dxaToPt(beforeVal);
  else if (styleInherited.pPr?.spaceBefore)
    props.spaceBefore = styleInherited.pPr.spaceBefore;
  if (afterVal > 0) props.spaceAfter = Metric.dxaToPt(afterVal);
  else if (styleInherited.pPr?.spaceAfter)
    props.spaceAfter = styleInherited.pPr.spaceAfter;
  if (lineVal > 0 && lineRule === "auto") props.lineHeight = lineVal / 240;
  else if (styleInherited.pPr?.lineHeight)
    props.lineHeight = styleInherited.pPr.lineHeight;
  const indAttr = pPr?.["w:ind"]?.[0]?._attr ?? pPr?.ind?.[0]?._attr ?? {};
  const leftVal = Number(indAttr?.["w:left"] ?? indAttr?.left ?? 0);
  const rightVal = Number(indAttr?.["w:right"] ?? indAttr?.right ?? 0);
  const firstLineVal = Number(
    indAttr?.["w:firstLine"] ?? indAttr?.firstLine ?? 0
  );
  const hangingVal = Number(indAttr?.["w:hanging"] ?? indAttr?.hanging ?? 0);
  if (leftVal > 0) props.indentPt = Metric.dxaToPt(leftVal);
  else if (styleInherited.pPr?.indentPt)
    props.indentPt = styleInherited.pPr.indentPt;
  if (rightVal > 0) props.indentRightPt = Metric.dxaToPt(rightVal);
  else if (styleInherited.pPr?.indentRightPt)
    props.indentRightPt = styleInherited.pPr.indentRightPt;
  if (firstLineVal > 0) props.firstLineIndentPt = Metric.dxaToPt(firstLineVal);
  else if (hangingVal > 0)
    props.firstLineIndentPt = -Metric.dxaToPt(hangingVal);
  else if (styleInherited.pPr?.firstLineIndentPt)
    props.firstLineIndentPt = styleInherited.pPr.firstLineIndentPt;
  if (!alignVal && styleInherited.pPr?.align)
    props.align = safeAlign(styleInherited.pPr.align);
  const numPr = pPr?.["w:numPr"]?.[0] ?? pPr?.numPr?.[0];
  if (numPr) {
    const ilvlNode = numPr?.["w:ilvl"]?.[0]?._attr ?? numPr?.ilvl?.[0]?._attr ?? {};
    const numIdNode = numPr?.["w:numId"]?.[0]?._attr ?? numPr?.numId?.[0]?._attr ?? {};
    const ilvl = Number(ilvlNode?.["w:val"] ?? ilvlNode?.val ?? 0);
    const numId = Number(numIdNode?.["w:val"] ?? numIdNode?.val ?? 0);
    props.listLv = ilvl;
    const numEntry = ctx.numMap.get(numId);
    if (numEntry) {
      const lvlInfo = numEntry.levels.get(ilvl) ?? numEntry.levels.get(0);
      props.listOrd = lvlInfo?.isOrdered ?? false;
    } else {
      props.listOrd = numId >= 2;
    }
  }
  const pbBeforeNode = pPr?.["w:pageBreakBefore"]?.[0] ?? pPr?.pageBreakBefore?.[0];
  const hasPageBreakBefore = pbBeforeNode != null && (pbBeforeNode?._attr?.["w:val"] ?? pbBeforeNode?._attr?.val ?? "1") !== "0";
  const children = p?.["_childOrder"];
  const kids = [];
  if (Array.isArray(children)) {
    const runsArr = toArr2(p?.["w:r"] ?? p?.r);
    const hlArr = toArr2(p?.["w:hyperlink"] ?? p?.hyperlink);
    const sdtArr = toArr2(p?.["w:sdt"] ?? p?.sdt);
    let ri = 0;
    let hi = 0;
    let si = 0;
    for (const tag of children) {
      if (tag === "w:r" || tag === "r") {
        const run = runsArr[ri++];
        if (run) {
          kids.push(
            ctx.shield.guard(
              () => hasDrawingDeep(run) ? decodeRunOrImage(run, ctx) : decodeRun(run, ctx, styleInherited.rPr),
              buildSpan(""),
              "docx:run"
            )
          );
        }
      } else if (tag === "w:hyperlink" || tag === "hyperlink") {
        const hl = hlArr[hi++];
        if (hl) {
          const rId = hl?._attr?.["r:id"] ?? hl?._attr?.id;
          const url = rId ? ctx.relsMap.get(rId) : "";
          const hlRuns = toArr2(hl?.["w:r"] ?? hl?.r);
          const hlKids = hlRuns.map(
            (r) => decodeRun(r, ctx, {
              ...styleInherited.rPr,
              u: true,
              color: "0000FF"
            })
          );
          kids.push({
            tag: "link",
            href: url || "",
            kids: hlKids
          });
        }
      } else if (tag === "w:sdt" || tag === "sdt") {
        const sdt = sdtArr[si++];
        if (sdt) {
          const sdtContent = sdt?.["w:sdtContent"]?.[0] ?? sdt?.sdtContent?.[0];
          if (sdtContent) {
            const innerRuns = toArr2(sdtContent?.["w:r"] ?? sdtContent?.r);
            for (const ir of innerRuns) {
              kids.push(
                ctx.shield.guard(
                  () => hasDrawingDeep(ir) ? decodeRunOrImage(ir, ctx) : decodeRun(ir, ctx, styleInherited.rPr),
                  buildSpan(""),
                  "docx:run"
                )
              );
            }
          }
        }
      }
    }
  } else {
    const runs = toArr2(p?.["w:r"] ?? p?.r);
    const legacyKids = ctx.shield.guardAll(
      runs,
      (run) => hasDrawingDeep(run) ? decodeRunOrImage(run, ctx) : decodeRun(run, ctx, styleInherited.rPr),
      () => buildSpan(""),
      "docx:run"
    );
    kids.push(...legacyKids);
  }
  const filteredKids = kids.filter(Boolean);
  if (hasPageBreakBefore) {
    filteredKids.unshift({ tag: "span", props: {}, kids: [buildPb()] });
  }
  return buildPara(filteredKids, props);
}
function decodeRunOrImage(run, ctx) {
  function findFirstDrawing(node) {
    if (!node || typeof node !== "object") return null;
    if (node["w:drawing"]) return node["w:drawing"][0];
    if (node["w:pict"]) return node["w:pict"][0];
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) {
        for (const v of value) {
          const found = findFirstDrawing(v);
          if (found) return found;
        }
      } else {
        const found = findFirstDrawing(value);
        if (found) return found;
      }
    }
    return null;
  }
  const drawing = findFirstDrawing(run);
  if (drawing) {
    const img = decodeDrawing(drawing, ctx);
    if (img) return img;
  }
  return decodeRun(run, ctx);
}
function decodeImageLayout(anchor) {
  const wrap = anchor?.["wp:wrapTop"]?.[0] ?? anchor?.wrapTop?.[0];
  const anchorPos = anchor?.["wp:anchorPos"]?.[0]?._attr ?? anchor?.anchorPos?.[0]?._attr ?? {};
  const layout = {
    wrap: "square",
    horzAlign: "left",
    vertAlign: "top",
    horzRelTo: "page",
    vertRelTo: "page",
    xPt: Number(anchorPos?.x ?? 0) / 12700,
    // emu to pt
    yPt: Number(anchorPos?.y ?? 0) / 12700
    // emu to pt
  };
  if (wrap?.["wp:none"]) layout.wrap = "none";
  else if (wrap?.["wp:square"]) layout.wrap = "square";
  else if (wrap?.["wp:tight"]) layout.wrap = "tight";
  else if (wrap?.["wp:through"]) layout.wrap = "through";
  else if (wrap?.["wp:behind"]) layout.wrap = "behind";
  else if (wrap?.["wp:inFront"]) layout.wrap = "front";
  return layout;
}
function decodeDrawing(drawing, ctx) {
  try {
    const inline = drawing?.["wp:inline"]?.[0] ?? drawing?.inline?.[0];
    const anchor = drawing?.["wp:anchor"]?.[0] ?? drawing?.anchor?.[0];
    const container = inline ?? anchor;
    if (!container) return null;
    const extent = container?.["wp:extent"]?.[0]?._attr ?? container?.extent?.[0]?._attr ?? {};
    const cx = Number(extent?.cx ?? 0);
    const cy = Number(extent?.cy ?? 0);
    const wPt = Metric.emuToPt(cx);
    const hPt = Metric.emuToPt(cy);
    const docPr = container?.["wp:docPr"]?.[0]?._attr ?? container?.docPr?.[0]?._attr ?? {};
    const alt = docPr?.descr ?? docPr?.name ?? "";
    const graphic = container?.["a:graphic"]?.[0] ?? container?.graphic?.[0];
    const graphicData = graphic?.["a:graphicData"]?.[0] ?? graphic?.graphicData?.[0];
    if (graphicData?.["c:chart"] || graphicData?.chart) {
      return {
        tag: "img",
        b64: "",
        // 플레이스홀더
        mime: "image/png",
        w: wPt,
        h: hPt,
        alt: `[\uCC28\uD2B8: ${alt || "\uCC28\uD2B8"}]`,
        layout: decodeImageLayout(anchor)
      };
    }
    const pic = graphicData?.["pic:pic"]?.[0] ?? graphicData?.pic?.[0];
    const blipFill = pic?.["pic:blipFill"]?.[0] ?? pic?.blipFill?.[0];
    const blip = blipFill?.["a:blip"]?.[0]?._attr ?? blipFill?.blip?.[0]?._attr ?? {};
    const rId = blip?.["r:embed"] ?? blip?.embed;
    if (!rId) return null;
    const target = ctx.relsMap.get(rId);
    if (!target) return null;
    let filePath = resolveDocxPath("word", target);
    let fileData = ctx.files.get(filePath);
    if (!fileData) {
      filePath = resolveDocxPath("word/_rels", target);
      fileData = ctx.files.get(filePath);
    }
    if (!fileData) {
      const fileName = target.split("/").pop() ?? "";
      for (const [k, v] of ctx.files) {
        if (fileName && (k.endsWith("/" + fileName) || k === fileName)) {
          fileData = v;
          filePath = k;
          break;
        }
      }
    }
    if (!fileData) {
      console.warn(`[DocxDecoder] image not found: "${target}"`);
      return null;
    }
    const ext = target.split(".").pop()?.toLowerCase() ?? "png";
    const mimeMap = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp"
    };
    const mime = mimeMap[ext] ?? "image/png";
    console.log(
      `[DocxDecoder] image loaded: ${filePath} (${mime}, ${fileData.length} bytes)`
    );
    const layout = inline ? { wrap: "inline" } : extractAnchorLayout(anchor);
    return buildImg(
      TextKit.base64Encode(fileData),
      mime,
      wPt,
      hPt,
      alt || void 0,
      layout
    );
  } catch {
    return null;
  }
}
var HIGHLIGHT_COLOR_MAP = {
  yellow: "FFFF00",
  green: "00FF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
  blue: "0000FF",
  red: "FF0000",
  darkBlue: "00008B",
  darkCyan: "008B8B",
  darkGreen: "006400",
  darkMagenta: "8B008B",
  darkRed: "8B0000",
  darkYellow: "808000",
  darkGray: "A9A9A9",
  lightGray: "D3D3D3",
  black: "000000",
  white: "FFFFFF"
};
function decodeRun(run, ctx, styleRpr) {
  const rPr = run?.["w:rPr"]?.[0] ?? run?.rPr?.[0] ?? {};
  const vanishNode = rPr?.["w:vanish"]?.[0] ?? rPr?.vanish?.[0];
  if (vanishNode != null) {
    const vanishVal = vanishNode?._attr?.["w:val"] ?? vanishNode?._attr?.val ?? "1";
    if (vanishVal !== "0") return buildSpan("");
  }
  const szAttr = rPr?.["w:sz"]?.[0]?._attr ?? rPr?.sz?.[0]?._attr ?? {};
  const szVal = szAttr?.["w:val"] ?? szAttr?.val;
  const szCsAttr = rPr?.["w:szCs"]?.[0]?._attr ?? rPr?.szCs?.[0]?._attr ?? {};
  const szCsVal = szCsAttr?.["w:val"] ?? szCsAttr?.val;
  const effectiveSzVal = szVal ?? szCsVal;
  const colorAttr = rPr?.["w:color"]?.[0]?._attr ?? rPr?.color?.[0]?._attr ?? {};
  const colorVal = colorAttr?.["w:val"] ?? colorAttr?.val;
  const fontAttr = rPr?.["w:rFonts"]?.[0]?._attr ?? rPr?.rFonts?.[0]?._attr ?? {};
  const fontName = fontAttr?.["w:ascii"] ?? fontAttr?.ascii ?? fontAttr?.["w:hAnsi"] ?? fontAttr?.hAnsi ?? fontAttr?.["w:eastAsia"] ?? fontAttr?.eastAsia;
  const underVal = rPr?.["w:u"]?.[0]?._attr?.["w:val"] ?? rPr?.["w:u"]?.[0]?._attr?.val;
  const shdAttr = rPr?.["w:shd"]?.[0]?._attr ?? rPr?.shd?.[0]?._attr ?? {};
  const shdBg = safeHex(shdAttr?.["w:fill"] ?? shdAttr?.fill);
  const hlAttr = rPr?.["w:highlight"]?.[0]?._attr ?? rPr?.highlight?.[0]?._attr ?? {};
  const hlVal = hlAttr?.["w:val"] ?? hlAttr?.val;
  const bgVal = (hlVal ? HIGHLIGHT_COLOR_MAP[hlVal] : void 0) ?? shdBg;
  const vertAlignVal = rPr?.["w:vertAlign"]?.[0]?._attr?.["w:val"] ?? rPr?.["w:vertAlign"]?.[0]?._attr?.val;
  const posAttr = rPr?.["w:position"]?.[0]?._attr ?? rPr?.position?.[0]?._attr ?? {};
  const posVal = Number(posAttr?.["w:val"] ?? posAttr?.val ?? 0);
  let isSup = vertAlignVal === "superscript";
  let isSub = vertAlignVal === "subscript";
  if (!isSup && !isSub && posVal !== 0) {
    if (posVal >= 4) isSup = true;
    else if (posVal <= -4) isSub = true;
  }
  const bNode = rPr?.["w:b"]?.[0] ?? rPr?.b?.[0];
  const isBold = bNode != null && (bNode?._attr?.["w:val"] ?? bNode?._attr?.val ?? "1") !== "0";
  const iNode = rPr?.["w:i"]?.[0] ?? rPr?.i?.[0];
  const isItalic = iNode != null && (iNode?._attr?.["w:val"] ?? iNode?._attr?.val ?? "1") !== "0";
  const sNode = rPr?.["w:strike"]?.[0] ?? rPr?.strike?.[0];
  const isStrike = sNode != null && (sNode?._attr?.["w:val"] ?? sNode?._attr?.val ?? "1") !== "0";
  const props = {
    b: (bNode != null ? isBold : styleRpr?.b) || void 0,
    i: (iNode != null ? isItalic : styleRpr?.i) || void 0,
    u: (underVal ? underVal !== "none" : styleRpr?.u) || void 0,
    s: (sNode != null ? isStrike : styleRpr?.s) || void 0,
    sup: isSup || void 0,
    sub: isSub || void 0,
    pt: effectiveSzVal ? Metric.halfPtToPt(Number(effectiveSzVal)) : styleRpr?.pt,
    color: safeHex(colorVal) ?? styleRpr?.color,
    font: fontName ? safeFont(fontName) : styleRpr?.font,
    bg: bgVal
  };
  const fldChar = run?.["w:fldChar"]?.[0]?._attr ?? run?.fldChar?.[0]?._attr;
  const instrText = run?.["w:instrText"]?.[0];
  const brNodes = toArr2(run?.["w:br"] ?? run?.br ?? []);
  for (const br of brNodes) {
    const brType = br?._attr?.["w:type"] ?? br?._attr?.type;
    if (brType === "page") {
      return { tag: "span", props, kids: [buildPb()] };
    }
  }
  const textNodes = toArr2(run?.["w:t"] ?? run?.t);
  const content = textNodes.map((t) => typeof t === "string" ? t : t?._ ?? t?._text ?? "").join("");
  if (instrText) {
    const instrStr = typeof instrText === "string" ? instrText : instrText?._text ?? "";
    if (instrStr.trim().toUpperCase() === "PAGE") {
      const pageNum = { tag: "pagenum", format: "decimal" };
      return { tag: "span", props, kids: [pageNum] };
    }
  }
  return buildSpan(content, props);
}
function parseBorderDef(bdrNode) {
  const sides = [
    ["top", "top"],
    ["bottom", "bottom"],
    ["left", "left"],
    ["right", "right"],
    ["insideH", "insideH"],
    ["insideV", "insideV"]
  ];
  const result = {};
  for (const [xml, prop] of sides) {
    const bdr = bdrNode?.["w:" + xml]?.[0]?._attr ?? bdrNode?.[xml]?.[0]?._attr;
    if (!bdr) continue;
    const val = bdr?.["w:val"] ?? bdr?.val;
    if (val === "none" || val === "nil") continue;
    result[prop] = safeStrokeDocx(
      val,
      Number(bdr?.["w:sz"] ?? bdr?.sz ?? 4),
      bdr?.["w:color"] ?? bdr?.color
    );
  }
  return result;
}
async function parseStylesMap(xml) {
  const map = /* @__PURE__ */ new Map();
  const trimmed = xml.trim();
  if (!trimmed) return map;
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    const stylesRoot = obj?.["w:styles"]?.[0] ?? obj?.styles?.[0] ?? obj;
    const styleArr = toArr2(stylesRoot?.["w:style"] ?? stylesRoot?.style);
    for (const style of styleArr) {
      const attr = style?._attr ?? {};
      const type = attr?.["w:type"] ?? attr?.type;
      if (type !== "table") continue;
      const id = attr?.["w:styleId"] ?? attr?.styleId;
      if (!id) continue;
      const tblPr = style?.["w:tblPr"]?.[0] ?? style?.tblPr?.[0];
      const tblBdrNode = tblPr?.["w:tblBorders"]?.[0] ?? tblPr?.tblBorders?.[0];
      const tblBorders = tblBdrNode ? parseBorderDef(tblBdrNode) : void 0;
      const tcStyle = style?.["w:tcStyle"]?.[0] ?? style?.tcStyle?.[0];
      const tcBdrNode = tcStyle?.["w:tcBdr"]?.[0] ?? tcStyle?.tcBdr?.[0];
      if (tcBdrNode) {
        const cellDef = parseBorderDef(tcBdrNode);
        if (!tblBorders) {
          map.set(id, { tblBorders: cellDef });
        } else {
          map.set(id, { tblBorders: { ...cellDef, ...tblBorders } });
        }
      } else if (tblBorders) {
        map.set(id, { tblBorders });
      }
    }
  } catch {
  }
  return map;
}
async function parseParaStyleMap(xml) {
  const map = /* @__PURE__ */ new Map();
  const trimmed = xml.trim();
  if (!trimmed) return map;
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    const stylesRoot = obj?.["w:styles"]?.[0] ?? obj?.styles?.[0] ?? obj;
    const styleArr = toArr2(stylesRoot?.["w:style"] ?? stylesRoot?.style);
    for (const style of styleArr) {
      const attr = style?._attr ?? {};
      const type = attr?.["w:type"] ?? attr?.type;
      if (type !== "paragraph" && type !== "character") continue;
      const id = attr?.["w:styleId"] ?? attr?.styleId;
      if (!id) continue;
      const basedOn = (style?.["w:basedOn"]?.[0]?._attr ?? style?.basedOn?.[0]?._attr)?.["w:val"];
      const def = { basedOn };
      const rPr = style?.["w:rPr"]?.[0] ?? style?.rPr?.[0];
      if (rPr) {
        const szAttr = rPr?.["w:sz"]?.[0]?._attr ?? rPr?.sz?.[0]?._attr ?? {};
        const szVal = szAttr?.["w:val"] ?? szAttr?.val;
        const colorAttr = rPr?.["w:color"]?.[0]?._attr ?? rPr?.color?.[0]?._attr ?? {};
        const colorVal = colorAttr?.["w:val"] ?? colorAttr?.val;
        const fontAttr = rPr?.["w:rFonts"]?.[0]?._attr ?? rPr?.rFonts?.[0]?._attr ?? {};
        const fontName = fontAttr?.["w:ascii"] ?? fontAttr?.ascii ?? fontAttr?.["w:eastAsia"] ?? fontAttr?.eastAsia;
        const bNode = rPr?.["w:b"]?.[0] ?? rPr?.b?.[0];
        const isBold = bNode != null && (bNode?._attr?.["w:val"] ?? bNode?._attr?.val ?? "1") !== "0";
        const iNode = rPr?.["w:i"]?.[0] ?? rPr?.i?.[0];
        const isItalic = iNode != null && (iNode?._attr?.["w:val"] ?? iNode?._attr?.val ?? "1") !== "0";
        const underVal = rPr?.["w:u"]?.[0]?._attr?.["w:val"] ?? rPr?.["w:u"]?.[0]?._attr?.val;
        const sNode = rPr?.["w:strike"]?.[0] ?? rPr?.strike?.[0];
        const isStrike = sNode != null && (sNode?._attr?.["w:val"] ?? sNode?._attr?.val ?? "1") !== "0";
        def.rPr = {
          b: isBold || void 0,
          i: isItalic || void 0,
          u: underVal && underVal !== "none" ? true : void 0,
          s: isStrike || void 0,
          pt: szVal ? Metric.halfPtToPt(Number(szVal)) : void 0,
          color: safeHex(colorVal),
          font: fontName ? safeFont(fontName) : void 0
        };
      }
      const pPr = style?.["w:pPr"]?.[0] ?? style?.pPr?.[0];
      if (pPr) {
        const spacingAttr = pPr?.["w:spacing"]?.[0]?._attr ?? pPr?.spacing?.[0]?._attr ?? {};
        const beforeVal = Number(
          spacingAttr?.["w:before"] ?? spacingAttr?.before ?? 0
        );
        const afterVal = Number(
          spacingAttr?.["w:after"] ?? spacingAttr?.after ?? 0
        );
        const lineVal = Number(
          spacingAttr?.["w:line"] ?? spacingAttr?.line ?? 0
        );
        const lineRule = spacingAttr?.["w:lineRule"] ?? spacingAttr?.lineRule ?? "auto";
        const indAttr = pPr?.["w:ind"]?.[0]?._attr ?? pPr?.ind?.[0]?._attr ?? {};
        const leftVal = Number(indAttr?.["w:left"] ?? indAttr?.left ?? 0);
        const rightVal = Number(indAttr?.["w:right"] ?? indAttr?.right ?? 0);
        const firstLineVal = Number(
          indAttr?.["w:firstLine"] ?? indAttr?.firstLine ?? 0
        );
        const hangingVal = Number(
          indAttr?.["w:hanging"] ?? indAttr?.hanging ?? 0
        );
        const alignVal = pPr?.["w:jc"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:jc"]?.[0]?._attr?.val;
        def.pPr = {
          align: alignVal,
          spaceBefore: beforeVal > 0 ? Metric.dxaToPt(beforeVal) : void 0,
          spaceAfter: afterVal > 0 ? Metric.dxaToPt(afterVal) : void 0,
          lineHeight: lineVal > 0 && lineRule === "auto" ? lineVal / 240 : void 0,
          indentPt: leftVal > 0 ? Metric.dxaToPt(leftVal) : void 0,
          indentRightPt: rightVal > 0 ? Metric.dxaToPt(rightVal) : void 0,
          firstLineIndentPt: firstLineVal > 0 ? Metric.dxaToPt(firstLineVal) : hangingVal > 0 ? -Metric.dxaToPt(hangingVal) : void 0
        };
      }
      map.set(id, def);
    }
  } catch {
  }
  return map;
}
function resolveParaStyle(styleId, map) {
  let merged = {};
  const visited = /* @__PURE__ */ new Set();
  let cur = styleId;
  while (cur && !visited.has(cur)) {
    visited.add(cur);
    const def = map.get(cur);
    if (!def) break;
    if (def.rPr) {
      merged.rPr = { ...def.rPr, ...merged.rPr };
    }
    if (def.pPr) {
      merged.pPr = { ...def.pPr, ...merged.pPr };
    }
    cur = def.basedOn;
  }
  return merged;
}
function resolveCellBorders(cp, ri, ci, rs, cs, rowCount, colCount, tblBdr) {
  const isTopEdge = ri === 0;
  const isBottomEdge = ri + rs >= rowCount;
  const isLeftEdge = ci === 0;
  const isRightEdge = ci + cs >= colCount;
  const resolved = { ...cp };
  if (!resolved.top) resolved.top = isTopEdge ? tblBdr.top : tblBdr.insideH;
  if (!resolved.bot)
    resolved.bot = isBottomEdge ? tblBdr.bottom : tblBdr.insideH;
  if (!resolved.left) resolved.left = isLeftEdge ? tblBdr.left : tblBdr.insideV;
  if (!resolved.right)
    resolved.right = isRightEdge ? tblBdr.right : tblBdr.insideV;
  return resolved;
}
function decodeGrid2(tbl, ctx) {
  const tblPr = tbl?.["w:tblPr"]?.[0] ?? tbl?.tblPr?.[0] ?? {};
  const tblLookAttr = tblPr?.["w:tblLook"]?.[0]?._attr ?? tblPr?.tblLook?.[0]?._attr ?? {};
  const look = {
    firstRow: tblLookAttr?.["w:firstRow"] === "1" || void 0,
    lastRow: tblLookAttr?.["w:lastRow"] === "1" || void 0,
    firstCol: tblLookAttr?.["w:firstColumn"] === "1" || tblLookAttr?.["w:firstCol"] === "1" || void 0,
    lastCol: tblLookAttr?.["w:lastColumn"] === "1" || tblLookAttr?.["w:lastCol"] === "1" || void 0,
    bandedRows: tblLookAttr?.["w:noHBand"] === "0" || void 0,
    bandedCols: tblLookAttr?.["w:noVBand"] === "0" || void 0
  };
  const tblStyleId = (tblPr?.["w:tblStyle"]?.[0]?._attr ?? tblPr?.tblStyle?.[0]?._attr)?.["w:val"];
  const styleDef = tblStyleId ? ctx.stylesMap.get(tblStyleId) : void 0;
  let tblBdr = styleDef?.tblBorders ?? {};
  const tblBordersNode = tblPr?.["w:tblBorders"]?.[0] ?? tblPr?.tblBorders?.[0];
  if (tblBordersNode) {
    const parsed = parseBorderDef(tblBordersNode);
    tblBdr = { ...tblBdr, ...parsed };
  }
  const defaultStroke = tblBdr.insideH ?? tblBdr.top;
  const gridProps = { look, defaultStroke };
  const tblGrid = tbl?.["w:tblGrid"]?.[0] ?? tbl?.tblGrid?.[0];
  if (tblGrid) {
    const gridCols = toArr2(tblGrid?.["w:gridCol"] ?? tblGrid?.gridCol ?? []);
    const colWidthsPt = gridCols.map(
      (gc) => Metric.dxaToPt(Number(gc?._attr?.["w:w"] ?? gc?._attr?.w ?? 0))
    ).filter((w) => w > 0);
    if (colWidthsPt.length > 0) gridProps.colWidths = colWidthsPt;
  }
  const rowArr = toArr2(tbl?.["w:tr"] ?? tbl?.tr);
  const rawGrid = rowArr.map((row) => {
    const cellArr = toArr2(row?.["w:tc"] ?? row?.tc);
    return cellArr.map((cell) => {
      const tcPr = cell?.["w:tcPr"]?.[0] ?? {};
      const gridSpan = Number(tcPr?.["w:gridSpan"]?.[0]?._attr?.["w:val"] ?? 1);
      const vMergeNode = tcPr?.["w:vMerge"]?.[0];
      const vMergeVal = vMergeNode?._attr?.["w:val"] ?? vMergeNode?._attr?.val;
      const vMergeRestart = vMergeVal === "restart";
      const vMergeContinue = vMergeNode != null && !vMergeRestart;
      return { cell, gridSpan, vMergeRestart, vMergeContinue };
    });
  });
  const rsMap = /* @__PURE__ */ new Map();
  for (let ri = 0; ri < rawGrid.length; ri++) {
    let gridCol = 0;
    for (let ci = 0; ci < rawGrid[ri].length; ci++) {
      const rc = rawGrid[ri][ci];
      if (rc.vMergeRestart) {
        let span = 1;
        for (let nr = ri + 1; nr < rawGrid.length; nr++) {
          let col = 0;
          let found = false;
          for (const nc of rawGrid[nr]) {
            if (col === gridCol && nc.vMergeContinue) {
              span++;
              found = true;
              break;
            }
            col += nc.gridSpan;
          }
          if (!found) break;
        }
        rsMap.set(`${ri},${ci}`, span);
      }
      gridCol += rc.gridSpan;
    }
  }
  const rowNodes = rawGrid.map((rawRow, ri) => {
    const row = rowArr[ri];
    const trPr = row?.["w:trPr"]?.[0] ?? row?.trPr?.[0] ?? {};
    const isHeaderRow = trPr?.["w:tblHeader"]?.[0] != null || trPr?.tblHeader?.[0] != null;
    if (ri === 0 && isHeaderRow) gridProps.headerRow = true;
    let rowHeightPt;
    const trHAttr = trPr?.["w:trHeight"]?.[0]?._attr ?? trPr?.trHeight?.[0]?._attr;
    if (trHAttr) {
      const hDxa = Number(trHAttr?.["w:val"] ?? trHAttr?.val ?? 0);
      if (hDxa > 0) rowHeightPt = Metric.dxaToPt(hDxa);
    }
    const cellNodes = [];
    for (let ci = 0; ci < rawRow.length; ci++) {
      const rc = rawRow[ci];
      if (rc.vMergeContinue) continue;
      const cell = rc.cell;
      const tcPr = cell?.["w:tcPr"]?.[0] ?? {};
      const bgAttr = tcPr?.["w:shd"]?.[0]?._attr ?? {};
      const bg = safeHex(bgAttr?.["w:fill"] ?? bgAttr?.fill);
      const tcBordersNode = tcPr?.["w:tcBorders"]?.[0] ?? tcPr?.tcBorders?.[0];
      const cp = { bg, isHeader: isHeaderRow || void 0 };
      if (tcBordersNode) {
        const dirs = [
          ["top", "top"],
          ["bottom", "bot"],
          ["left", "left"],
          ["right", "right"]
        ];
        for (const [xmlTag, propKey] of dirs) {
          const bdr = tcBordersNode?.["w:" + xmlTag]?.[0]?._attr ?? tcBordersNode?.[xmlTag]?.[0]?._attr;
          if (!bdr) continue;
          const val = bdr?.["w:val"] ?? bdr?.val;
          if (val === "none" || val === "nil") {
          } else {
            cp[propKey] = safeStrokeDocx(
              val,
              Number(bdr?.["w:sz"] ?? bdr?.sz ?? 4),
              bdr?.["w:color"] ?? bdr?.color
            );
          }
        }
      }
      const vaAttr = tcPr?.["w:vAlign"]?.[0]?._attr ?? tcPr?.vAlign?.[0]?._attr ?? {};
      const vaVal = vaAttr?.["w:val"] ?? vaAttr?.val;
      if (vaVal) {
        const vaMap = {
          top: "top",
          center: "mid",
          bottom: "bot"
        };
        cp.va = vaMap[vaVal];
      }
      const tcMar = tcPr?.["w:tcMar"]?.[0] ?? tcPr?.tcMar?.[0];
      if (tcMar) {
        const top = tcMar?.["w:top"]?.[0]?._attr ?? tcMar?.top?.[0]?._attr;
        const bot = tcMar?.["w:bottom"]?.[0]?._attr ?? tcMar?.bottom?.[0]?._attr;
        const left = tcMar?.["w:left"]?.[0]?._attr ?? tcMar?.left?.[0]?._attr;
        const right = tcMar?.["w:right"]?.[0]?._attr ?? tcMar?.right?.[0]?._attr;
        if (top) cp.padT = Metric.dxaToPt(Number(top?.["w:w"] ?? top?.w ?? 0));
        if (bot) cp.padB = Metric.dxaToPt(Number(bot?.["w:w"] ?? bot?.w ?? 0));
        if (left)
          cp.padL = Metric.dxaToPt(Number(left?.["w:w"] ?? left?.w ?? 0));
        if (right)
          cp.padR = Metric.dxaToPt(Number(right?.["w:w"] ?? right?.w ?? 0));
      }
      const rs = rsMap.get(`${ri},${ci}`) ?? 1;
      let gridColIdx = 0;
      for (let prevCi = 0; prevCi < ci; prevCi++) {
        if (!rawRow[prevCi].vMergeContinue)
          gridColIdx += rawRow[prevCi].gridSpan;
      }
      const colCount = gridProps.colWidths?.length ?? rawGrid[0]?.reduce((s, c) => s + c.gridSpan, 0) ?? 1;
      const resolvedCp = resolveCellBorders(
        cp,
        ri,
        gridColIdx,
        rs,
        rc.gridSpan,
        rawGrid.length,
        colCount,
        tblBdr
      );
      const paras = toArr2(cell?.["w:p"] ?? cell?.p).map(
        (p) => decodePara2(p, ctx)
      );
      cellNodes.push(
        buildCell(paras.length > 0 ? paras : [buildPara([buildSpan("")])], {
          cs: rc.gridSpan,
          rs,
          props: resolvedCp
        })
      );
    }
    return buildRow(cellNodes, rowHeightPt);
  });
  return buildGrid(rowNodes, gridProps);
}
function decodeGridSimple2(tbl) {
  const rowArr = toArr2(tbl?.["w:tr"] ?? tbl?.tr);
  const rowNodes = rowArr.map((row) => {
    const cellArr = toArr2(row?.["w:tc"] ?? row?.tc);
    return buildRow(
      cellArr.map((c) => buildCell([buildPara([buildSpan(cellText2(c))])]))
    );
  });
  return buildGrid(rowNodes);
}
function decodeGridFlat2(tbl) {
  return buildGrid([
    buildRow([buildCell([buildPara([buildSpan(tableText2(tbl))])])])
  ]);
}
function decodeGridText2(tbl) {
  return buildPara([buildSpan(tableText2(tbl))]);
}
function cellText2(cell) {
  return toArr2(cell?.["w:p"] ?? cell?.p).map(
    (p) => toArr2(p?.["w:r"] ?? p?.r).map(
      (r) => toArr2(r?.["w:t"] ?? r?.t).map((t) => typeof t === "string" ? t : t?._ ?? "").join("")
    ).join("")
  ).join(" ");
}
function tableText2(tbl) {
  return toArr2(tbl?.["w:tr"] ?? tbl?.tr).map(
    (row) => toArr2(row?.["w:tc"] ?? row?.tc).map((c) => cellText2(c)).join("	")
  ).join("\n");
}
function parseHeading(style) {
  if (!style) return void 0;
  const m = style.match(/[Hh]eading(\d)/);
  if (m) {
    const n = Number(m[1]);
    if (n >= 1 && n <= 6) return n;
  }
  return void 0;
}
registry.registerDecoder(new DocxDecoder());
function extractAnchorLayout(anchor) {
  const attr = anchor?._attr ?? {};
  const behindDoc = attr.behindDoc === "1";
  let wrap = "square";
  if (anchor?.["wp:wrapNone"]?.[0] != null)
    wrap = behindDoc ? "behind" : "none";
  else if (anchor?.["wp:wrapTight"]?.[0] != null) wrap = "tight";
  else if (anchor?.["wp:wrapThrough"]?.[0] != null) wrap = "through";
  else if (anchor?.["wp:wrapSquare"]?.[0] != null) wrap = "square";
  else if (anchor?.["wp:wrapTopAndBottom"]?.[0] != null) wrap = "square";
  else if (anchor?.["wp:wrapBehind"]?.[0] != null || behindDoc) wrap = "behind";
  const posH = anchor?.["wp:positionH"]?.[0];
  const horzRelTo = parseHorzRelTo(posH?._attr?.relativeFrom);
  const horzAlignTxt = posH?.["wp:align"]?.[0]?._text;
  const horzOffsetTxt = posH?.["wp:posOffset"]?.[0]?._text;
  const horzAlign = horzAlignTxt ? parseHorzAlign(horzAlignTxt) : void 0;
  const xPt = horzOffsetTxt && !horzAlignTxt ? Metric.emuToPt(Number(horzOffsetTxt)) : void 0;
  const posV = anchor?.["wp:positionV"]?.[0];
  const vertRelTo = parseVertRelTo(posV?._attr?.relativeFrom);
  const vertAlignTxt = posV?.["wp:align"]?.[0]?._text;
  const vertOffsetTxt = posV?.["wp:posOffset"]?.[0]?._text;
  const vertAlign = vertAlignTxt ? parseVertAlign(vertAlignTxt) : void 0;
  const yPt = vertOffsetTxt && !vertAlignTxt ? Metric.emuToPt(Number(vertOffsetTxt)) : void 0;
  const distT = attr.distT ? Metric.emuToPt(Number(attr.distT)) : void 0;
  const distB = attr.distB ? Metric.emuToPt(Number(attr.distB)) : void 0;
  const distL = attr.distL ? Metric.emuToPt(Number(attr.distL)) : void 0;
  const distR = attr.distR ? Metric.emuToPt(Number(attr.distR)) : void 0;
  const zOrder = attr.relativeHeight ? Number(attr.relativeHeight) : void 0;
  return {
    wrap,
    horzAlign,
    vertAlign,
    horzRelTo,
    vertRelTo,
    xPt,
    yPt,
    distT,
    distB,
    distL,
    distR,
    behindDoc,
    zOrder
  };
}
var HORZ_RELTO_MAP = {
  margin: "margin",
  leftMargin: "margin",
  rightMargin: "margin",
  insideMargin: "margin",
  outsideMargin: "margin",
  column: "column",
  page: "page",
  character: "para",
  paragraph: "para"
};
var VERT_RELTO_MAP = {
  margin: "margin",
  topMargin: "margin",
  bottomMargin: "margin",
  insideMargin: "margin",
  outsideMargin: "margin",
  line: "line",
  page: "page",
  paragraph: "para"
};
var HORZ_ALIGN_MAP = {
  left: "left",
  center: "center",
  right: "right",
  inside: "left",
  outside: "right"
};
var VERT_ALIGN_MAP = {
  top: "top",
  center: "center",
  bottom: "bottom",
  inside: "top",
  outside: "bottom"
};
function parseHorzRelTo(v) {
  return HORZ_RELTO_MAP[v ?? ""] ?? "column";
}
function parseVertRelTo(v) {
  return VERT_RELTO_MAP[v ?? ""] ?? "para";
}
function parseHorzAlign(v) {
  return HORZ_ALIGN_MAP[v ?? ""];
}
function parseVertAlign(v) {
  return VERT_ALIGN_MAP[v ?? ""];
}

// src/decoders/md/MdDecoder.ts
var MdDecoder = class extends BaseDecoder {
  getFormat() {
    return "md";
  }
  async decode(data) {
    const shield = new ShieldedParser();
    const warns = [];
    try {
      const text = this.bytesToString(data);
      const lines = text.split(/\r?\n/);
      const kids = [];
      let i = 0;
      while (i < lines.length) {
        const line = lines[i];
        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
          const level = headingMatch[1].length;
          kids.push(buildPara([buildSpan(headingMatch[2], { b: level <= 2 })], { heading: level }));
          i++;
          continue;
        }
        if (line.includes("|") && i + 1 < lines.length && lines[i + 1].match(/^\s*\|?\s*[-:]+\s*\|/)) {
          const tableResult = shield.guard(() => parseMdTable(lines, i), null, `md:table@${i}`);
          if (tableResult) {
            kids.push(tableResult.node);
            i = tableResult.nextLine;
            continue;
          }
        }
        if (line.match(/^[-*_]{3,}$/)) {
          kids.push(buildPara([buildSpan("")], {}));
          i++;
          continue;
        }
        const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.+)$/);
        if (listMatch) {
          kids.push(buildPara(parseInline(listMatch[3]), {
            listLv: Math.floor(listMatch[1].length / 2),
            listOrd: /\d+\./.test(listMatch[2])
          }));
          i++;
          continue;
        }
        const bqMatch = line.match(/^>\s*(.*)$/);
        if (bqMatch) {
          kids.push(buildPara([buildSpan(bqMatch[1])], { indentPt: 28 }));
          i++;
          continue;
        }
        if (line.startsWith("```")) {
          const codeLines = [];
          i++;
          while (i < lines.length && !lines[i].startsWith("```")) {
            codeLines.push(lines[i]);
            i++;
          }
          i++;
          kids.push(buildPara([buildSpan(codeLines.join("\n"), { font: "Courier New" })], {}));
          continue;
        }
        if (line.trim() === "") {
          i++;
          continue;
        }
        const alignMatch = line.match(/^<div\s+align="(center|right|left)">(.*?)<\/div>$/i);
        if (alignMatch) {
          const align = alignMatch[1].toLowerCase();
          kids.push(buildPara(parseInline(alignMatch[2]), { align }));
          i++;
          continue;
        }
        kids.push(buildPara(parseInline(line), {}));
        i++;
      }
      warns.push(...shield.flush());
      const sheet = buildSheet(kids.length > 0 ? kids : [buildPara([buildSpan("")])], A4);
      return succeed(buildRoot({}, [sheet]), warns);
    } catch (e) {
      warns.push(...shield.flush());
      return fail(`MD decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
};
function parseInline(text) {
  const result = [];
  let rem = text;
  while (rem.length > 0) {
    let m = rem.match(/^(.*?)!\[([^\]]*)\]\((data:([^;]+);base64,([^)]+))\)(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      const mime = m[4];
      const validMimes = ["image/png", "image/jpeg", "image/gif", "image/bmp"];
      result.push(buildImg(m[5], validMimes.includes(mime) ? mime : "image/png", 100, 100, m[2] || void 0));
      rem = m[6];
      continue;
    }
    m = rem.match(/^(.*?)!\[([^\]]*)\]\(([^)]+)\)(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(`[\uC774\uBBF8\uC9C0: ${m[2] || m[3]}]`));
      rem = m[4];
      continue;
    }
    m = rem.match(/^(.*?)\*\*\*(.+?)\*\*\*(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { b: true, i: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)\*\*(.+?)\*\*(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { b: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)\*(.+?)\*(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { i: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)~~(.+?)~~(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { s: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)<u>(.+?)<\/u>(.*)/si);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { u: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)<sup>(.+?)<\/sup>(.*)/si);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { sup: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)<sub>(.+?)<\/sub>(.*)/si);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { sub: true }));
      rem = m[3];
      continue;
    }
    m = rem.match(/^(.*?)`(.+?)`(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      result.push(buildSpan(m[2], { font: "Courier New" }));
      rem = m[3];
      continue;
    }
    result.push(buildSpan(rem));
    break;
  }
  return result.length > 0 ? result : [buildSpan(text)];
}
function parseMdTable(lines, startLine) {
  const parse = (line) => line.split("|").map((c) => c.trim()).filter((c, i, arr) => i > 0 || c !== "");
  const headers = parse(lines[startLine]);
  let cur = startLine + 2;
  const rows = [];
  while (cur < lines.length) {
    if (!lines[cur].includes("|")) break;
    const cells = parse(lines[cur]);
    if (cells.length === 0) break;
    rows.push(cells);
    cur++;
  }
  const allRows = [headers, ...rows];
  const gridRows = allRows.map(
    (row, ri) => buildRow(row.map((cell) => buildCell([buildPara([buildSpan(cell, ri === 0 ? { b: true } : {})])])))
  );
  return { node: buildGrid(gridRows), nextLine: cur };
}
registry.registerDecoder(new MdDecoder());

// src/decoders/html/HtmlDecoder.ts
var HtmlDecoder = class extends BaseDecoder {
  getFormat() {
    return "html";
  }
  async decode(data) {
    const shield = new ShieldedParser();
    const warns = [];
    try {
      const html = this.bytesToString(data);
      const tokens = shield.guard(() => tokenize(html), [], "html:tokenize");
      const kids = shield.guard(() => parseTokens(tokens), [], "html:parse");
      warns.push(...shield.flush());
      const sheet = buildSheet(kids.length > 0 ? kids : [buildPara([buildSpan("")])], A4);
      return succeed(buildRoot({}, [sheet]), warns);
    } catch (e) {
      warns.push(...shield.flush());
      return fail(`HTML decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
};
function tokenize(html) {
  const tokens = [];
  let i = 0;
  while (i < html.length) {
    if (html[i] === "<") {
      if (html[i + 1] === "!") {
        const end2 = html.indexOf(">", i);
        i = end2 + 1;
        continue;
      }
      const isClose = html[i + 1] === "/";
      const start = isClose ? i + 2 : i + 1;
      const end = html.indexOf(">", i);
      if (end === -1) break;
      const tagContent = html.slice(start, end).trim();
      const spaceIdx = tagContent.search(/\s/);
      const name = spaceIdx > 0 ? tagContent.slice(0, spaceIdx) : tagContent;
      const attrsStr = spaceIdx > 0 ? tagContent.slice(spaceIdx + 1).trim() : "";
      const attrs = {};
      if (attrsStr) {
        const attrRegex = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
        let m;
        while ((m = attrRegex.exec(attrsStr)) !== null) {
          attrs[m[1].toLowerCase()] = m[2] ?? m[3] ?? m[4] ?? "";
        }
      }
      tokens.push({
        type: "tag",
        name: name.toLowerCase(),
        attrs,
        selfClose: html[end - 1] === "/",
        close: isClose
      });
      i = end + 1;
    } else {
      const end = html.indexOf("<", i);
      const text = end === -1 ? html.slice(i) : html.slice(i, end);
      if (text.trim()) {
        tokens.push({ type: "text", content: text });
      }
      i = end === -1 ? html.length : end;
    }
  }
  return tokens;
}
function parseTokens(tokens) {
  const kids = [];
  let i = 0;
  while (i < tokens.length) {
    const t = tokens[i];
    if (t.type === "tag" && !t.close) {
      switch (t.name) {
        case "html":
          i++;
          let bodyStart = -1;
          let depth = 1;
          while (i < tokens.length && depth > 0) {
            if (tokens[i].type === "tag" && !tokens[i].close && tokens[i].name === "html") depth++;
            else if (tokens[i].type === "tag" && tokens[i].close && tokens[i].name === "html") depth--;
            else if (tokens[i].type === "tag" && !tokens[i].close && tokens[i].name === "body") {
              bodyStart = i + 1;
            }
            i++;
          }
          if (bodyStart > 0) {
            let bodyEnd = bodyStart;
            let bodyDepth = 1;
            while (bodyEnd < tokens.length && bodyDepth > 0) {
              if (tokens[bodyEnd].type === "tag" && !tokens[bodyEnd].close && tokens[bodyEnd].name === "body") bodyDepth++;
              else if (tokens[bodyEnd].type === "tag" && tokens[bodyEnd].close && tokens[bodyEnd].name === "body") bodyDepth--;
              bodyEnd++;
            }
            bodyEnd--;
            const bodyTokens = tokens.slice(bodyStart, bodyEnd);
            const bodyKids = parseTokens(bodyTokens);
            kids.push(...bodyKids);
          }
          continue;
        case "head":
        case "style":
        case "script":
          i = skipBlock(tokens, i, t.name);
          continue;
        case "body":
        case "div":
        case "section":
        case "article":
        case "main":
          const start = i + 1;
          let end = start;
          let divDepth = 1;
          while (end < tokens.length && divDepth > 0) {
            const t2 = tokens[end];
            if (t2.type === "tag" && !t2.close) {
              if (["html", "head", "body", "div", "section", "article", "main"].includes(t2.name ?? "")) divDepth++;
            } else if (t2.type === "tag" && t2.close) {
              if (["html", "head", "body", "div", "section", "article", "main"].includes(t2.name ?? "")) divDepth--;
            }
            end++;
          }
          end--;
          const subTokens = tokens.slice(start, end);
          const subKids = parseTokens(subTokens);
          kids.push(...subKids);
          i = end + 1;
          continue;
        case "p":
          i++;
          const paraKids = collectInline(tokens, i, ["p", "div", "br"]);
          i = paraKids.nextI;
          const align = t.attrs?.style?.includes("text-align: center") ? "center" : t.attrs?.style?.includes("text-align: right") ? "right" : t.attrs?.style?.includes("text-align: left") ? "left" : void 0;
          kids.push(buildPara(paraKids.nodes, { align }));
          continue;
        case "br":
          kids.push(buildPara([buildSpan("")], {}));
          i++;
          continue;
        case "img":
          i++;
          const src = t.attrs?.src;
          const alt = t.attrs?.alt || "";
          if (src?.startsWith("data:")) {
            const match = src.match(/^data:([^;]+);base64,(.+)$/);
            if (match) {
              kids.push(buildPara([buildImg(match[2], match[1], 100, 100, alt)], {}));
            }
          }
          continue;
        case "table":
          i++;
          const rows = [];
          while (i < tokens.length) {
            if (tokens[i].type === "tag" && tokens[i].close && tokens[i].name === "table") {
              i++;
              break;
            }
            if (tokens[i].type === "tag" && tokens[i].name === "tr" && !tokens[i].close) {
              i++;
              const cells = [];
              while (i < tokens.length) {
                if (tokens[i].type === "tag" && tokens[i].close && tokens[i].name === "tr") {
                  i++;
                  break;
                }
                if (tokens[i].type === "tag" && (tokens[i].name === "td" || tokens[i].name === "th") && !tokens[i].close) {
                  i++;
                  const cellKids = collectInline(tokens, i, ["td", "th", "tr"]);
                  i = cellKids.nextI;
                  const isHeader = tokens[i - 2]?.name === "th";
                  const paraKids2 = cellKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, b: isHeader } } : n);
                  cells.push(buildCell([buildPara(paraKids2, {})]));
                } else if (tokens[i].type === "text" && tokens[i].content?.trim()) {
                  cells.push(buildCell([buildPara([buildSpan(tokens[i].content.trim())])]));
                  i++;
                } else {
                  i++;
                }
              }
              if (cells.length > 0) rows.push(buildRow(cells));
            } else {
              i++;
            }
          }
          if (rows.length > 0) kids.push(buildGrid(rows));
          continue;
        case "ul":
        case "ol":
          i++;
          const isOrdered = t.name === "ol";
          while (i < tokens.length) {
            if (tokens[i].type === "tag" && tokens[i].close && tokens[i].name === t.name) {
              i++;
              break;
            }
            if (tokens[i].type === "tag" && tokens[i].name === "li" && !tokens[i].close) {
              i++;
              const liKids = collectInline(tokens, i, ["li", "ul", "ol"]);
              i = liKids.nextI;
              kids.push(buildPara(liKids.nodes, { listOrd: isOrdered }));
            } else {
              i++;
            }
          }
          continue;
        default:
          i++;
      }
    } else if (t.type === "text" && t.content?.trim()) {
      kids.push(buildPara([buildSpan(t.content.trim())], {}));
      i++;
    } else {
      i++;
    }
  }
  return kids;
}
function collectInline(tokens, start, stopTags) {
  const nodes = [];
  let i = start;
  while (i < tokens.length) {
    const t = tokens[i];
    if (t.type === "tag" && !t.close) {
      if (t.name && stopTags.includes(t.name)) {
        break;
      }
      switch (t.name) {
        case "b":
        case "strong":
          i++;
          const boldKids = collectInline(tokens, i, ["b", "strong", ...stopTags]);
          i = boldKids.nextI;
          nodes.push(...boldKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, b: true } } : n));
          continue;
        case "i":
        case "em":
          i++;
          const italicKids = collectInline(tokens, i, ["i", "em", ...stopTags]);
          i = italicKids.nextI;
          nodes.push(...italicKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, i: true } } : n));
          continue;
        case "u":
          i++;
          const underlineKids = collectInline(tokens, i, ["u", ...stopTags]);
          i = underlineKids.nextI;
          nodes.push(...underlineKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, u: true } } : n));
          continue;
        case "s":
        case "strike":
          i++;
          const strikeKids = collectInline(tokens, i, ["s", "strike", ...stopTags]);
          i = strikeKids.nextI;
          nodes.push(...strikeKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, s: true } } : n));
          continue;
        case "span":
          i++;
          const spanKids = collectInline(tokens, i, ["span", ...stopTags]);
          i = spanKids.nextI;
          const color = t.attrs?.style?.match(/color:\s*([^;]+)/)?.[1];
          nodes.push(...spanKids.nodes.map((n) => n.tag === "span" ? { ...n, props: { ...n.props, color: color || n.props.color } } : n));
          continue;
        case "img":
          const src = t.attrs?.src;
          const alt = t.attrs?.alt || "";
          if (src?.startsWith("data:")) {
            const match = src.match(/^data:([^;]+);base64,(.+)$/);
            if (match) {
              nodes.push(buildImg(match[2], match[1], 100, 100, alt));
            }
          }
          i++;
          continue;
        default:
          i++;
      }
    } else if (t.type === "text") {
      if (t.content?.trim()) {
        nodes.push(buildSpan(t.content.trim()));
      }
      i++;
    } else {
      i++;
    }
  }
  return { nodes: nodes.length > 0 ? nodes : [buildSpan("")], nextI: i };
}
function skipBlock(tokens, start, name) {
  let i = start + 1;
  let depth = 1;
  while (i < tokens.length && depth > 0) {
    if (tokens[i].type === "tag") {
      if (!tokens[i].close && tokens[i].name === name) depth++;
      if (tokens[i].close && tokens[i].name === name) depth--;
    }
    i++;
  }
  return i;
}
registry.registerDecoder(new HtmlDecoder());

// src/core/BaseEncoder.ts
var BaseEncoder = class {
  constructor() {
    this.format = this.getFormat();
    this.aliases = this.getAliases();
  }
  /** 별칭 목록 반환 (하위 클래스에서 필요 시 오버라이드) */
  getAliases() {
    return [];
  }
  // ─── 공통 유틸리티 메서드 ───────────────────────────
  /** 문서 내 모든 이미지 노드 수집 */
  collectImages(doc) {
    const images = [];
    this.collectImagesRecursive(doc, images);
    return images;
  }
  /** 재귀적으로 이미지 수집 */
  collectImagesRecursive(node, images) {
    if (node.tag === "img") {
      images.push(node);
      return;
    }
    const children = this.getChildren(node);
    for (const child of children) {
      this.collectImagesRecursive(child, images);
    }
  }
  /** 노드의 자식 노드 반환 */
  getChildren(node) {
    switch (node.tag) {
      case "root":
        return node.kids;
      case "sheet":
        return [...node.kids, ...node.headers?.default ?? [], ...node.footers?.default ?? []];
      case "para":
        return node.kids;
      case "span":
        return node.kids;
      case "link":
        return node.kids;
      case "row":
        return node.kids;
      case "cell":
        return node.kids;
      case "grid":
        return node.kids;
      default:
        return [];
    }
  }
  /** base64 문자열을 Uint8Array 로 변환 */
  base64ToBytes(b64) {
    return TextKit.base64Decode(b64);
  }
  /** Uint8Array 를 base64 문자열로 변환 */
  bytesToBase64(data) {
    return TextKit.base64Encode(data);
  }
  /** XML 이스케이프 */
  escapeXml(s) {
    return TextKit.escapeXml(s);
  }
  /** XML 언이스케이프 */
  unescapeXml(s) {
    return TextKit.unescapeXml(s);
  }
  /** 문자열을 UTF-8 바이트로 변환 */
  stringToBytes(s) {
    return TextKit.encode(s);
  }
  /** 바이트를 UTF-8 문자열로 변환 */
  bytesToString(data) {
    return TextKit.decode(data);
  }
  /** ZIP 압축 */
  async zip(entries) {
    return ArchiveKit.zip(entries);
  }
  /** ZIP 해제 */
  async unzip(data) {
    return ArchiveKit.unzip(data);
  }
  /** deflate 압축 */
  async deflate(data) {
    return ArchiveKit.deflate(data);
  }
  /** inflate 해제 */
  async inflate(data) {
    return ArchiveKit.inflate(data);
  }
  /** 제어 문자 제거 */
  stripControl(s) {
    return TextKit.stripControl(s);
  }
  /** 공백 정규화 */
  normalizeWhitespace(s) {
    return TextKit.normalizeWhitespace(s);
  }
};

// src/encoders/hwpx/HwpxEncoder.ts
var NS = [
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"',
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"',
  'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"',
  'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"',
  'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"',
  'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"',
  'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"',
  'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"',
  'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:opf="http://www.idpf.org/2007/opf"',
  'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"',
  'xmlns:epub="http://www.idpf.org/2007/ops"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
].join(" ");
var LINESEG_FLAGS_FIRST = 1441792;
var LINESEG_FLAGS_OTHER = 393216;
var LANG_GROUPS = [
  "HANGUL",
  "LATIN",
  "HANJA",
  "JAPANESE",
  "OTHER",
  "SYMBOL",
  "USER"
];
var LangFontBank = class {
  constructor() {
    // 언어 그룹별 독립 폰트 맵: face → localId (0-based)
    this.maps = new Map(
      LANG_GROUPS.map((g) => [g, /* @__PURE__ */ new Map()])
    );
    this.registerAll("\uD568\uCD08\uB86C\uBC14\uD0D5");
  }
  /** 모든 언어 그룹에 동일 폰트 등록 */
  registerAll(face) {
    for (const g of LANG_GROUPS) {
      const m = this.maps.get(g);
      if (!m.has(face)) m.set(face, m.size);
    }
  }
  /** 특정 언어 그룹에 폰트 등록, 이미 있으면 기존 ID 반환 */
  register(lang, face) {
    const m = this.maps.get(lang);
    if (m.has(face)) return m.get(face);
    const id = m.size;
    m.set(face, id);
    return id;
  }
  /** 폰트 이름 → 한글 폰트 여부 판별 (ANYTOHWP 방식) */
  isKorean(face) {
    return /[\uAC00-\uD7A3\u3131-\u318E]/.test(face) || ["\uB9D1\uC740", "\uB098\uB214", "\uAD74\uB9BC", "\uB3CB\uC6C0", "\uBC14\uD0D5", "\uD568\uCD08\uB86C", "\uD55C\uCEF4", "HY"].some(
      (k) => face.includes(k)
    );
  }
  /** TextProps.font 문자열에서 적절한 HANGUL/LATIN 그룹에 등록 */
  registerFont(rawFace) {
    const face = safeFontToKr(rawFace) || "\uD568\uCD08\uB86C\uBC14\uD0D5";
    const isKor = this.isKorean(face);
    const hangulId = this.register("HANGUL", isKor ? face : "\uD568\uCD08\uB86C\uBC14\uD0D5");
    const latinId = this.register("LATIN", isKor ? "\uD568\uCD08\uB86C\uBC14\uD0D5" : face);
    for (const g of [
      "HANJA",
      "JAPANESE",
      "OTHER",
      "SYMBOL",
      "USER"
    ]) {
      this.register(g, isKor ? face : "\uD568\uCD08\uB86C\uBC14\uD0D5");
    }
    return { hangulId, latinId };
  }
  /** 언어 그룹별 폰트 목록 반환 */
  getFaces(lang) {
    return [...this.maps.get(lang).keys()];
  }
  getId(lang, face) {
    return this.maps.get(lang).get(face) ?? 0;
  }
  /** hh:fontfaces XML 생성 */
  toXml() {
    let xml = `<hh:fontfaces itemCnt="${LANG_GROUPS.length}">`;
    for (const lang of LANG_GROUPS) {
      const faces = this.getFaces(lang);
      xml += `<hh:fontface lang="${lang}" fontCnt="${faces.length}">`;
      faces.forEach((face, i) => {
        xml += `<hh:font id="${i}" face="${esc(face)}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_UNKNOWN" weight="0" proportion="0" contrast="0" strokeVariation="0" armStyle="0" letterform="0" midline="252" xHeight="255"/></hh:font>`;
      });
      xml += `</hh:fontface>`;
    }
    return xml + `</hh:fontfaces>`;
  }
};
var KIND_MAP = {
  solid: "SOLID",
  dash: "DASH",
  dot: "DOT",
  double: "DOUBLE",
  none: "NONE",
  dash_dot: "DASH_DOT",
  dash_dot_dot: "DASH_DOT_DOT"
};
var BorderFillBank = class {
  constructor() {
    this.fills = [];
    this.keyMap = /* @__PURE__ */ new Map();
    this._addXml(
      this._buildXml(void 0, void 0, void 0, void 0, void 0)
    );
    const defS = { kind: "solid", pt: 0.5, color: "000000" };
    this._addXml(this._buildXml(defS, defS, defS, defS, void 0));
  }
  _strokeXml(tag, s) {
    const type = s && s.kind !== "none" ? KIND_MAP[s.kind] ?? "SOLID" : "NONE";
    const w = s && s.kind !== "none" ? `${(s.pt * 0.3528).toFixed(2)} mm` : "0.12 mm";
    const c = s ? s.color.startsWith("#") ? s.color : `#${s.color}` : "#000000";
    return `<hh:${tag} type="${type}" width="${w}" color="${c}"/>`;
  }
  _buildXml(top, right, bottom, left, bg) {
    const fill = bg ? `<hc:fillBrush><hc:winBrush faceColor="${bg.startsWith("#") ? bg : "#" + bg}" hatchColor="none" alpha="0"/></hc:fillBrush>` : "";
    return `<hh:borderFill id="__ID__" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>` + this._strokeXml("leftBorder", left) + this._strokeXml("rightBorder", right) + this._strokeXml("topBorder", top) + this._strokeXml("bottomBorder", bottom) + `<hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>` + fill + `</hh:borderFill>`;
  }
  _addXml(xml) {
    const id = this.fills.length + 1;
    this.fills.push({ id, xml: xml.replace("__ID__", String(id)) });
    return id;
  }
  _key(top, right, bottom, left, bg) {
    const sk = (s) => s ? `${s.kind}:${s.pt.toFixed(2)}:${s.color}` : "none";
    return `${sk(top)}|${sk(right)}|${sk(bottom)}|${sk(left)}|${bg ?? ""}`;
  }
  /** 균일 테두리 등록 */
  addUniform(s, bg) {
    const key = this._key(s, s, s, s, bg);
    if (this.keyMap.has(key)) return this.keyMap.get(key);
    const id = this._addXml(this._buildXml(s, s, s, s, bg));
    this.keyMap.set(key, id);
    return id;
  }
  /** 방향별 테두리 등록 */
  addPerSide(top, right, bottom, left, bg) {
    const key = this._key(top, right, bottom, left, bg);
    if (this.keyMap.has(key)) return this.keyMap.get(key);
    const id = this._addXml(this._buildXml(top, right, bottom, left, bg));
    this.keyMap.set(key, id);
    return id;
  }
  /** CellProps에서 적절한 borderFill ID 계산 (하드코딩 "1" 완전 제거) */
  addFromCellProps(cp, defStroke) {
    const d = defStroke ?? DEFAULT_STROKE;
    const top = cp.top ?? d;
    const right = cp.right ?? d;
    const bottom = cp.bot ?? d;
    const left = cp.left ?? d;
    const bg = cp.bg;
    const uniform = top.kind === right.kind && top.kind === bottom.kind && top.kind === left.kind && top.pt === right.pt && top.pt === bottom.pt && top.pt === left.pt && top.color === right.color && top.color === bottom.color && top.color === left.color;
    return uniform ? this.addUniform(top, bg) : this.addPerSide(top, right, bottom, left, bg);
  }
  toXml() {
    return `<hh:borderFills itemCnt="${this.fills.length}">${this.fills.map((f) => f.xml).join("")}</hh:borderFills>`;
  }
};
function readPixelDims(b64, mime) {
  try {
    const raw = TextKit.base64Decode(b64);
    const view = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    if (mime.includes("png")) {
      if (raw.length >= 24 && view.getUint32(0) === 2303741511 && view.getUint32(4) === 218765834) {
        return { w: view.getUint32(16), h: view.getUint32(20) };
      }
    } else if (mime.includes("jpeg") || mime.includes("jpg")) {
      let off = 2;
      while (off < raw.length - 4) {
        const marker = view.getUint16(off);
        off += 2;
        if (marker === 65472 || marker === 65474) {
          return { w: view.getUint16(off + 5), h: view.getUint16(off + 3) };
        }
        if ((marker & 65280) !== 65280) break;
        const segLen = view.getUint16(off);
        off += segLen;
      }
    }
  } catch {
  }
  return null;
}
function charPrKey(p) {
  return `${p.b ? 1 : 0}|${p.i ? 1 : 0}|${p.u ? 1 : 0}|${p.s ? 1 : 0}|${p.pt ?? 10}|${p.color ?? "000000"}|${p.font ?? ""}|${p.bg ?? ""}`;
}
function paraPrKey(p) {
  return `${p.align ?? "left"}|${p.listOrd ?? ""}|${p.listLv ?? 0}|${p.indentPt ?? 0}|${p.firstLineIndentPt ?? 0}|${p.spaceBefore ?? 0}|${p.spaceAfter ?? 0}|${p.lineHeight ?? 0}|${p.styleId ?? ""}`;
}
function registerCharPr(props, ctx) {
  const key = charPrKey(props);
  const existing = ctx.charPrMap.get(key);
  if (existing !== void 0) return existing;
  const rawFont = props.font ?? "\uD568\uCD08\uB86C\uBC14\uD0D5";
  const { hangulId, latinId } = ctx.fontBank.registerFont(rawFont);
  const id = ctx.charPrs.length;
  ctx.charPrs.push({
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? "BOTTOM" : "NONE",
    strikeout: props.s ? "SOLID" : "NONE",
    textColor: props.color ? `#${props.color}` : "#000000",
    hangulId,
    latinId,
    bg: props.bg
  });
  ctx.charPrMap.set(key, id);
  return id;
}
function registerParaPr(props, ctx) {
  const key = paraPrKey(props);
  const existing = ctx.paraPrMap.get(key);
  if (existing !== void 0) return existing;
  const id = ctx.paraPrs.length;
  const def = {
    id,
    align: (props.align ?? "left").toUpperCase(),
    leftHwp: Metric.ptToHwp(props.indentPt ?? 0),
    rightHwp: Metric.ptToHwp(props.indentRightPt ?? 0),
    intentHwp: Metric.ptToHwp(props.firstLineIndentPt ?? 0),
    prevHwp: Metric.ptToHwp(props.spaceBefore ?? 0),
    nextHwp: Metric.ptToHwp(props.spaceAfter ?? 0),
    lineSpacing: props.lineHeight ? Math.round(props.lineHeight * 100) : 160
  };
  if (props.listOrd !== void 0) {
    def.listType = props.listOrd ? "DIGIT" : "BULLET";
    def.listLevel = props.listLv ?? 0;
  }
  ctx.paraPrs.push(def);
  ctx.paraPrMap.set(key, id);
  return id;
}
function mimeToExt(mime) {
  if (mime.includes("jpeg")) return "jpg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  return "png";
}
function registerImage(img, ctx) {
  if (ctx.imgMap.has(img)) return;
  const ext = mimeToExt(img.mime);
  const id = `BIN${String(ctx.nextBinNum).padStart(4, "0")}`;
  const name = `${id}.${ext}`;
  ctx.nextBinNum++;
  const data = TextKit.base64Decode(img.b64);
  ctx.bins.push({ id, name, data });
  ctx.imgMap.set(img, id);
}
var STYLE_NAME_MAP = {
  Normal: "\uBC14\uD0D5\uAE00",
  "Heading 1": "\uAC1C\uC694 1",
  "Heading 2": "\uAC1C\uC694 2",
  "Heading 3": "\uAC1C\uC694 3",
  "Heading 4": "\uAC1C\uC694 4",
  "Heading 5": "\uAC1C\uC694 5",
  "Heading 6": "\uAC1C\uC694 6",
  "Body Text": "\uBCF8\uBB38"
};
function registerStyle(styleId, paraPrId, charPrId, ctx) {
  if (!styleId || ctx.styleIdToHwpxId.has(styleId)) return;
  if (styleId === "Normal") {
    ctx.styleIdToHwpxId.set(styleId, 0);
    return;
  }
  const hwpxId = ctx.hwpxStyles.length;
  ctx.styleIdToHwpxId.set(styleId, hwpxId);
  ctx.hwpxStyles.push({
    id: hwpxId,
    name: STYLE_NAME_MAP[styleId] ?? styleId,
    engName: "",
    paraPrIDRef: paraPrId,
    charPrIDRef: charPrId
  });
}
function scanPara(para, ctx) {
  const paraPrId = registerParaPr(para.props, ctx);
  let firstCharPrId = 0;
  let hasFirstSpan = false;
  function scanKids(kids) {
    for (const kid of kids) {
      if (kid.tag === "span") {
        const cId = registerCharPr(kid.props, ctx);
        if (!hasFirstSpan) {
          firstCharPrId = cId;
          hasFirstSpan = true;
        }
      } else if (kid.tag === "img") {
        registerImage(kid, ctx);
      } else if (kid.tag === "link") {
        scanKids(kid.kids);
      }
    }
  }
  scanKids(para.kids);
  if (para.props.styleId)
    registerStyle(para.props.styleId, paraPrId, firstCharPrId, ctx);
}
function scanGrid(grid, ctx) {
  const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
  ctx.borderFillBank.addUniform(defStroke);
  for (const row of grid.kids) {
    for (const cell of row.kids) {
      ctx.borderFillBank.addFromCellProps(cell.props, defStroke);
      for (const p of cell.kids) {
        if (p.tag === "grid") scanGrid(p, ctx);
        else scanPara(p, ctx);
      }
    }
  }
}
function scanContent(kids, ctx) {
  for (const kid of kids) {
    if (kid.tag === "para") scanPara(kid, ctx);
    else if (kid.tag === "grid") scanGrid(kid, ctx);
  }
}
var HwpxEncoder = class extends BaseEncoder {
  getFormat() {
    return "hwpx";
  }
  getAliases() {
    return [HWPX_MIME_TYPE, "application/hwp+zip"];
  }
  async encode(doc) {
    try {
      const sheet = doc.kids[0];
      const dims = normalizeDims(sheet?.dims ?? A4);
      const safeML = dims.ml > 0 ? dims.ml : 70.87;
      const safeMR = dims.mr > 0 ? dims.mr : 70.87;
      const availableWidth = Math.round(
        Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(safeML) - Metric.ptToHwp(safeMR)
      );
      const ctx = {
        fontBank: new LangFontBank(),
        // ANYTOHWP 방식 언어별 폰트
        borderFillBank: new BorderFillBank(),
        // 하드코딩 없는 테두리 관리
        charPrs: [],
        charPrMap: /* @__PURE__ */ new Map(),
        paraPrs: [],
        paraPrMap: /* @__PURE__ */ new Map(),
        bins: [],
        nextBinNum: 1,
        nextElementId: 1e4,
        availableWidth,
        imgMap: /* @__PURE__ */ new WeakMap(),
        nextZOrder: 0,
        styleIdToHwpxId: /* @__PURE__ */ new Map(),
        hwpxStyles: []
      };
      registerCharPr({}, ctx);
      registerParaPr({}, ctx);
      ctx.hwpxStyles.push({
        id: 0,
        name: "\uBC14\uD0D5\uAE00",
        engName: "Normal",
        paraPrIDRef: 0,
        charPrIDRef: 0
      });
      ctx.styleIdToHwpxId.set("Normal", 0);
      scanContent(sheet?.kids ?? [], ctx);
      if (sheet?.headers?.default) for (const p of sheet.headers.default) scanPara(p, ctx);
      if (sheet?.footers?.default) for (const p of sheet.footers.default) scanPara(p, ctx);
      const sectionData = this.stringToBytes(buildSectionXml(sheet, dims, ctx));
      const headerData = this.stringToBytes(buildHeaderXml(dims, doc.meta, ctx));
      const previewText = extractPreviewText(sheet);
      const entries = [
        {
          name: "mimetype",
          data: new TextEncoder().encode(HWPX_MIME_TYPE),
          compression: "STORE",
          mime: ""
        },
        {
          name: "version.xml",
          data: this.stringToBytes(VERSION_XML),
          mime: "application/xml"
        },
        {
          name: "META-INF/container.xml",
          data: this.stringToBytes(CONTAINER_XML),
          mime: "application/xml"
        },
        {
          name: "META-INF/container.rdf",
          data: this.stringToBytes(CONTAINER_RDF),
          mime: "application/rdf+xml"
        },
        {
          name: "Contents/content.hpf",
          data: this.stringToBytes(buildContentHpf(ctx, doc.meta)),
          mime: "application/hwpml-package+xml"
        },
        {
          name: "Contents/header.xml",
          data: headerData,
          mime: "application/xml"
        },
        {
          name: "Contents/section0.xml",
          data: sectionData,
          mime: "application/xml"
        },
        {
          name: "Preview/PrvText.txt",
          data: this.stringToBytes(previewText),
          mime: "text/plain"
        },
        {
          name: "settings.xml",
          data: this.stringToBytes(buildSettingsXml()),
          mime: "application/xml"
        }
      ];
      for (const bin of ctx.bins) {
        const ext = bin.name.split(".").pop()?.toLowerCase() ?? "png";
        const ct = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/bmp";
        entries.push({ name: `BinData/${bin.name}`, data: bin.data, mime: ct });
      }
      return succeed(await this.zip(entries));
    } catch (e) {
      return fail(`HWPX \uC778\uCF54\uB529 \uC624\uB958: ${e?.message ?? String(e)}`);
    }
  }
};
var VERSION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.owpml.org/owpml/2024/version" targetApplication="WORDPROCESSING" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 0"/>`;
var CONTAINER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>`;
var CONTAINER_RDF = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>`;
function buildContentHpf(ctx, meta) {
  const title = esc(meta?.title ?? "");
  const creator = esc(meta?.author ?? "text");
  const subject = esc(meta?.subject ?? "text");
  const desc = esc(meta?.desc ?? "text");
  const keyword = esc(meta?.keywords ?? "text");
  const created = meta?.created ?? (/* @__PURE__ */ new Date()).toISOString().replace(/\.\d{3}Z$/, "Z");
  const modified = meta?.modified ?? created;
  let items = `<opf:item id="header"   href="Contents/header.xml"   media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="settings" href="settings.xml"          media-type="application/xml"/>`;
  for (const bin of ctx.bins) {
    const ext = bin.name.split(".").pop()?.toLowerCase() ?? "png";
    const ct = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/bmp";
    items += `<opf:item id="${bin.id}" href="BinData/${bin.name}" media-type="${ct}" isEmbeded="1"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title>${title}</opf:title><opf:language>ko</opf:language><opf:meta name="creator"      content="text">${creator}</opf:meta><opf:meta name="subject"      content="text">${subject}</opf:meta><opf:meta name="description"  content="text">${desc}</opf:meta><opf:meta name="CreatedDate"  content="text">${created}</opf:meta><opf:meta name="ModifiedDate" content="text">${modified}</opf:meta><opf:meta name="keyword"      content="text">${keyword}</opf:meta><opf:meta name="trackchageConfig" content="text">0</opf:meta></opf:metadata><opf:manifest>${items}</opf:manifest><opf:spine><opf:itemref idref="header"/><opf:itemref idref="section0"/></opf:spine></opf:package>`;
}
function buildSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/><config:config-item-set name="PrintInfo"><config:config-item name="PrintAutoFootNote" type="boolean">false</config:config-item><config:config-item name="PrintAutoHeadNote" type="boolean">false</config:config-item><config:config-item name="PrintMethod" type="short">4</config:config-item><config:config-item name="OverlapSize" type="short">0</config:config-item><config:config-item name="PrintCropMark" type="short">0</config:config-item><config:config-item name="BinderHoleType" type="short">0</config:config-item><config:config-item name="ZoomX" type="short">100</config:config-item><config:config-item name="ZoomY" type="short">100</config:config-item></config:config-item-set></ha:HWPApplicationSetting>`;
}
function buildNumberingsXml() {
  return `<hh:numberings itemCnt="1"><hh:numbering id="1" start="1"><hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1" autoIndent="0" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="BULLET" charPrIDRef="0" checkable="0"/></hh:numbering></hh:numberings>`;
}
function buildBulletsXml() {
  return `<hh:bullets itemCnt="1"><hh:bullet id="1" charPrIDRef="0" start="1" numFormat="BULLET"><hh:paraHead level="1" numChar="&#x2022;"/><hh:paraHead level="2" numChar="&#x2022;"/><hh:paraHead level="3" numChar="&#x2022;"/></hh:bullet></hh:bullets>`;
}
function buildHeaderXml(dims, meta, ctx) {
  const fontFacesXml = ctx.fontBank.toXml();
  let charPrXml = "";
  for (const cp of ctx.charPrs) {
    const bold = cp.bold ? "<hh:bold/>" : "";
    const italic = cp.italic ? "<hh:italic/>" : "";
    const hid = cp.hangulId;
    const lid = cp.latinId;
    const shadeColor = cp.bg ? cp.bg.startsWith("#") ? cp.bg : `#${cp.bg}` : "none";
    charPrXml += `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" shadeColor="${shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1"><hh:fontRef hangul="${hid}" latin="${lid}" hanja="${hid}" japanese="${hid}" other="${lid}" symbol="${lid}" user="${lid}"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>` + bold + italic + `<hh:underline type="${cp.underline}" shape="SOLID" color="#000000"/><hh:strikeout shape="${cp.strikeout}" color="#000000"/><hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>`;
  }
  let paraPrXml = "";
  for (const pp of ctx.paraPrs) {
    paraPrXml += `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0"><hh:align horizontal="${pp.align}" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/><hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/><hh:margin><hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/><hc:left value="${pp.leftHwp}" unit="HWPUNIT"/><hc:right value="${pp.rightHwp}" unit="HWPUNIT"/><hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/><hc:next value="${pp.nextHwp}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="${pp.lineSpacing}" unit="HWPUNIT"/><hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`;
  }
  const borderFillXml = ctx.borderFillBank.toXml();
  const stylesXml2 = `<hh:styles itemCnt="${ctx.hwpxStyles.length}">` + ctx.hwpxStyles.map(
    (s) => `<hh:style id="${s.id}" type="PARA" name="${esc(s.name)}" engName="${esc(s.engName)}" paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="0" langID="1042" lockForm="0"/>`
  ).join("") + `</hh:styles>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head ${NS} version="1.2" secCnt="1"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList>` + fontFacesXml + borderFillXml + `<hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>` + buildNumberingsXml() + buildBulletsXml() + `<hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties>` + stylesXml2 + `</hh:refList><hh:compatibleDocument targetProgram="MS_WORD">        
          <hh:layoutCompatibility>
            <hh:applyFontWeightToBold />
            <hh:useInnerUnderline />
            <hh:useLowercaseStrikeout />
            <hh:extendLineheightToOffset />
            <hh:treatQuotationAsLatin />
            <hh:doNotAlignWhitespaceOnRight />
            <hh:doNotAdjustWordInJustify />
            <hh:baseCharUnitOnEAsian />
            <hh:baseCharUnitOfIndentOnFirstChar />
            <hh:adjustLineheightToFont />
            <hh:adjustBaselineInFixedLinespacing />
            <hh:applyPrevspacingBeneathObject />
            <hh:applyNextspacingOfLastPara />
            <hh:adjustParaBorderfillToSpacing />
            <hh:connectParaBorderfillOfEqualBorder />
            <hh:adjustParaBorderOffsetWithBorder />
            <hh:extendLineheightToParaBorderOffset />
            <hh:applyParaBorderToOutside />
            <hh:applyMinColumnWidthTo1mm />
            <hh:applyTabPosBasedOnSegment />
            <hh:breakTabOverLine />
            <hh:adjustVertPosOfLine />
            <hh:doNotAlignLastForbidden />
            <hh:adjustMarginFromAdjustLineheight />
            <hh:baseLineSpacingOnLineGrid />
            <hh:applyCharSpacingToCharGrid />
            <hh:doNotApplyGridInHeaderFooter />
            <hh:applyExtendHeaderFooterEachSection />
            <hh:doNotApplyLinegridAtNoLinespacing />
            <hh:doNotAdjustEmptyAnchorLine />
            <hh:overlapBothAllowOverlap />
            <hh:extendVertLimitToPageMargins />
            <hh:doNotHoldAnchorOfTable />
            <hh:doNotFormattingAtBeneathAnchor />
            <hh:adjustBaselineOfObjectToBottom />
        </hh:layoutCompatibility>
    </hh:compatibleDocument><hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption><hh:trackchageConfig flags="56"/></hh:head>`;
}
function buildHeaderFooterRunXml(sheet, dims, ctx) {
  const headers = sheet.headers || {};
  const footers = sheet.footers || {};
  const hasAny = Object.keys(headers).length > 0 || Object.keys(footers).length > 0;
  if (!hasAny) return "";
  const availW = ctx.availableWidth;
  const mtHwp = Metric.ptToHwp(dims.mt);
  const mbHwp = Metric.ptToHwp(dims.mb);
  const headerZoneH = dims.headerPt ? Math.max(100, mtHwp - Metric.ptToHwp(dims.headerPt)) : 3600;
  const footerZoneH = dims.footerPt ? Math.max(100, mbHwp - Metric.ptToHwp(dims.footerPt)) : 3600;
  let inner = "";
  const hideFirst = !!(headers.first || footers.first);
  inner += `<hp:ctrl><hp:pageHiding hideHeader="${hideFirst ? 1 : 0}" hideFooter="${hideFirst ? 1 : 0}" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0"/></hp:ctrl>`;
  for (const [type, paras] of Object.entries(headers)) {
    const applyPageType = type === "even" ? "EVEN" : type === "default" || type === "first" ? "BOTH" : "ODD";
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const parasXml = paras.map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml).join("");
    ctx.nextElementId = savedId;
    inner += `<hp:ctrl><hp:header id="1" applyPageType="${applyPageType}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${headerZoneH}" hasTextRef="0" hasNumRef="0">` + parasXml + `</hp:subList></hp:header></hp:ctrl>`;
  }
  for (const [type, paras] of Object.entries(footers)) {
    const applyPageType = type === "even" ? "EVEN" : type === "default" || type === "first" ? "BOTH" : "ODD";
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const parasXml = paras.map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml).join("");
    ctx.nextElementId = savedId;
    inner += `<hp:ctrl><hp:footer id="2" applyPageType="${applyPageType}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${footerZoneH}" hasTextRef="0" hasNumRef="0">` + parasXml + `</hp:subList></hp:footer></hp:ctrl>`;
  }
  return `<hp:run charPrIDRef="0" charTcId="0">${inner}</hp:run>`;
}
function buildSectionXml(sheet, dims, ctx) {
  const secPrXml = buildSecPrXml(dims);
  const kids = sheet?.kids ?? [];
  const hfRunXml = sheet ? buildHeaderFooterRunXml(sheet, dims, ctx) : "";
  const availWidth = Math.max(
    1e3,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr)
  );
  ctx.availableWidth = availWidth;
  let contentXml = "";
  let vertPos = 0;
  for (let i = 0; i < kids.length; i++) {
    const kid = kids[i];
    const isFirst = i === 0;
    const curSecPr = isFirst ? secPrXml : "";
    const curHfRun = isFirst ? hfRunXml : "";
    if (kid.tag === "para") {
      const { xml, nextVertPos } = encodeParaPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        availWidth,
        curHfRun
      );
      contentXml += xml;
      vertPos = nextVertPos;
    } else if (kid.tag === "grid") {
      const { xml, nextVertPos } = encodeGridPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        curHfRun
      );
      contentXml += xml;
      vertPos = nextVertPos;
    }
  }
  if (kids.length === 0) {
    const fs = 1e3;
    const vs = 1600;
    const { xml: linesegXml } = buildLinesegarray(" ", 0, fs, vs / (fs / 100), availWidth);
    contentXml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + secPrXml + hfRunXml + `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>` + linesegXml + `</hp:p>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS} xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">${contentXml}</hs:sec>`;
}
function buildSecPrXml(dims) {
  const wHwp = Metric.ptToHwp(dims.wPt);
  const hHwp = Metric.ptToHwp(dims.hPt);
  const ml = Metric.ptToHwp(dims.ml);
  const mr = Metric.ptToHwp(dims.mr);
  const mt = Metric.ptToHwp(dims.mt);
  const mb = Metric.ptToHwp(dims.mb);
  const headerZone = dims.headerPt ? Math.max(0, mt - Metric.ptToHwp(dims.headerPt)) : 0;
  const footerZone = dims.footerPt ? Math.max(0, mb - Metric.ptToHwp(dims.footerPt)) : 0;
  const pageBorderFill = `<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>`;
  return `<hp:secPr id="0" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${dims.wPt >= dims.hPt ? "WIDELY" : "NARROWLY"}" width="${wHwp}" height="${hHwp}" gutterType="LEFT_ONLY"><hp:margin header="${headerZone}" footer="${footerZone}" gutter="0" left="${ml}" right="${mr}" top="${mt}" bottom="${mb}"/></hp:pagePr><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>` + pageBorderFill + `</hp:secPr>`;
}
function buildLinesegarray(text, vertPosStart, fontSize, lineSpacingPct, horzSize) {
  const vertsizeLine = Math.round(fontSize * lineSpacingPct / 100);
  const spacing = vertsizeLine - fontSize;
  const baseline = Math.round(fontSize * 0.83);
  const isKorean = /[\uAC00-\uD7A3\u3131-\u318E]/.test(text);
  const charW = isKorean ? fontSize : Math.round(fontSize * 0.47);
  const charsPerLine = Math.max(1, Math.floor(horzSize / charW));
  const lineCount = text.length === 0 ? 1 : Math.ceil(text.length / charsPerLine);
  const linesegParts = [];
  for (let i = 0; i < lineCount; i++) {
    const flags = i === 0 ? LINESEG_FLAGS_FIRST : LINESEG_FLAGS_OTHER;
    linesegParts.push(
      `<hp:lineseg textpos="${i * charsPerLine}" vertpos="${vertPosStart + i * vertsizeLine}" vertsize="${vertsizeLine}" textheight="${fontSize}" baseline="${baseline}" spacing="${spacing}" horzpos="0" horzsize="${horzSize}" flags="${flags}"/>`
    );
  }
  return {
    xml: `<hp:linesegarray>${linesegParts.join("")}</hp:linesegarray>`,
    totalHeight: lineCount * vertsizeLine
  };
}
function extractParaText(para) {
  let text = "";
  const walk = (kids) => {
    for (const k of kids) {
      if (k.tag === "span") {
        for (const c of k.kids) if (c.tag === "txt") text += c.content;
      } else if (k.tag === "link") {
        walk(k.kids);
      }
    }
  };
  walk(para.kids);
  return text;
}
function fontSizeForPara(para, ctx) {
  for (const kid of para.kids) {
    if (kid.tag === "span") {
      const id = ctx.charPrMap.get(charPrKey(kid.props));
      if (id !== void 0 && ctx.charPrs[id]) return ctx.charPrs[id].height;
    }
  }
  return 1e3;
}
function encodeParaPositioned(para, ctx, vertPos, secPr = "", availWidth, hfRun = "") {
  const gridKid = para.kids.find((k) => k.tag === "grid");
  if (gridKid) {
    return encodeTablePara(para, gridKid, ctx, vertPos, secPr, hfRun);
  }
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const styleIDRef = para.props.styleId ? ctx.styleIdToHwpxId.get(para.props.styleId) ?? 0 : 0;
  const fontSize = fontSizeForPara(para, ctx);
  const paraPr = ctx.paraPrs[paraPrId];
  const lineSpacing = paraPr?.lineSpacing ?? 160;
  const spacing = Math.max(0, Math.round(fontSize * (lineSpacing / 100 - 1)));
  let vertSize = fontSize + spacing;
  const horzSize = availWidth ?? ctx.availableWidth;
  const isCourierFont = (kids) => kids.some(
    (k) => k.tag === "span" && k.props.font?.toLowerCase().includes("courier") || k.tag === "link" && isCourierFont(k.kids)
  );
  const isCode = availWidth === void 0 && (para.props.styleId?.toLowerCase().includes("code") || isCourierFont(para.kids));
  if (isCode)
    return encodeCodeBlockPositioned(
      para,
      ctx,
      vertPos,
      secPr,
      fontSize,
      spacing,
      vertSize
    );
  let runsXml = encodeParaKids(para.kids, ctx);
  if (!runsXml) runsXml = `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>`;
  const paraText = extractParaText(para);
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText,
    vertPos,
    fontSize,
    lineSpacing,
    horzSize
  );
  const hasPageBreak = para.kids.some(
    (k) => k.tag === "span" && k.kids.some((c) => c.tag === "pb")
  );
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="${styleIDRef}" pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0" paraTcId="0">` + secPr + hfRun + runsXml + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight };
}
function encodeTablePara(para, grid, ctx, vertPos, secPr, hfRun) {
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  const fontSize = 1e3;
  const totalHeight = Math.max(1600, tblHeight);
  const baseline = 850;
  const spacing = Math.max(0, totalHeight - fontSize);
  const linesegXml = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" textheight="${fontSize}" baseline="${baseline}" spacing="${spacing}" horzpos="0" horzsize="${ctx.availableWidth}" flags="1441792"/></hp:linesegarray>`;
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + secPr + gridXml + hfRun + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight };
}
function encodeCodeBlockPositioned(para, ctx, vertPos, secPr, fontSize, spacing, vertSize) {
  const codeBfId = ctx.borderFillBank.addUniform(
    { kind: "solid", pt: 0.5, color: "aaaaaa" },
    "f4f4f4"
  );
  const cellW = ctx.availableWidth;
  const innerW = Math.max(cellW - 510, 100);
  const subListId = ctx.nextElementId++;
  const { xml: innerXml } = encodeParaPositioned(para, ctx, 0, "", innerW);
  const paraText = extractParaText(para);
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText,
    vertPos,
    fontSize,
    160,
    // 코드 블록 기본 줄간격 160%
    ctx.availableWidth
  );
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" paraTcId="0">` + secPr + `<hp:run charPrIDRef="0" charTcId="0"><hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="${codeBfId}" noAdjust="0"><hp:sz width="${cellW}" widthRelTo="ABSOLUTE" height="0" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="138" right="138" top="138" bottom="138"/><hp:inMargin left="138" right="138" top="138" bottom="138"/><hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${codeBfId}"><hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` + innerXml + `</hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${cellW}" height="0"/><hp:cellMargin left="283" right="283" top="141" bottom="141"/></hp:tc></hp:tr></hp:tbl><hp:t xml:space="preserve"> </hp:t></hp:run>` + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight };
}
function encodeParaKids(kids, ctx) {
  let xml = "";
  let currentRunCharPrId = null;
  let currentRunContent = "";
  const flushRun = () => {
    if (currentRunCharPrId !== null) {
      const content = currentRunContent || `<hp:t xml:space="preserve"> </hp:t>`;
      xml += `<hp:run charPrIDRef="${currentRunCharPrId}" charTcId="0">${content}</hp:run>`;
    }
    currentRunCharPrId = null;
    currentRunContent = "";
  };
  for (const kid of kids) {
    if (kid.tag === "span") {
      const span = kid;
      const charPrId = ctx.charPrMap.get(charPrKey(span.props)) ?? 0;
      if (currentRunCharPrId !== null && currentRunCharPrId !== charPrId) {
        flushRun();
      }
      currentRunCharPrId = charPrId;
      currentRunContent += encodeRunInner(span);
    } else if (kid.tag === "link") {
      const link = kid;
      let charPrId = 0;
      if (link.kids.length > 0 && link.kids[0].tag === "span") {
        charPrId = ctx.charPrMap.get(charPrKey(link.kids[0].props)) ?? 0;
      }
      if (currentRunCharPrId !== null && currentRunCharPrId !== charPrId) {
        flushRun();
      }
      currentRunCharPrId = charPrId;
      currentRunContent += encodeLinkInner(link, ctx);
    } else if (kid.tag === "img") {
      flushRun();
      xml += encodeImgWrapped(kid, ctx);
    }
  }
  flushRun();
  return xml;
}
function encodeRunInner(span) {
  let xml = "";
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const content = esc(kid.content);
      if (content) xml += `<hp:t xml:space="preserve">${content}</hp:t>`;
    } else if (kid.tag === "br") {
      xml += `<hp:t xml:space="preserve">
</hp:t>`;
    } else if (kid.tag === "pagenum") {
      const fmt = kid.format === "roman" ? "ROMAN_LOWER" : kid.format === "romanCaps" ? "ROMAN_UPPER" : "DIGIT";
      xml += `<hp:pageNum pageStartsOn="BOTH" formatType="${fmt}"/>`;
    }
  }
  return xml;
}
function encodeLinkInner(link, ctx) {
  const fieldId = 6e8 + ctx.nextElementId++ % 1e8;
  const instanceId = 21e8 + ctx.nextElementId++ % 1e8;
  const url = link.href;
  let xml = `<hp:ctrl><hp:fieldBegin id="${instanceId}" type="HYPERLINK" name="" editable="0" dirty="1" zorder="-1" fieldid="${fieldId}"><hp:parameters cnt="6" name=""><hp:integerParam name="Prop">0</hp:integerParam><hp:stringParam name="Command">${esc(url.replace(/:/g, "\\:"))};1;5;-1;</hp:stringParam><hp:stringParam name="Path">${esc(url)}</hp:stringParam><hp:stringParam name="Category">HWPHYPERLINK_TYPE_URL</hp:stringParam><hp:stringParam name="TargetType">HWPHYPERLINK_TARGET_HYPERLINK</hp:stringParam><hp:stringParam name="DocOpenType">HWPHYPERLINK_JUMP_DONTCARE</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl>`;
  for (const kid of link.kids) {
    if (kid.tag === "span") {
      xml += encodeRunInner(kid);
    }
  }
  xml += `<hp:ctrl><hp:fieldEnd beginIDRef="${instanceId}"/></hp:ctrl>`;
  return xml;
}
var WRAP_MAP = {
  inline: "TOP_AND_BOTTOM",
  square: "SQUARE",
  tight: "BOTH_SIDES",
  through: "BOTH_SIDES",
  none: "FRONT_TEXT",
  behind: "BEHIND_TEXT",
  front: "FRONT_TEXT"
};
var FLOW_MAP = {
  inline: "BOTH_SIDES",
  square: "LARGEST_ONLY",
  tight: "BOTH_SIDES",
  through: "BOTH_SIDES",
  none: "BOTH_SIDES",
  behind: "BOTH_SIDES",
  front: "BOTH_SIDES"
};
function encodeImage(img, ctx) {
  if (!img.b64) {
    return `<hp:t xml:space="preserve">${esc(img.alt || "[\uAC1C\uCCB4]")}</hp:t>`;
  }
  const binId = ctx.imgMap.get(img);
  if (!binId) return "";
  const pixelDims = readPixelDims(img.b64, img.mime);
  let wHwp, hHwp;
  if (pixelDims && pixelDims.w > 0 && pixelDims.h > 0) {
    wHwp = Metric.ptToHwp(pixelDims.w * 72 / 96);
    hHwp = Metric.ptToHwp(pixelDims.h * 72 / 96);
  } else {
    wHwp = Metric.ptToHwp(img.w);
    hHwp = Metric.ptToHwp(img.h);
  }
  if (wHwp > ctx.availableWidth) {
    hHwp = Math.round(hHwp * ctx.availableWidth / wHwp);
    wHwp = ctx.availableWidth;
  }
  const rotationCenterX = Math.round(wHwp / 2);
  const rotationCenterY = Math.round(hHwp / 2);
  const layout = img.layout;
  const isInline = !layout || layout.wrap === "inline";
  const textWrap = layout ? WRAP_MAP[layout.wrap] ?? "SQUARE" : "SQUARE";
  const textFlow = layout ? FLOW_MAP[layout.wrap] ?? "BOTH_SIDES" : "BOTH_SIDES";
  const zOrder = ctx.nextZOrder++;
  return `<hp:pic id="${ctx.nextElementId++}" zOrder="${zOrder}" numberingType="PICTURE" textWrap="${textWrap}" textFlow="${textFlow}" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="0" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${wHwp}" height="${hHwp}"/><hp:curSz width="${wHwp}" height="${hHwp}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="${rotationCenterX}" centerY="${rotationCenterY}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${wHwp}" y="0"/><hc:pt2 x="${wHwp}" y="${hHwp}"/><hc:pt3 x="0" y="${hHwp}"/></hp:imgRect><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="${wHwp}" dimheight="${hHwp}"/><hc:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:effects/><hp:sz width="${wHwp}" widthRelTo="ABSOLUTE" height="${hHwp}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="${isInline ? 1 : 0}" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic>`;
}
function encodeImgWrapped(img, ctx) {
  const content = encodeImage(img, ctx);
  if (!img.b64) {
    return `<hp:run charPrIDRef="0" charTcId="0">${content}</hp:run>`;
  }
  return `<hp:run charPrIDRef="0" charTcId="0">${content}<hp:t xml:space="preserve"> </hp:t></hp:run>`;
}
function encodeGridPositioned(grid, ctx, vertPos, secPr = "", hfRun = "") {
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  const totalHeight = Math.max(1600, tblHeight);
  const fontSize = 1e3;
  const baseline = Math.round(fontSize * 0.83);
  const spacing = Math.max(0, totalHeight - fontSize);
  const linesegXml = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" textheight="${fontSize}" baseline="${baseline}" spacing="${spacing}" horzpos="0" horzsize="${ctx.availableWidth}" flags="${LINESEG_FLAGS_FIRST}"/></hp:linesegarray>`;
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + secPr + hfRun + gridXml + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight };
}
function buildGridXml(grid, ctx) {
  const rowCount = grid.kids.length;
  const tableMap = Array.from({ length: rowCount }, () => []);
  for (let ri = 0; ri < rowCount; ri++) {
    let ci = 0;
    for (const cell of grid.kids[ri].kids) {
      while (tableMap[ri][ci]) ci++;
      tableMap[ri][ci] = { type: "real", cell };
      for (let rr = 0; rr < cell.rs; rr++) {
        const tri = ri + rr;
        if (tri >= rowCount) break;
        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue;
          tableMap[tri][ci + cc] = { type: "absorbed" };
        }
      }
      ci += cell.cs;
    }
  }
  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++)
    colCount = Math.max(colCount, tableMap[ri].length);
  if (colCount === 0) colCount = 1;
  const totalW = ctx.availableWidth;
  const colWidths = [];
  if (grid.props.colWidths && grid.props.colWidths.length === colCount) {
    for (const wPt of grid.props.colWidths) {
      colWidths.push(Metric.ptToHwp(wPt));
    }
  } else {
    const defW = Math.round(totalW / colCount);
    for (let c = 0; c < colCount; c++) colWidths.push(defW);
  }
  const rawTotal = colWidths.reduce((s, w) => s + w, 0);
  if (rawTotal > totalW && rawTotal > 0) {
    const scale = totalW / rawTotal;
    for (let i = 0; i < colWidths.length; i++) {
      colWidths[i] = Math.max(100, Math.round(colWidths[i] * scale));
    }
  }
  const actualTotal = colWidths.reduce((s, w) => s + w, 0);
  const rowHeights = [];
  for (let ri = 0; ri < rowCount; ri++) {
    if (grid.kids[ri].heightPt != null && grid.kids[ri].heightPt > 0) {
      rowHeights.push(Metric.ptToHwp(grid.kids[ri].heightPt));
    } else {
      let maxH = 0;
      for (let ci = 0; ci < colCount; ci++) {
        const entry = tableMap[ri][ci];
        if (entry?.type === "real") {
          const h = estimateCellHeight(entry.cell, ctx);
          if (h > maxH) maxH = h;
        }
      }
      rowHeights.push(maxH || Math.round(1e3 * 1.6));
    }
  }
  const totalH = rowHeights.reduce((s, h) => s + h, 0);
  const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
  const tblBfId = ctx.borderFillBank.addUniform(defStroke);
  let rowsXml = "";
  for (let ri = 0; ri < rowCount; ri++) {
    let cellsXml = "";
    for (let ci = 0; ci < colCount; ci++) {
      const entry = tableMap[ri][ci];
      if (!entry || entry.type === "absorbed") continue;
      const cell = entry.cell;
      const cp = cell.props;
      const cellBfId = ctx.borderFillBank.addFromCellProps(cp, defStroke);
      let cellW = 0;
      for (let sc = ci; sc < ci + cell.cs && sc < colWidths.length; sc++)
        cellW += colWidths[sc];
      if (!cellW) cellW = Math.round(totalW / colCount) * cell.cs;
      const subListId = ctx.nextElementId++;
      const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : 141;
      const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : 141;
      const padT = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : 141;
      const padB = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : 141;
      const innerW = Math.max(cellW - padL - padR, 100);
      let parasXml = "";
      if (cell.kids.length > 0) {
        for (const kid of cell.kids) {
          if (kid.tag === "grid") {
            const { xml: tblXml } = buildGridXml(kid, ctx);
            const pid = ctx.nextElementId++;
            const rid = ctx.nextElementId++;
            parasXml += `<hp:p id="${pid}" paraPrIDRef="0" styleIDRef="0" paraTcId="0"><hp:run id="${rid}" charPrIDRef="0" charTcId="0">${tblXml}</hp:run></hp:p>`;
          } else {
            parasXml += encodeParaPositioned(kid, ctx, 0, "", innerW).xml;
          }
        }
      } else {
        parasXml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" paraTcId="0"><hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run></hp:p>`;
      }
      const vAlign = cp.va === "mid" ? "CENTER" : cp.va === "bot" ? "BOTTOM" : "TOP";
      cellsXml += `<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}"><hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${cellW}" height="${rowHeights[ri]}"/><hp:cellMargin left="${padL}" right="${padR}" top="${padT}" bottom="${padB}"/><hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vAlign}" linkListIDRef="0" linkListNextIDRef="0" textWidth="${innerW}" textHeight="${Math.max(100, rowHeights[ri] - padT - padB)}" hasTextRef="0" hasNumRef="0">` + parasXml + `</hp:subList></hp:tc>`;
    }
    rowsXml += `<hp:tr>${cellsXml}</hp:tr>`;
  }
  const alignMap = {
    left: "LEFT",
    right: "RIGHT",
    center: "CENTER",
    justify: "JUSTIFY"
  };
  const horzAlign = alignMap[grid.props.align ?? "left"] ?? "LEFT";
  const headerRow = grid.props.headerRow ? ' repeatHeader="1"' : "";
  const xml = `<hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE"${headerRow} rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="0"><hp:sz width="${actualTotal}" widthRelTo="ABSOLUTE" height="${totalH}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="${horzAlign}" vertOffset="0" horzOffset="0"/><hp:outMargin left="138" right="138" top="138" bottom="138"/><hp:inMargin left="0" right="0" top="0" bottom="0"/>` + rowsXml + `</hp:tbl>`;
  return { xml, height: totalH };
}
function estimateCellHeight(cell, ctx) {
  const topPad = 141;
  const botPad = 141;
  let h = 0;
  for (const kid of cell.kids) {
    if (kid.tag === "grid") {
      h += 1600;
      continue;
    }
    const para = kid;
    const fs = fontSizeForPara(para, ctx);
    const ppId = ctx.paraPrMap.get(paraPrKey(para.props));
    const pp = ppId !== void 0 ? ctx.paraPrs[ppId] : null;
    const ls = pp?.lineSpacing ?? 160;
    const before = pp?.prevHwp ?? 0;
    const after = pp?.nextHwp ?? 0;
    h += Math.round(fs * ls / 100) + before + after;
  }
  if (!h) h = Math.round(1e3 * 1.6);
  return h + topPad + botPad;
}
function extractPreviewText(sheet) {
  if (!sheet) return "";
  const lines = [];
  for (const kid of sheet.kids) {
    if (kid.tag === "para") {
      const text = kid.kids.flatMap(
        (k) => k.tag === "span" ? k.kids.flatMap((c) => c.tag === "txt" ? [c.content] : []) : []
      ).join("");
      if (text) lines.push(text);
    } else if (kid.tag === "grid") {
      for (const row of kid.kids) {
        const cells = row.kids.map(
          (cell) => cell.kids.flatMap(
            (p) => p.tag === "para" ? p.kids.flatMap(
              (k) => k.tag === "span" ? k.kids.flatMap((c) => c.tag === "txt" ? [c.content] : []) : []
            ) : []
          ).join("")
        );
        lines.push(cells.join("	"));
      }
    }
  }
  return lines.join("\r\n");
}
function esc(s) {
  if (!s) return "";
  s = s.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
  s = s.replace(/湰灧/g, "").replace(/\uFEFF/g, "");
  s = s.replace(
    /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu,
    ""
  );
  return TextKit.escapeXml(s);
}
registry.registerEncoder(new HwpxEncoder());

// src/encoders/docx/DocxEncoder.ts
var DocxEncoder = class extends BaseEncoder {
  getFormat() {
    return "docx";
  }
  async encode(doc) {
    try {
      const sheets = doc.kids.length > 0 ? doc.kids : [];
      const firstSheet = sheets[0];
      const dims = normalizeDims(firstSheet?.dims ?? A4);
      const allKids = sheets.flatMap((s) => s?.kids ?? []);
      const images = [];
      const ctx = {
        images,
        nextId: 10,
        nextImgNum: 1,
        warns: [],
        imgMap: /* @__PURE__ */ new WeakMap()
      };
      collectImages(allKids, ctx);
      let headerParas = firstSheet?.headers?.default;
      let footerParas = firstSheet?.footers?.default;
      const hasHeader = headerParas && headerParas.length > 0;
      const hasFooter = footerParas && footerParas.length > 0;
      if (hasHeader) collectImagesFromParas(headerParas, ctx);
      if (hasFooter) collectImagesFromParas(footerParas, ctx);
      const headerRId = hasHeader ? `rId${ctx.nextId++}` : "";
      const footerRId = hasFooter ? `rId${ctx.nextId++}` : "";
      const numInfo = collectNumbering(allKids);
      const kids = allKids;
      const entries = [
        {
          name: "[Content_Types].xml",
          data: this.stringToBytes(contentTypes(images, hasHeader, hasFooter))
        },
        { name: "_rels/.rels", data: this.stringToBytes(pkgRels()) },
        {
          name: "word/document.xml",
          data: this.stringToBytes(
            documentXml(kids, dims, ctx, headerRId, footerRId)
          )
        },
        { name: "word/styles.xml", data: this.stringToBytes(stylesXml()) },
        { name: "word/settings.xml", data: this.stringToBytes(settingsXml()) },
        {
          name: "word/_rels/document.xml.rels",
          data: this.stringToBytes(
            docRels(images, headerRId, footerRId, numInfo.hasLists)
          )
        },
        { name: "docProps/app.xml", data: this.stringToBytes(appXml()) },
        {
          name: "docProps/core.xml",
          data: this.stringToBytes(coreXml(doc.meta))
        }
      ];
      if (numInfo.hasLists) {
        entries.push({
          name: "word/numbering.xml",
          data: this.stringToBytes(numberingXml(numInfo))
        });
      }
      if (hasHeader) {
        entries.push({
          name: "word/header1.xml",
          data: this.stringToBytes(headerFooterXml("hdr", headerParas, ctx))
        });
      }
      if (hasFooter) {
        entries.push({
          name: "word/footer1.xml",
          data: this.stringToBytes(headerFooterXml("ftr", footerParas, ctx))
        });
      }
      for (const img of images) {
        entries.push({ name: `word/media/${img.name}`, data: img.data });
      }
      return succeed(await this.zip(entries));
    } catch (e) {
      return fail(`DOCX encode error: ${e?.message ?? String(e)}`);
    }
  }
};
function mimeToExt2(mime) {
  if (mime.includes("jpeg")) return "jpeg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  return "png";
}
function collectImages(kids, ctx) {
  for (const kid of kids) {
    if (kid.tag === "para") collectImagesFromPara(kid, ctx);
    else if (kid.tag === "grid") {
      for (const row of kid.kids)
        for (const cell of row.kids)
          for (const p of cell.kids)
            if (p.tag === "para") collectImagesFromPara(p, ctx);
            else collectImages([p], ctx);
    }
  }
}
function collectImagesFromParas(paras, ctx) {
  for (const p of paras) collectImagesFromPara(p, ctx);
}
function collectImagesFromPara(para, ctx) {
  for (const kid of para.kids) {
    if (kid.tag === "img") registerImage2(kid, ctx);
  }
}
function registerImage2(img, ctx) {
  if (ctx.imgMap.has(img)) return;
  const ext = mimeToExt2(img.mime);
  const name = `image${ctx.nextImgNum++}.${ext}`;
  const rId = `rId${ctx.nextId++}`;
  const data = TextKit.base64Decode(img.b64);
  ctx.images.push({ rId, name, data, ext });
  ctx.imgMap.set(img, rId);
}
function collectNumbering(kids) {
  let hasBullet = false;
  let hasNumbered = false;
  for (const kid of kids) {
    if (kid.tag === "para") {
      if (kid.props.listOrd === true) hasNumbered = true;
      else if (kid.props.listOrd === false) hasBullet = true;
    }
  }
  return { hasLists: hasBullet || hasNumbered, hasBullet, hasNumbered };
}
function contentTypes(images, hasHeader, hasFooter) {
  const imgDefaults = /* @__PURE__ */ new Set();
  for (const img of images) imgDefaults.add(img.ext);
  let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>`;
  for (const ext of imgDefaults) {
    const ct = ext === "png" ? "image/png" : ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/bmp";
    defaults += `
  <Default Extension="${ext}" ContentType="${ct}"/>`;
  }
  let overrides = `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;
  if (hasHeader)
    overrides += `
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
  if (hasFooter)
    overrides += `
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  ${defaults}
  ${overrides}
</Types>`;
}
function pkgRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}
function docRels(images, headerRId, footerRId, hasLists) {
  let rels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`;
  if (hasLists) {
    rels += `
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
  }
  for (const img of images) {
    rels += `
  <Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`;
  }
  if (headerRId) {
    rels += `
  <Relationship Id="${headerRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`;
  }
  if (footerRId) {
    rels += `
  <Relationship Id="${footerRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`;
}
function appXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>hwpkit</Application>
</Properties>`;
}
function coreXml(meta) {
  const now = (/* @__PURE__ */ new Date()).toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${esc2(meta.title ?? "")}</dc:title>
  <dc:creator>${esc2(meta.author ?? "")}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${meta.created ?? now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}
function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="\uB9D1\uC740 \uACE0\uB515" w:eastAsia="\uB9D1\uC740 \uACE0\uB515" w:hAnsi="\uB9D1\uC740 \uACE0\uB515"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
}
function settingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:bordersDoNotSurroundHeader/>
  <w:bordersDoNotSurroundFooter/>
  <w:defaultTabStop w:val="800"/>
  <w:compat>
    <w:spaceForUL/>
    <w:balanceSingleByteDoubleByteWidth/>
    <w:doNotLeaveBackslashAlone/>
    <w:ulTrailSpace/>
    <w:doNotExpandShiftReturn/>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
  </w:compat>
</w:settings>`;
}
function numberingXml(info) {
  let abstractNums = "";
  let nums = "";
  if (info.hasBullet) {
    abstractNums += `<w:abstractNum w:abstractNumId="0">`;
    for (let lvl = 0; lvl < 9; lvl++) {
      const marker = lvl === 0 ? "\u25CF" : lvl === 1 ? "\u25CB" : "\u25A0";
      const indent = (lvl + 1) * 720;
      abstractNums += `<w:lvl w:ilvl="${lvl}"><w:numFmt w:val="bullet"/><w:lvlText w:val="${marker}"/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    abstractNums += `</w:abstractNum>`;
    nums += `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
  }
  if (info.hasNumbered) {
    abstractNums += `<w:abstractNum w:abstractNumId="1">`;
    for (let lvl = 0; lvl < 9; lvl++) {
      const fmt = lvl % 3 === 0 ? "decimal" : lvl % 3 === 1 ? "lowerLetter" : "lowerRoman";
      const indent = (lvl + 1) * 720;
      abstractNums += `<w:lvl w:ilvl="${lvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${lvl + 1}."/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    abstractNums += `</w:abstractNum>`;
    nums += `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${abstractNums}
  ${nums}
</w:numbering>`;
}
function headerFooterXml(type, paras, ctx) {
  const tag = type === "hdr" ? "w:hdr" : "w:ftr";
  const body = paras.map((p) => encodeParaInner(p, ctx)).join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
${body}
</${tag}>`;
}
function documentXml(kids, dims, ctx, headerRId, footerRId) {
  const body = kids.map((k) => encodeContent(k, ctx, dims)).join("\n");
  let sectRefs = "";
  if (headerRId)
    sectRefs += `
      <w:headerReference w:type="default" r:id="${headerRId}"/>`;
  if (footerRId)
    sectRefs += `
      <w:footerReference w:type="default" r:id="${footerRId}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${body}
    <w:sectPr>${sectRefs}
      <w:pgSz w:w="${Metric.ptToDxa(dims.wPt)}" w:h="${Metric.ptToDxa(dims.hPt)}" w:orient="${dims.orient ?? "portrait"}"/>
      <w:pgMar w:top="${Metric.ptToDxa(dims.mt)}" w:right="${Metric.ptToDxa(dims.mr)}" w:bottom="${Metric.ptToDxa(dims.mb)}" w:left="${Metric.ptToDxa(dims.ml)}" w:header="709" w:footer="709" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}
function encodeContent(node, ctx, dims) {
  return node.tag === "grid" ? encodeGrid(node, ctx, dims) : encodeParaInner(node, ctx);
}
function encodeParaInner(para, ctx) {
  const align = para.props.align ?? "left";
  const headStyle = para.props.heading ? `<w:pStyle w:val="Heading${para.props.heading}"/>` : "";
  let numPr = "";
  if (para.props.listOrd !== void 0) {
    const numId = para.props.listOrd ? 2 : 1;
    const ilvl = para.props.listLv ?? 0;
    numPr = `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
  }
  let spacingXml = "";
  const { spaceBefore, spaceAfter, lineHeight, lineHeightFixed } = para.props;
  if (spaceBefore !== void 0 || spaceAfter !== void 0 || lineHeight !== void 0 || lineHeightFixed !== void 0) {
    const parts = [];
    if (spaceBefore !== void 0)
      parts.push(`w:before="${Math.max(0, Metric.ptToDxa(spaceBefore))}"`);
    if (spaceAfter !== void 0)
      parts.push(`w:after="${Math.max(0, Metric.ptToDxa(spaceAfter))}"`);
    if (lineHeightFixed !== void 0) {
      parts.push(
        `w:line="${Math.max(1, Metric.ptToDxa(lineHeightFixed))}" w:lineRule="exact"`
      );
    } else if (lineHeight !== void 0) {
      parts.push(`w:line="${Math.round(lineHeight * 240)}" w:lineRule="auto"`);
    }
    spacingXml = `<w:spacing ${parts.join(" ")}/>`;
  }
  let indentXml = "";
  const leftDxa = Math.round(Metric.ptToDxa(para.props.indentPt ?? 0));
  const rightDxa = Math.round(Metric.ptToDxa(para.props.indentRightPt ?? 0));
  const firstPt = para.props.firstLineIndentPt ?? 0;
  const indParts = [];
  if (leftDxa > 0) indParts.push(`w:left="${leftDxa}"`);
  if (rightDxa > 0) indParts.push(`w:right="${rightDxa}"`);
  if (firstPt > 0)
    indParts.push(`w:firstLine="${Math.round(Metric.ptToDxa(firstPt))}"`);
  if (firstPt < 0)
    indParts.push(`w:hanging="${Math.round(Metric.ptToDxa(-firstPt))}"`);
  if (indParts.length > 0) indentXml = `<w:ind ${indParts.join(" ")}/>`;
  const runs = para.kids.map((k) => {
    if (k.tag === "span") return encodeRun(k, ctx);
    if (k.tag === "img") return encodeImage2(k, ctx);
    return "";
  }).join("");
  return `    <w:p>
      <w:pPr>${headStyle}${numPr}${spacingXml}${indentXml}<w:jc w:val="${align === "justify" ? "both" : align}"/></w:pPr>
      ${runs}
    </w:p>`;
}
function encodeRun(span, _ctx) {
  const p = span.props;
  const rPr = [];
  if (p.b) rPr.push("<w:b/>");
  if (p.i) rPr.push("<w:i/>");
  if (p.u) rPr.push('<w:u w:val="single"/>');
  if (p.s) rPr.push("<w:strike/>");
  if (p.sup) rPr.push('<w:vertAlign w:val="superscript"/>');
  if (p.sub) rPr.push('<w:vertAlign w:val="subscript"/>');
  if (p.pt)
    rPr.push(
      `<w:sz w:val="${Metric.ptToHalfPt(p.pt)}"/><w:szCs w:val="${Metric.ptToHalfPt(p.pt)}"/>`
    );
  if (p.color) rPr.push(`<w:color w:val="${p.color}"/>`);
  if (p.font)
    rPr.push(
      `<w:rFonts w:ascii="${esc2(p.font)}" w:hAnsi="${esc2(p.font)}" w:eastAsia="${esc2(p.font)}"/>`
    );
  if (p.bg) rPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`);
  const parts = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const content = kid.content.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      if (content) {
        parts.push(
          `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:t xml:space="preserve">${esc2(content)}</w:t></w:r>`
        );
      }
    } else if (kid.tag === "pagenum") {
      parts.push(
        `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:instrText> PAGE </w:instrText></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:t>1</w:t></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`
      );
    } else if (kid.tag === "br") {
      parts.push(`<w:r><w:br/></w:r>`);
    } else if (kid.tag === "pb") {
      parts.push(`<w:r><w:br w:type="page"/></w:r>`);
    }
  }
  return parts.join("");
}
function encodeImage2(img, ctx) {
  const rId = ctx.imgMap.get(img);
  if (!rId) return "";
  const cx = Metric.ptToEmu(img.w > 0 ? img.w : 72);
  const cy = Metric.ptToEmu(img.h > 0 ? img.h : 72);
  const alt = esc2(img.alt ?? "");
  const docPrId = ctx.nextId++;
  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const layout = img.layout;
  const isInline = !layout || layout.wrap === "inline";
  const forceAnchor = layout?.wrap === "topAndBottom" || layout?.wrap === "square" || layout?.wrap === "tight" || layout?.wrap === "behind" || layout?.wrap === "front";
  if (isInline && !forceAnchor) {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:inline></w:drawing></w:r>`;
  }
  return `<w:r><w:drawing>${encodeAnchor(img, cx, cy, alt, docPrId, graphic, layout)}</w:drawing></w:r>`;
}
function encodeAnchor(_img, cx, cy, alt, docPrId, graphic, layout) {
  const distT = Metric.ptToEmu(layout.distT ?? 0);
  const distB = Metric.ptToEmu(layout.distB ?? 0);
  const distL = Metric.ptToEmu(layout.distL ?? 9144);
  const distR = Metric.ptToEmu(layout.distR ?? 9144);
  const behindDoc = layout.behindDoc || layout.wrap === "behind" ? "1" : "0";
  const relH = layout.zOrder ?? 251658240;
  const horzRelFrom = HORZ_RELTO_DOCX[layout.horzRelTo ?? "column"] ?? "column";
  let posH;
  if (layout.xPt != null) {
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.xPt)}</wp:posOffset></wp:positionH>`;
  } else {
    const ha = HORZ_ALIGN_DOCX[layout.horzAlign ?? "left"] ?? "left";
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:align>${ha}</wp:align></wp:positionH>`;
  }
  const vertRelFrom = VERT_RELTO_DOCX[layout.vertRelTo ?? "para"] ?? "paragraph";
  let posV;
  if (layout.yPt != null) {
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.yPt)}</wp:posOffset></wp:positionV>`;
  } else {
    const va = VERT_ALIGN_DOCX[layout.vertAlign ?? "top"] ?? "top";
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:align>${va}</wp:align></wp:positionV>`;
  }
  const wrapXml = WRAP_DOCX[layout.wrap] ?? '<wp:wrapSquare wrapText="bothSides"/>';
  return `<wp:anchor distT="${distT}" distB="${distB}" distL="${distL}" distR="${distR}" simplePos="0" relativeHeight="${relH}" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}<wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:anchor>`;
}
var HORZ_RELTO_DOCX = {
  margin: "margin",
  column: "column",
  page: "page",
  para: "paragraph"
};
var VERT_RELTO_DOCX = {
  margin: "margin",
  line: "line",
  page: "page",
  para: "paragraph"
};
var HORZ_ALIGN_DOCX = {
  left: "left",
  center: "center",
  right: "right"
};
var VERT_ALIGN_DOCX = {
  top: "top",
  center: "center",
  bottom: "bottom"
};
var WRAP_DOCX = {
  square: '<wp:wrapSquare wrapText="bothSides"/>',
  tight: '<wp:wrapTight><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>',
  through: '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>',
  // ECMA-376 §20.4.2.15: wrapTopAndBottom — 텍스트가 이미지 위아래로만 흐름
  topAndBottom: "<wp:wrapTopAndBottom/>",
  none: "<wp:wrapNone/>",
  behind: "<wp:wrapNone/>",
  front: "<wp:wrapNone/>"
};
function encodeGrid(grid, ctx, dims = A4) {
  const gp = grid.props;
  const look = gp.look;
  const firstRow = look?.firstRow ? "1" : "0";
  const lastRow = look?.lastRow ? "1" : "0";
  const firstCol = look?.firstCol ? "1" : "0";
  const lastCol = look?.lastCol ? "1" : "0";
  const noHBand = look?.bandedRows ? "0" : "1";
  const noVBand = look?.bandedCols ? "0" : "1";
  const d = dims ?? A4;
  const availDxa = Metric.ptToDxa(d.wPt - d.ml - d.mr);
  const tableMap = Array.from(
    { length: grid.kids.length },
    () => []
  );
  for (let ri = 0; ri < grid.kids.length; ri++) {
    let c = 0;
    for (const cell of grid.kids[ri].kids) {
      while (tableMap[ri][c]) c++;
      tableMap[ri][c] = { type: "real", cell, width: cell.cs };
      for (let rr = 0; rr < cell.rs; rr++) {
        const targetRi = ri + rr;
        if (targetRi >= grid.kids.length) break;
        if (!tableMap[targetRi]) tableMap[targetRi] = [];
        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue;
          if (rr > 0 && cc === 0) {
            tableMap[targetRi][c + cc] = { type: "continue", width: cell.cs };
          } else {
            tableMap[targetRi][c + cc] = { type: "absorbed" };
          }
        }
      }
      c += cell.cs;
    }
  }
  let colCount = 0;
  for (let ri = 0; ri < grid.kids.length; ri++) {
    colCount = Math.max(colCount, tableMap[ri].length);
  }
  if (colCount === 0) colCount = 1;
  for (let ri = 0; ri < grid.kids.length; ri++) {
    for (let c = 0; c < colCount; c++) {
      if (!tableMap[ri][c]) tableMap[ri][c] = { type: "void" };
    }
  }
  const defaultColDxa = Math.round(availDxa / colCount);
  let colWidthsDxa = [];
  if (grid.props.colWidths && grid.props.colWidths.length > 0) {
    const srcPt = [...grid.props.colWidths];
    while (srcPt.length < colCount) srcPt.push(0);
    srcPt.length = colCount;
    const knownTotalPt = srcPt.filter((w) => w > 0).reduce((s, w) => s + w, 0);
    const zeroCount = srcPt.filter((w) => w <= 0).length;
    const availPt = Metric.dxaToPt(availDxa);
    const remainingPt = Math.max(0, availPt - knownTotalPt);
    const zeroFillPt = zeroCount > 0 ? remainingPt / zeroCount : 0;
    for (let i = 0; i < srcPt.length; i++) {
      if (srcPt[i] <= 0) {
        srcPt[i] = zeroFillPt > 0 ? zeroFillPt : availPt / colCount;
      }
    }
    colWidthsDxa = srcPt.map((w) => Math.round(Metric.ptToDxa(w)));
    const computedTotalDxa = colWidthsDxa.reduce((s, w) => s + w, 0);
    if (computedTotalDxa > availDxa) {
      const scale = availDxa / computedTotalDxa;
      colWidthsDxa = colWidthsDxa.map((w) => Math.round(w * scale));
    }
  } else {
    for (let c = 0; c < colCount; c++) colWidthsDxa.push(defaultColDxa);
  }
  const totalDxa = colWidthsDxa.reduce((s, w) => s + w, 0);
  const gridCols = colWidthsDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const rows = tableMap.map((rowMap, ri) => {
    const cellXmls = [];
    for (let c = 0; c < colCount; c++) {
      const mapEntry = rowMap[c];
      if (mapEntry.type === "absorbed") continue;
      const isContinue = mapEntry.type === "continue";
      const isReal = mapEntry.type === "real";
      const isVoid = mapEntry.type === "void";
      if (isContinue || isReal || isVoid) {
        let cw = 0;
        const cellWidth = mapEntry.width || 1;
        const safeColWidths = colWidthsDxa.length >= colCount ? colWidthsDxa : [
          ...colWidthsDxa,
          ...Array(colCount - colWidthsDxa.length).fill(defaultColDxa)
        ];
        for (let sc = c; sc < c + cellWidth && sc < safeColWidths.length; sc++) {
          cw += safeColWidths[sc];
        }
        if (cw <= 0) cw = defaultColDxa * cellWidth;
        const tcPrParts = [];
        tcPrParts.push(`<w:tcW w:w="${Math.round(cw)}" w:type="dxa"/>`);
        if (cellWidth > 1) {
          tcPrParts.push(`<w:gridSpan w:val="${cellWidth}"/>`);
        }
        if (isContinue) {
          tcPrParts.push(`<w:vMerge/>`);
        }
        let cellContent = "";
        if (isReal) {
          const cell = mapEntry.cell;
          const cp = cell.props;
          if (cell.rs > 1) tcPrParts.push(`<w:vMerge w:val="restart"/>`);
          const borders = encodeCellBorders(cp);
          if (borders) tcPrParts.push(borders);
          if (cp.bg)
            tcPrParts.push(
              `<w:shd w:val="clear" w:color="auto" w:fill="${cp.bg}"/>`
            );
          if (cp.va) {
            const vaMap = {
              top: "top",
              mid: "center",
              bot: "bottom"
            };
            tcPrParts.push(`<w:vAlign w:val="${vaMap[cp.va] ?? "top"}"/>`);
          }
          const cPadT = cp.padT != null ? Math.round(Metric.ptToDxa(cp.padT)) : null;
          const cPadB = cp.padB != null ? Math.round(Metric.ptToDxa(cp.padB)) : null;
          const cPadL = cp.padL != null ? Math.round(Metric.ptToDxa(cp.padL)) : null;
          const cPadR = cp.padR != null ? Math.round(Metric.ptToDxa(cp.padR)) : null;
          if (cPadT != null || cPadB != null || cPadL != null || cPadR != null) {
            const t = cPadT ?? 28;
            const b = cPadB ?? 28;
            const l = cPadL ?? 72;
            const r = cPadR ?? 72;
            tcPrParts.push(
              `<w:tcMar><w:top w:w="${t}" w:type="dxa"/><w:left w:w="${l}" w:type="dxa"/><w:bottom w:w="${b}" w:type="dxa"/><w:right w:w="${r}" w:type="dxa"/></w:tcMar>`
            );
          }
          const parts = [];
          for (const kid of cell.kids) {
            if (kid.tag === "grid") {
              parts.push(encodeGrid(kid, ctx));
            } else if (kid.tag === "para") {
              parts.push(encodeParaInner(kid, ctx));
            }
          }
          const lastKid = cell.kids[cell.kids.length - 1];
          if (cell.kids.length === 0 || lastKid?.tag === "grid") {
            parts.push('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
          }
          cellContent = parts.join("");
        } else {
          cellContent = `<w:p><w:pPr/></w:p>`;
        }
        const tcPr = `<w:tcPr>${tcPrParts.join("")}</w:tcPr>`;
        cellXmls.push(`      <w:tc>${tcPr}${cellContent}</w:tc>`);
      }
    }
    const trPrParts = [];
    if (ri === 0 && (gp.headerRow || look?.firstRow)) {
      trPrParts.push("<w:tblHeader/>");
    }
    const originalRow = grid.kids[ri];
    if (originalRow?.heightPt != null && originalRow.heightPt > 0) {
      const hDxa = Math.round(Metric.ptToDxa(originalRow.heightPt));
      trPrParts.push(`<w:trHeight w:val="${hDxa}" w:hRule="atLeast"/>`);
    }
    const trPr = trPrParts.length > 0 ? `<w:trPr>${trPrParts.join("")}</w:trPr>` : "";
    return `    <w:tr>${trPr}
${cellXmls.join("\n")}
    </w:tr>`;
  }).join("\n");
  let tblBorders = "";
  const strokeKindMap = {
    solid: "single",
    dash: "dash",
    dot: "dot",
    double: "double",
    none: "none",
    dotDash: "dotDash",
    dotDotDash: "dotDotDash",
    triple: "triple",
    thinThickSmallGap: "thinThickSmallGap",
    thickThinSmallGap: "thickThinSmallGap",
    thinThickThinSmallGap: "thinThickThinSmallGap"
  };
  if (gp.defaultStroke) {
    const s = gp.defaultStroke;
    const val = strokeKindMap[s.kind] ?? "single";
    if (val === "none" || s.pt <= 0) {
      tblBorders = '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>';
    } else {
      const sz = Math.max(2, Math.round(s.pt * 8));
      const clr = s.color ? s.color.replace("#", "") : "auto";
      const bdr = `w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"`;
      tblBorders = `<w:tblBorders><w:top ${bdr}/><w:left ${bdr}/><w:bottom ${bdr}/><w:right ${bdr}/><w:insideH ${bdr}/><w:insideV ${bdr}/></w:tblBorders>`;
    }
  }
  const tblAlignMap = {
    left: "start",
    center: "center",
    right: "end",
    justify: "start"
  };
  const tblJc = gp.align ? `<w:jc w:val="${tblAlignMap[gp.align] ?? "start"}"/>` : "";
  return `    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${Math.round(totalDxa)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="${firstRow}" w:lastRow="${lastRow}" w:firstColumn="${firstCol}" w:lastColumn="${lastCol}" w:noHBand="${noHBand}" w:noVBand="${noVBand}"/>${tblBorders}${tblJc}<w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="72" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
${rows}
    </w:tbl>`;
}
function encodeCellBorders(cp) {
  if (!cp.top && !cp.bot && !cp.left && !cp.right) return "";
  const strokeKindMap = {
    solid: "single",
    dash: "dash",
    dot: "dot",
    double: "double",
    none: "none",
    dotDash: "dotDash",
    dotDotDash: "dotDotDash",
    triple: "triple"
  };
  const encode = (s, tag) => {
    if (!s || !tag) return "";
    const val = strokeKindMap[s.kind] ?? "single";
    if (val === "none" || s.pt <= 0) {
      return `<w:${tag} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`;
    }
    const sz = Math.max(2, Math.round(s.pt * 8));
    const clr = s.color ? s.color.replace("#", "") : "auto";
    return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"/>`;
  };
  return `<w:tcBorders>${encode(cp.top, "top")}${encode(cp.bot, "bottom")}${encode(cp.left, "left")}${encode(cp.right, "right")}</w:tcBorders>`;
}
function esc2(s) {
  if (!s) return "";
  s = s.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
  s = s.replace(/湰灧/g, "");
  s = s.replace(/\uFEFF/g, "");
  s = s.replace(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]/g, "");
  return TextKit.escapeXml(s);
}
registry.registerEncoder(new DocxEncoder());

// src/encoders/md/MdEncoder.ts
var MdEncoder = class extends BaseEncoder {
  getFormat() {
    return "md";
  }
  async encode(doc, options) {
    const includeImages = options?.includeImages !== false;
    try {
      const warns = [];
      const parts = [];
      for (const sheet of doc.kids) {
        if (sheet.headers && sheet.headers.default && sheet.headers.default.length > 0) warns.push("[SHIELD] MD: \uBA38\uB9AC\uAE00(header) \uD45C\uD604 \uBD88\uAC00 \u2014 \uC190\uC2E4\uB428");
        if (sheet.footers && sheet.footers.default && sheet.footers.default.length > 0) warns.push("[SHIELD] MD: \uBC14\uB2E5\uAE00(footer) \uD45C\uD604 \uBD88\uAC00 \u2014 \uC190\uC2E4\uB428");
        for (const kid of sheet.kids) parts.push(encodeContent2(kid, warns, includeImages));
      }
      return succeed(this.stringToBytes(parts.join("\n\n")), warns);
    } catch (e) {
      return fail(`MD encode error: ${e?.message ?? String(e)}`);
    }
  }
};
function encodeContent2(node, warns, includeImages) {
  return node.tag === "grid" ? encodeGrid2(node, warns, includeImages) : encodePara(node, warns, includeImages);
}
function encodePara(para, warns, includeImages) {
  const text = para.kids.map((k) => {
    if (k.tag === "span") return encodeSpan(k, warns);
    if (k.tag === "img") return encodeImage3(k, includeImages);
    return "";
  }).join("");
  if (para.props.heading) return `${"#".repeat(para.props.heading)} ${text}`;
  if (para.props.listOrd !== void 0) {
    const indent = "  ".repeat(para.props.listLv ?? 0);
    return `${indent}${para.props.listOrd ? "1." : "-"} ${text}`;
  }
  if (para.props.align && para.props.align !== "left" && para.props.align !== "justify") {
    return `<div align="${para.props.align}">${text}</div>`;
  }
  return text;
}
function encodeSpan(span, warns) {
  let hasPageNum = false;
  const textParts = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") textParts.push(kid.content);
    else if (kid.tag === "pagenum") {
      hasPageNum = true;
      warns.push("[SHIELD] MD: \uD398\uC774\uC9C0 \uBC88\uD638 \uD45C\uD604 \uBD88\uAC00 \u2014 \uC190\uC2E4\uB428");
    }
  }
  let r = textParts.join("");
  if (hasPageNum && r === "") r = "[\uD398\uC774\uC9C0 \uBC88\uD638]";
  const cssStyles = [];
  if (span.props.font) cssStyles.push(`font-family: ${span.props.font}`);
  if (span.props.pt) cssStyles.push(`font-size: ${span.props.pt}pt`);
  if (span.props.color) cssStyles.push(`color: #${span.props.color}`);
  if (span.props.bg) cssStyles.push(`background-color: #${span.props.bg}`);
  const hasHtmlStyle = cssStyles.length > 0;
  if (hasHtmlStyle) {
    if (span.props.b) cssStyles.push("font-weight: bold");
    if (span.props.i) cssStyles.push("font-style: italic");
    if (span.props.s) cssStyles.push("text-decoration: line-through");
    if (span.props.u) {
      const existing = cssStyles.find((s) => s.startsWith("text-decoration:"));
      if (existing) {
        const idx = cssStyles.indexOf(existing);
        cssStyles[idx] = existing.replace("line-through", "underline line-through");
        if (!existing.includes("line-through")) cssStyles[idx] = existing + " underline";
      } else {
        cssStyles.push("text-decoration: underline");
      }
    }
    const styleAttr = cssStyles.join("; ");
    if (span.props.sup) return `<sup style="${styleAttr}">${r}</sup>`;
    if (span.props.sub) return `<sub style="${styleAttr}">${r}</sub>`;
    return `<span style="${styleAttr}">${r}</span>`;
  }
  if (span.props.b && span.props.i) r = `***${r}***`;
  else if (span.props.b) r = `**${r}**`;
  else if (span.props.i) r = `*${r}*`;
  if (span.props.s) r = `~~${r}~~`;
  if (span.props.u) r = `<u>${r}</u>`;
  if (span.props.sup) r = `<sup>${r}</sup>`;
  if (span.props.sub) r = `<sub>${r}</sub>`;
  return r;
}
function encodeImage3(img, includeImages) {
  if (!includeImages) {
    return `![${img.alt ?? ""}]`;
  }
  return `![${img.alt ?? ""}](data:${img.mime};base64,${img.b64})`;
}
function strokeToCss(s) {
  if (!s || s.kind === "none" || s.pt <= 0) return void 0;
  const kindMap = { solid: "solid", dash: "dashed", dot: "dotted", double: "double", none: "none" };
  const style = kindMap[s.kind] ?? "solid";
  const px = Math.max(1, Math.round(s.pt * 96 / 72));
  const color = s.color.startsWith("#") ? s.color : `#${s.color}`;
  return `${px}px ${style} ${color}`;
}
function encodeGrid2(grid, warns, includeImages) {
  if (grid.kids.length === 0) return "";
  const rowCount = grid.kids.length;
  const occupancy = Array.from({ length: rowCount }, () => /* @__PURE__ */ new Set());
  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let ci = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(ci)) ci++;
      if (cell.rs > 1) {
        for (let r = ri + 1; r < ri + cell.rs && r < rowCount; r++) {
          for (let c = ci; c < ci + cell.cs; c++) occupancy[r].add(c);
        }
      }
      ci += cell.cs;
    }
    while (occupancy[ri].has(ci)) ci++;
    if (ci > colCount) colCount = ci;
  }
  let rows = "";
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let cells = "";
    let colIdx = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(colIdx)) colIdx++;
      const cs = cell.cs > 1 ? ` colspan="${cell.cs}"` : "";
      const rs = cell.rs > 1 ? ` rowspan="${cell.rs}"` : "";
      const styles = ["padding:4px 6px", "vertical-align:top"];
      const top = strokeToCss(cell.props.top);
      const bot = strokeToCss(cell.props.bot);
      const left = strokeToCss(cell.props.left);
      const right = strokeToCss(cell.props.right);
      if (top) styles.push(`border-top:${top}`);
      if (bot) styles.push(`border-bottom:${bot}`);
      if (left) styles.push(`border-left:${left}`);
      if (right) styles.push(`border-right:${right}`);
      if (cell.props.bg) styles.push(`background-color:#${cell.props.bg}`);
      if (cell.props.va === "mid") styles[1] = "vertical-align:middle";
      else if (cell.props.va === "bot") styles[1] = "vertical-align:bottom";
      const tag = grid.props.headerRow && ri === 0 || cell.props.isHeader ? "th" : "td";
      const content = cell.kids.map((p) => p.tag === "para" ? encodePara(p, warns, includeImages) : encodeGrid2(p, warns, includeImages)).join("\n");
      cells += `<${tag}${cs}${rs} style="${styles.join(";")}">${content}</${tag}>`;
      colIdx += cell.cs;
    }
    rows += `<tr>${cells}</tr>
`;
  }
  return `<table style="border-collapse:collapse;width:100%">
<tbody>
${rows}</tbody>
</table>
`;
}
registry.registerEncoder(new MdEncoder());

// src/encoders/html/HtmlEncoder.ts
var HtmlEncoder = class extends BaseEncoder {
  getFormat() {
    return "html";
  }
  async encode(doc) {
    try {
      const warns = [];
      const bodyParts = [];
      for (const sheet of doc.kids) {
        if (sheet.headers?.default && sheet.headers.default.length > 0) {
          const hText = sheet.headers.default.map((p) => encodePara2(p, warns)).join("");
          bodyParts.push(`<div class="hwp-header">${hText}</div>`);
        }
        for (const kid of sheet.kids) {
          bodyParts.push(encodeContent3(kid, warns));
        }
        if (sheet.footers?.default && sheet.footers.default.length > 0) {
          const fText = sheet.footers.default.map((p) => encodePara2(p, warns)).join("");
          bodyParts.push(`<div class="hwp-footer">${fText}</div>`);
        }
      }
      const title = this.escapeXml(doc.meta?.title ?? "");
      const html = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${title}</title>
<style>
${BASE_CSS}
</style>
</head>
<body>
<div class="hwp-doc">
${bodyParts.join("\n")}
</div>
</body>
</html>`;
      return succeed(this.stringToBytes(html), warns);
    } catch (e) {
      return fail(`HTML encode error: ${e?.message ?? String(e)}`);
    }
  }
};
var BASE_CSS = `
body { margin: 0; padding: 0; background: #f0f0f0; }
.hwp-doc { max-width: 800px; margin: 0 auto; background: #fff; padding: 40px 60px; box-shadow: 0 0 8px rgba(0,0,0,0.15); }
.hwp-header, .hwp-footer { color: #666; font-size: 0.9em; border-bottom: 1px solid #ddd; margin-bottom: 8px; padding-bottom: 4px; }
.hwp-footer { border-top: 1px solid #ddd; border-bottom: none; margin-top: 8px; padding-top: 4px; }
p { margin: 0; padding: 0; line-height: 1; }
table { border-collapse: collapse; width: 100%; margin: 8px 0; }
td, th { border: 1px solid #ccc; padding: 4px 8px; vertical-align: top; }
img { max-width: 100%; height: auto; }
`.trim();
function encodeContent3(node, warns) {
  return node.tag === "grid" ? encodeGrid3(node, warns) : encodePara2(node, warns);
}
function encodePara2(para, warns) {
  const kids = para.kids.map((k) => {
    if (k.tag === "span") return encodeSpan2(k, warns);
    if (k.tag === "img") return encodeImage4(k);
    if (k.tag === "link") {
      const link = k;
      const inner = link.kids.map((s) => encodeSpan2(s, warns)).join("");
      return `<a href="${TextKit.escapeXml(link.href)}">${inner}</a>`;
    }
    return "";
  }).join("");
  if (para.props.heading) {
    const tag = `h${para.props.heading}`;
    return `<${tag}>${kids}</${tag}>
`;
  }
  if (para.props.listOrd !== void 0) {
    const indent = (para.props.listLv ?? 0) * 20;
    const style = indent > 0 ? ` style="margin-left:${indent}px"` : "";
    const marker = para.props.listOrd ? `<span class="list-marker">1. </span>` : `<span class="list-marker">\u2022 </span>`;
    return `<p${style}>${marker}${kids}</p>
`;
  }
  const align = para.props.align;
  const styleAttrs = [];
  if (align && align !== "left") styleAttrs.push(`text-align:${align}`);
  if (para.props.indentPt) styleAttrs.push(`margin-left:${para.props.indentPt.toFixed(1)}pt`);
  if (para.props.spaceBefore) styleAttrs.push(`margin-top:${para.props.spaceBefore.toFixed(1)}pt`);
  if (para.props.spaceAfter) styleAttrs.push(`margin-bottom:${para.props.spaceAfter.toFixed(1)}pt`);
  if (para.props.lineHeight) styleAttrs.push(`line-height:${para.props.lineHeight}`);
  const styleAttr = styleAttrs.length > 0 ? ` style="${styleAttrs.join(";")}"` : "";
  return `<p${styleAttr}>${kids || "&nbsp;"}</p>
`;
}
function encodeSpan2(span, warns) {
  const parts = [];
  let hasPageNum = false;
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const content = kid.content.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      if (content) parts.push(TextKit.escapeXml(content));
    } else if (kid.tag === "br") {
      parts.push("<br>");
    } else if (kid.tag === "pb") {
      parts.push('<div style="page-break-after:always"></div>');
    } else if (kid.tag === "pagenum") {
      hasPageNum = true;
      warns.push("[SHIELD] HTML: \uD398\uC774\uC9C0 \uBC88\uD638 \u2014 \uC815\uC801 \uAC12\uC73C\uB85C \uB300\uCCB4\uB428");
      parts.push('<span class="page-num">[\uD398\uC774\uC9C0]</span>');
    }
  }
  let text = parts.join("");
  if (hasPageNum && text.trim() === '<span class="page-num">[\uD398\uC774\uC9C0]</span>') {
  }
  const p = span.props;
  const css = [];
  if (p.font) css.push(`font-family:${TextKit.escapeXml(p.font)}`);
  if (p.pt) css.push(`font-size:${p.pt}pt`);
  if (p.color) css.push(`color:#${p.color}`);
  if (p.bg) css.push(`background-color:#${p.bg}`);
  if (p.b) css.push("font-weight:bold");
  if (p.i) css.push("font-style:italic");
  const decorations = [];
  if (p.u) decorations.push("underline");
  if (p.s) decorations.push("line-through");
  if (decorations.length > 0) css.push(`text-decoration:${decorations.join(" ")}`);
  if (p.sup) return `<sup${css.length ? ` style="${css.join(";")}"` : ""}>${text}</sup>`;
  if (p.sub) return `<sub${css.length ? ` style="${css.join(";")}"` : ""}>${text}</sub>`;
  if (css.length > 0) return `<span style="${css.join(";")}">${text}</span>`;
  return text;
}
function encodeImage4(img) {
  const wStyle = img.w ? ` width="${Math.round(img.w / 72 * 96)}px"` : "";
  const hStyle = img.h ? ` height="${Math.round(img.h / 72 * 96)}px"` : "";
  const alt = TextKit.escapeXml(img.alt ?? "");
  return `<img src="data:${img.mime};base64,${img.b64}" alt="${alt}"${wStyle}${hStyle}>`;
}
function encodeGrid3(grid, warns) {
  if (grid.kids.length === 0) return "";
  const rowCount = grid.kids.length;
  const occupancy = Array.from({ length: rowCount }, () => /* @__PURE__ */ new Set());
  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let ci = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(ci)) ci++;
      if (cell.rs > 1) {
        for (let r = ri + 1; r < ri + cell.rs && r < rowCount; r++) {
          for (let c = ci; c < ci + cell.cs; c++) occupancy[r].add(c);
        }
      }
      ci += cell.cs;
    }
    while (occupancy[ri].has(ci)) ci++;
    if (ci > colCount) colCount = ci;
  }
  let rows = "";
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let cells = "";
    let ci = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(ci)) ci++;
      const isHeader = cell.props.isHeader || grid.props.headerRow && ri === 0;
      const tag = isHeader ? "th" : "td";
      const cs = cell.cs > 1 ? ` colspan="${cell.cs}"` : "";
      const rs = cell.rs > 1 ? ` rowspan="${cell.rs}"` : "";
      const styleAttrs = [];
      if (cell.props.bg) styleAttrs.push(`background-color:#${cell.props.bg}`);
      const va = cell.props.va;
      if (va === "mid") styleAttrs.push("vertical-align:middle");
      else if (va === "bot") styleAttrs.push("vertical-align:bottom");
      const styleAttr = styleAttrs.length > 0 ? ` style="${styleAttrs.join(";")}"` : "";
      const content = cell.kids.map((p) => p.tag === "para" ? encodePara2(p, warns) : encodeGrid3(p, warns)).join("");
      cells += `<${tag}${cs}${rs}${styleAttr}>${content}</${tag}>`;
      ci += cell.cs;
    }
    rows += `<tr>${cells}</tr>
`;
  }
  return `<table>
<tbody>
${rows}</tbody>
</table>
`;
}
registry.registerEncoder(new HtmlEncoder());

// src/encoders/hwp/HwpEncoder.ts
import pako3 from "pako";
var T = 16;
var TAG_DOCUMENT_PROPERTIES = T + 0;
var TAG_ID_MAPPINGS = T + 1;
var TAG_BIN_DATA = T + 2;
var TAG_FACE_NAME2 = T + 3;
var TAG_BORDER_FILL2 = T + 4;
var TAG_CHAR_SHAPE2 = T + 5;
var TAG_PARA_SHAPE2 = T + 9;
var TAG_STYLE = T + 10;
var TAG_PARA_HEADER2 = T + 50;
var TAG_PARA_TEXT2 = T + 51;
var TAG_PARA_CHAR_SHAPE2 = T + 52;
var TAG_PARA_LINE_SEG = T + 53;
var TAG_CTRL_HEADER2 = T + 55;
var TAG_LIST_HEADER2 = T + 56;
var TAG_PAGE_DEF2 = T + 57;
var TAG_FOOTNOTE_SHAPE = T + 58;
var TAG_TABLE = T + 61;
var TAG_SHAPE_COMPONENT_PICTURE = T + 69;
var CTRL_TABLE2 = 1952607264;
var CTRL_SECD = 1936024420;
var CTRL_PIC = 611346787;
var CTRL_FIELD_BEGIN = 1684825637;
var CTRL_FIELD_END = 1684825692;
var BORDER_W_PT2 = [
  0.28,
  0.34,
  0.43,
  0.57,
  0.71,
  0.85,
  1.13,
  1.42,
  1.7,
  1.98,
  2.84,
  4.25,
  5.67,
  8.5,
  11.34,
  14.17
];
var BORDER_KIND_IDX = {
  solid: 0,
  dot: 1,
  dash: 2,
  double: 7,
  triple: 8,
  none: 0
};
var ALIGN_CODE = {
  justify: 0,
  left: 1,
  right: 2,
  center: 3,
  distribute: 4
};
var BufWriter = class {
  constructor() {
    this.chunks = [];
    this._sz = 0;
  }
  get size() {
    return this._sz;
  }
  u8(v) {
    this.chunks.push(new Uint8Array([v & 255]));
    this._sz++;
    return this;
  }
  u16(v) {
    this.chunks.push(new Uint8Array([v & 255, v >> 8 & 255]));
    this._sz += 2;
    return this;
  }
  u32(v) {
    const b = new Uint8Array(4);
    b[0] = v & 255;
    b[1] = v >>> 8 & 255;
    b[2] = v >>> 16 & 255;
    b[3] = v >>> 24 & 255;
    this.chunks.push(b);
    this._sz += 4;
    return this;
  }
  i32(v) {
    return this.u32(v < 0 ? v + 4294967296 : v);
  }
  i16(v) {
    return this.u16(v < 0 ? v + 65536 : v);
  }
  bytes(d) {
    this.chunks.push(d);
    this._sz += d.length;
    return this;
  }
  zeros(n) {
    this.chunks.push(new Uint8Array(n));
    this._sz += n;
    return this;
  }
  utf16(s) {
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i));
    return this;
  }
  colorRef(hex) {
    const h = (hex || "000000").replace("#", "").padStart(6, "0");
    return this.u8(parseInt(h.slice(0, 2), 16)).u8(parseInt(h.slice(2, 4), 16)).u8(parseInt(h.slice(4, 6), 16)).u8(0);
  }
  build() {
    const out = new Uint8Array(this._sz);
    let off = 0;
    for (const c of this.chunks) {
      out.set(c, off);
      off += c.length;
    }
    return out;
  }
};
function mkRec(tag, level, data) {
  const sz = data.length;
  const enc = Math.min(sz, 4095);
  const hdr = enc << 20 | (level & 1023) << 10 | tag & 1023;
  const w = new BufWriter().u32(hdr);
  if (enc >= 4095) w.u32(sz);
  w.bytes(data);
  return w.build();
}
function readPixelDims2(data, mime) {
  try {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    if (mime.includes("png")) {
      if (data.length >= 24 && view.getUint32(0) === 2303741511 && view.getUint32(4) === 218765834) {
        return { w: view.getUint32(16), h: view.getUint32(20) };
      }
    } else if (mime.includes("jpeg") || mime.includes("jpg")) {
      let off = 2;
      while (off < data.length - 4) {
        const marker = view.getUint16(off);
        off += 2;
        if (marker === 65472 || marker === 65474) {
          return { w: view.getUint16(off + 5), h: view.getUint16(off + 3) };
        }
        if ((marker & 65280) !== 65280) break;
        off += view.getUint16(off);
      }
    }
  } catch {
  }
  return null;
}
var LANG_GROUPS2 = [
  "HANGUL",
  "LATIN",
  "HANJA",
  "JAPANESE",
  "OTHER",
  "SYMBOL",
  "USER"
];
function isKoreanFont(face) {
  return /[\uAC00-\uD7A3\u3131-\u318E]/.test(face) || ["\uB9D1\uC740", "\uB098\uB214", "\uAD74\uB9BC", "\uB3CB\uC6C0", "\uBC14\uD0D5", "\uD568\uCD08\uB86C", "\uD55C\uCEF4", "HY"].some(
    (k) => face.includes(k)
  );
}
var HwpStyleBank = class {
  // id=0 → 모두 0
  constructor() {
    this.DEF_STROKE = { kind: "solid", pt: 0.5, color: "000000" };
    // 언어별 독립 폰트 목록 (ANYTOHWP langFontFaces)
    this.langFonts = new Map(
      LANG_GROUPS2.map((g) => [g, []])
    );
    this.langFontIdx = new Map(
      LANG_GROUPS2.map((g) => [g, /* @__PURE__ */ new Map()])
    );
    // charShape, parShape, borderFill 레지스트리
    this.csProps = [{}];
    this.csIdx = /* @__PURE__ */ new Map([[csKey({}), 0]]);
    this.psProps = [{}];
    this.psIdx = /* @__PURE__ */ new Map([[psKey({}), 0]]);
    this.bfData = [];
    this.bfIdx = /* @__PURE__ */ new Map();
    // charShape마다 언어별 fontId를 기록
    this.csFontIds = [[0, 0, 0, 0, 0, 0, 0]];
    for (const g of LANG_GROUPS2) this._registerLangFont(g, "\uD568\uCD08\uB86C\uBC14\uD0D5");
    this.addBorderFill(this.DEF_STROKE);
  }
  _registerLangFont(lang, face) {
    const idx = this.langFontIdx.get(lang);
    if (idx.has(face)) return idx.get(face);
    const id = this.langFonts.get(lang).length;
    this.langFonts.get(lang).push(face);
    idx.set(face, id);
    return id;
  }
  /** 폰트 이름 → 언어별 7개 ID 반환 (ANYTOHWP 방식) */
  registerFontForLangs(rawFace) {
    const face = safeFontToKr(rawFace) || "\uD568\uCD08\uB86C\uBC14\uD0D5";
    const isKor = isKoreanFont(face);
    const hangulFace = isKor ? face : "\uD568\uCD08\uB86C\uBC14\uD0D5";
    const latinFace = isKor ? "\uD568\uCD08\uB86C\uBC14\uD0D5" : face;
    const ids = [];
    for (const lang of LANG_GROUPS2) {
      const f = lang === "LATIN" ? latinFace : hangulFace;
      ids.push(this._registerLangFont(lang, f));
    }
    return ids;
  }
  /** 언어별 폰트 목록 반환 */
  getFontsForLang(lang) {
    return [...this.langFonts.get(lang) ?? []];
  }
  /** 폰트 수 반환 (mkIdMappings용) */
  getFontCount(lang) {
    return this.langFonts.get(lang)?.length ?? 0;
  }
  addCharShape(p) {
    const k = csKey(p);
    if (this.csIdx.has(k)) return this.csIdx.get(k);
    const id = this.csProps.length;
    const fIds = p.font ? this.registerFontForLangs(p.font) : [0, 0, 0, 0, 0, 0, 0];
    this.csProps.push(p);
    this.csFontIds.push(fIds);
    this.csIdx.set(k, id);
    return id;
  }
  addParaShape(p) {
    const k = psKey(p);
    if (this.psIdx.has(k)) return this.psIdx.get(k);
    const id = this.psProps.length;
    this.psProps.push(p);
    this.psIdx.set(k, id);
    return id;
  }
  addBorderFill(s, bg) {
    const k = bfKey(s, bg);
    if (this.bfIdx.has(k)) return this.bfIdx.get(k);
    const id = this.bfData.length + 1;
    this.bfData.push({ uniform: true, s, bg });
    this.bfIdx.set(k, id);
    return id;
  }
  addBorderFillPerSide(l, r, t, b, bg) {
    const k = bfPerSideKey(l, r, t, b, bg);
    if (this.bfIdx.has(k)) return this.bfIdx.get(k);
    const id = this.bfData.length + 1;
    this.bfData.push({ uniform: false, l, r, t, b, bg });
    this.bfIdx.set(k, id);
    return id;
  }
};
function csKey(p) {
  return [
    p.font ?? "",
    p.pt ?? 10,
    p.b ? 1 : 0,
    p.i ? 1 : 0,
    p.u ? 1 : 0,
    p.s ? 1 : 0,
    p.sup ? 1 : 0,
    p.sub ? 1 : 0,
    p.color ?? "000000"
  ].join("|");
}
function psKey(p) {
  return [
    p.align ?? "left",
    p.indentPt ?? 0,
    p.firstLineIndentPt ?? 0,
    p.spaceBefore ?? 0,
    p.spaceAfter ?? 0,
    p.lineHeight ?? 1
  ].join("|");
}
function bfKey(s, bg) {
  return `${s.kind}|${s.pt}|${s.color}|${bg ?? ""}`;
}
function bfPerSideKey(l, r, t, b, bg) {
  return `${bfKey(l)}/${bfKey(r)}/${bfKey(t)}/${bfKey(b)}/${bg ?? ""}`;
}
function collectNode(node, bank) {
  if (node.tag === "para") {
    bank.addParaShape(node.props);
    for (const kid of node.kids) {
      if (kid.tag === "span") bank.addCharShape(kid.props);
    }
  } else if (node.tag === "grid") {
    if (node.props.defaultStroke) bank.addBorderFill(node.props.defaultStroke);
    for (const row of node.kids) {
      for (const cell of row.kids) {
        const defStroke = node.props.defaultStroke ?? bank.DEF_STROKE;
        const cp = cell.props;
        if (cp.top || cp.bot || cp.left || cp.right) {
          bank.addBorderFillPerSide(
            cp.left ?? defStroke,
            cp.right ?? defStroke,
            cp.top ?? defStroke,
            cp.bot ?? defStroke,
            cp.bg
          );
        } else {
          bank.addBorderFill(defStroke, cp.bg);
        }
        for (const para of cell.kids) collectNode(para, bank);
      }
    }
  }
}
function mkDocumentProperties() {
  return new BufWriter().u16(1).u16(1).u16(1).u16(1).u16(1).u16(1).u16(1).u32(0).u32(0).u32(0).build();
}
function mkIdMappings(bank, nBinData = 0) {
  const w = new BufWriter();
  w.u32(nBinData);
  for (const lang of LANG_GROUPS2) w.u32(bank.getFontCount(lang));
  w.u32(bank.bfData.length);
  w.u32(bank.csProps.length);
  w.u32(0);
  w.u32(0);
  w.u32(0);
  w.u32(bank.psProps.length);
  w.u32(1);
  w.u32(0);
  w.u32(0);
  w.u32(0);
  return w.build();
}
function mkStyle(name, engName, paraPrId, charPrId) {
  return new BufWriter().u16(name.length).utf16(name).u16(engName.length).utf16(engName).u16(paraPrId).u16(charPrId).u16(0).u16(1042).u16(0).build();
}
function mkFaceName(name) {
  return new BufWriter().u8(0).u16(name.length).utf16(name).u8(0).u16(0).zeros(10).u16(0).build();
}
function borderWidthIdx(pt) {
  let best = 0;
  for (let i = 0; i < BORDER_W_PT2.length; i++) {
    if (Math.abs(BORDER_W_PT2[i] - pt) < Math.abs(BORDER_W_PT2[best] - pt))
      best = i;
  }
  return best;
}
function mkBorderFill(s, bg) {
  const w = new BufWriter();
  const t = BORDER_KIND_IDX[s.kind] ?? 0;
  const wi = borderWidthIdx(s.pt);
  const col = s.color || "000000";
  w.u16(0);
  for (let i = 0; i < 4; i++) w.u8(t);
  for (let i = 0; i < 4; i++) w.u8(wi);
  for (let i = 0; i < 4; i++) w.colorRef(col);
  w.u8(0).u8(0).colorRef("000000");
  if (bg) {
    w.u32(1).colorRef(bg).colorRef("FFFFFF").u32(0);
  } else {
    w.u32(0);
  }
  return w.build();
}
function mkBorderFillPerSide(l, r, t, b, bg) {
  const w = new BufWriter();
  w.u16(0);
  w.u8(BORDER_KIND_IDX[l.kind] ?? 0).u8(BORDER_KIND_IDX[r.kind] ?? 0).u8(BORDER_KIND_IDX[t.kind] ?? 0).u8(BORDER_KIND_IDX[b.kind] ?? 0);
  w.u8(borderWidthIdx(l.pt)).u8(borderWidthIdx(r.pt)).u8(borderWidthIdx(t.pt)).u8(borderWidthIdx(b.pt));
  w.colorRef(l.color || "000000").colorRef(r.color || "000000").colorRef(t.color || "000000").colorRef(b.color || "000000");
  w.u8(0).u8(0).colorRef("000000");
  if (bg) {
    w.u32(1).colorRef(bg).colorRef("FFFFFF").u32(0);
  } else {
    w.u32(0);
  }
  return w.build();
}
function mkCharShape(fontIds, p) {
  const height = Math.round((p.pt ?? 10) * 100);
  let attr = 0;
  if (p.i) attr |= 1 << 0;
  if (p.b) attr |= 1 << 1;
  if (p.u) attr |= 1 << 2;
  if (p.s) attr |= 1 << 18;
  if (p.sup) attr |= 1 << 16;
  if (p.sub) attr |= 2 << 16;
  const w = new BufWriter();
  for (const id of fontIds) w.u16(id);
  for (let i = 0; i < 7; i++) w.u8(100);
  for (let i = 0; i < 7; i++) w.u8(0);
  for (let i = 0; i < 7; i++) w.u8(100);
  for (let i = 0; i < 7; i++) w.u8(0);
  w.i32(height).u32(attr).u8(0).u8(0);
  w.colorRef(p.color ?? "000000");
  w.colorRef("000000");
  w.colorRef(p.bg ?? "FFFFFF");
  w.colorRef("000000");
  w.u16(0);
  w.colorRef("000000");
  return w.build();
}
function mkParaShape(p) {
  const alignVal = ALIGN_CODE[p.align ?? "left"] ?? 1;
  const attr1 = (alignVal & 7) << 2;
  const lineSpacePct = p.lineHeight ? Math.round(p.lineHeight * 100) : 160;
  return new BufWriter().u32(attr1).i32(Metric.ptToHwp(p.indentPt ?? 0)).i32(Metric.ptToHwp(p.indentRightPt ?? 0)).i32(Metric.ptToHwp(p.firstLineIndentPt ?? 0)).i32(Metric.ptToHwp(p.spaceBefore ?? 0)).i32(Metric.ptToHwp(p.spaceAfter ?? 0)).i32(lineSpacePct).u16(0).u16(0).u16(0).i16(0).i16(0).i16(0).i16(0).u32(0).u32(4).u32(lineSpacePct).build();
}
function mkBinData(id, ext) {
  return new BufWriter().u16(2).u16(id).u16(ext.length).utf16(ext).build();
}
function buildDocInfoStream(bank, images = []) {
  const chunks = [];
  chunks.push(mkRec(TAG_DOCUMENT_PROPERTIES, 0, mkDocumentProperties()));
  chunks.push(mkRec(TAG_ID_MAPPINGS, 1, mkIdMappings(bank, images.length)));
  for (const img of images) {
    chunks.push(mkRec(TAG_BIN_DATA, 1, mkBinData(img.id, img.ext)));
  }
  for (const lang of LANG_GROUPS2) {
    for (const face of bank.getFontsForLang(lang)) {
      chunks.push(mkRec(TAG_FACE_NAME2, 1, mkFaceName(face)));
    }
  }
  for (const entry of bank.bfData) {
    chunks.push(
      mkRec(
        TAG_BORDER_FILL2,
        1,
        entry.uniform ? mkBorderFill(entry.s, entry.bg) : mkBorderFillPerSide(entry.l, entry.r, entry.t, entry.b, entry.bg)
      )
    );
  }
  for (let i = 0; i < bank.csProps.length; i++) {
    chunks.push(
      mkRec(TAG_CHAR_SHAPE2, 1, mkCharShape(bank.csFontIds[i], bank.csProps[i]))
    );
  }
  for (const p of bank.psProps) {
    chunks.push(mkRec(TAG_PARA_SHAPE2, 1, mkParaShape(p)));
  }
  chunks.push(mkRec(TAG_STYLE, 1, mkStyle("\uBC14\uD0D5\uAE00", "Normal", 0, 0)));
  return concatU8(chunks);
}
function mkPageDef(dims) {
  return new BufWriter().u32(Metric.ptToHwp(dims.wPt)).u32(Metric.ptToHwp(dims.hPt)).u32(Metric.ptToHwp(dims.ml)).u32(Metric.ptToHwp(dims.mr)).u32(Metric.ptToHwp(dims.mt)).u32(Metric.ptToHwp(dims.mb)).zeros(12).u32(dims.orient === "landscape" ? 1 : 0).build();
}
function mkParaHeader(nchars, ctrlMask, psId, csCount, lineAlignCount = 0, instanceId = 0) {
  return new BufWriter().u32(nchars).u32(ctrlMask).u16(psId).u8(0).u8(0).u16(csCount).u16(0).u16(lineAlignCount).u32(instanceId).u16(0).build();
}
function mkParaText(text) {
  const w = new BufWriter();
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    w.u16(c);
  }
  w.u16(13);
  return w.build();
}
function mkParaCharShape(pairs) {
  const w = new BufWriter();
  for (const [pos, id] of pairs) w.u32(pos).u32(id);
  return w.build();
}
function calcLineHeight(type, value, textHeight) {
  switch (type) {
    case 0:
      return Math.floor(textHeight * value / 100);
    case 1:
      return value;
    case 2:
      return Math.max(textHeight, value);
    case 3:
      return textHeight + value;
    case 4:
      return Math.floor(textHeight * value);
    default:
      return Math.floor(textHeight * value / 100);
  }
}
function mkLineSeg(textStartPos, vertPos, vertSize, textHeight, baseline, spacing, horzPos, horzSize, flags) {
  return new BufWriter().u32(textStartPos).i32(vertPos).i32(vertSize).i32(textHeight).i32(baseline).i32(spacing).i32(horzPos).i32(horzSize).u32(flags).build();
}
function buildDefaultLineSeg(availWidthHwp, fontHwp, nchars, paraProps, vertPos = 0) {
  const ratio = paraProps?.lineHeight ? Math.round(paraProps.lineHeight * 100) : 160;
  const vertSize = calcLineHeight(0, ratio, fontHwp);
  const baseline = Math.round(fontHwp * 0.85);
  const spacing = vertSize - fontHwp;
  const flags = 3;
  return mkLineSeg(
    0,
    vertPos,
    vertSize,
    fontHwp,
    baseline,
    spacing,
    0,
    availWidthHwp,
    flags
  );
}
function mkSecdParaText() {
  const lo = CTRL_SECD & 65535;
  const hi = CTRL_SECD >>> 16 & 65535;
  return new BufWriter().u16(2).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(2).u16(13).build();
}
function mkTableParaText() {
  const lo = CTRL_TABLE2 & 65535;
  const hi = CTRL_TABLE2 >>> 16 & 65535;
  return new BufWriter().u16(11).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(11).u16(13).build();
}
function mkPicParaText() {
  const lo = CTRL_PIC & 65535;
  const hi = CTRL_PIC >>> 16 & 65535;
  return new BufWriter().u16(11).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(11).u16(13).build();
}
function mkShapeComponentPicture(binDataId, wHwp, hHwp) {
  const w = new BufWriter();
  w.u32(CTRL_PIC).zeros(15);
  w.u32(0).u32(0).u32(wHwp).u32(hHwp);
  w.u32(0).u32(0).u32(wHwp).u32(hHwp);
  w.zeros(36);
  w.u16(binDataId).u8(0).u8(0).u8(0).zeros(5);
  return w.build();
}
function mkObjectCtrl(ctrlId, wHwp, hHwp, instanceId, layout) {
  let attr = 136978960;
  if (layout?.wrap === "inline") attr |= 1 << 3;
  return new BufWriter().u32(ctrlId).u32(attr).i32(layout?.yPt ? Metric.ptToHwp(layout.yPt) : 0).i32(layout?.xPt ? Metric.ptToHwp(layout.xPt) : 0).u32(wHwp).u32(hHwp).i32(layout?.zOrder ?? 0).u16(layout?.distL ? Metric.ptToHwp(layout.distL) : 0).u16(layout?.distR ? Metric.ptToHwp(layout.distR) : 0).u16(layout?.distT ? Metric.ptToHwp(layout.distT) : 0).u16(layout?.distB ? Metric.ptToHwp(layout.distB) : 0).u32(instanceId).i32(0).u16(0).build();
}
function mkFieldBeginCtrl(instanceId) {
  return new BufWriter().u32(CTRL_FIELD_BEGIN).u32(2).zeros(28).u32(instanceId).zeros(6).build();
}
function mkFieldEndCtrl(beginId) {
  return new BufWriter().u32(CTRL_FIELD_END).u32(0).zeros(28).u32(beginId).zeros(6).build();
}
function encodePicPara(imgNode, binDataId, bank, lv, idGen, availWidthHwp) {
  const rawData = TextKit.base64Decode(imgNode.b64);
  const pixDims = readPixelDims2(rawData, imgNode.mime);
  let wHwp, hHwp;
  if (pixDims && pixDims.w > 0 && pixDims.h > 0) {
    wHwp = Metric.ptToHwp(pixDims.w * 72 / 96);
    hHwp = Metric.ptToHwp(pixDims.h * 72 / 96);
  } else {
    wHwp = Metric.ptToHwp(imgNode.w);
    hHwp = Metric.ptToHwp(imgNode.h);
  }
  if (wHwp > availWidthHwp) {
    hHwp = Math.round(hHwp * availWidthHwp / wHwp);
    wHwp = availWidthHwp;
  }
  const CTRL_MASK = 1 << 11;
  const instanceId = idGen();
  const psId = bank.addParaShape({});
  return [
    mkRec(
      TAG_PARA_HEADER2,
      lv,
      mkParaHeader(9, CTRL_MASK, psId, 1, 1, instanceId)
    ),
    mkRec(TAG_PARA_TEXT2, lv + 1, mkPicParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE2, lv + 1, mkParaCharShape([[0, 0]])),
    mkRec(
      TAG_PARA_LINE_SEG,
      lv + 1,
      buildDefaultLineSeg(availWidthHwp, hHwp, 9)
    ),
    mkRec(
      TAG_CTRL_HEADER2,
      lv + 1,
      mkObjectCtrl(CTRL_PIC, wHwp, hHwp, idGen(), imgNode.layout)
    ),
    mkRec(
      TAG_SHAPE_COMPONENT_PICTURE,
      lv + 2,
      mkShapeComponentPicture(binDataId, wHwp, hHwp)
    )
  ];
}
function encodePara3(para, bank, lv, instanceId, availWidthHwp, mask = 0, vertPos = 0) {
  let text = "";
  const csPairs = [];
  let pos = 0;
  let fontHwp = 1e3;
  const ctrlRecords = [];
  for (const kid of para.kids) {
    if (kid.tag === "span" && kid.props.pt && kid.props.pt > 0) {
      fontHwp = Metric.ptToHwp(kid.props.pt);
      break;
    }
  }
  let localIdCounter = 1e4;
  const localIdGen = () => localIdCounter++;
  function processKids(kids) {
    for (const kid of kids) {
      if (kid.tag === "span") {
        const span = kid;
        const csId = bank.addCharShape(span.props);
        if (!csPairs.length || csPairs[csPairs.length - 1][1] !== csId) {
          csPairs.push([pos, csId]);
        }
        for (const t of span.kids) {
          if (t.tag === "txt") {
            text += t.content;
            pos += t.content.length;
          }
        }
      } else if (kid.tag === "link") {
        const link = kid;
        mask |= 1 << 11;
        const fieldBeginId = localIdGen();
        text += String.fromCharCode(3);
        pos += 1;
        ctrlRecords.push(
          mkRec(TAG_CTRL_HEADER2, lv + 1, mkFieldBeginCtrl(fieldBeginId))
        );
        processKids(link.kids);
        text += String.fromCharCode(4);
        pos += 1;
        ctrlRecords.push(
          mkRec(TAG_CTRL_HEADER2, lv + 1, mkFieldEndCtrl(fieldBeginId))
        );
      }
    }
  }
  processKids(para.kids);
  if (!csPairs.length) csPairs.push([0, 0]);
  const psId = bank.addParaShape(para.props);
  const nchars = text.length + 1;
  return [
    mkRec(
      TAG_PARA_HEADER2,
      lv,
      mkParaHeader(nchars, mask, psId, csPairs.length, 1, instanceId)
    ),
    mkRec(TAG_PARA_TEXT2, lv + 1, mkParaText(text)),
    mkRec(TAG_PARA_CHAR_SHAPE2, lv + 1, mkParaCharShape(csPairs)),
    mkRec(
      TAG_PARA_LINE_SEG,
      lv + 1,
      buildDefaultLineSeg(availWidthHwp, fontHwp, nchars, para.props, vertPos)
    ),
    ...ctrlRecords
  ];
}
function mkTableCtrl(wHwp, hHwp, instanceId, align = "left") {
  const alignFlags = { left: 0, center: 1, right: 2, justify: 3 }[align] ?? 0;
  return new BufWriter().u32(CTRL_TABLE2).u32(136978961).i32(0).i32(0).u32(wHwp).u32(hHwp).i32(7).u16(140).u16(140).u16(140).u16(140).u32(instanceId).i32(alignFlags).u16(0).build();
}
function mkTableRecord(rowCnt, colCnt, rowHwp, bfId) {
  const w = new BufWriter();
  w.u32(67108870).u16(rowCnt).u16(colCnt).u16(0);
  w.u16(510).u16(510).u16(141).u16(141);
  for (const h of rowHwp) w.u16(Math.max(1, h & 65535));
  w.u16(bfId).u16(0);
  return w.build();
}
function mkCellListHeader(paraCount, row, col, rs, cs, wHwp, hHwp, bfId, padL = 141, padR = 141, padT = 141, padB = 141) {
  return new BufWriter().u16(paraCount).u32(0).u16(0).u16(col).u16(row).u16(rs).u16(cs).u32(wHwp).u32(hHwp).u16(padL).u16(padR).u16(padT).u16(padB).u16(bfId).zeros(13).build();
}
var DEFAULT_ROW_HEIGHT_PT = 14;
function encodeGrid4(grid, bank, lv, idGen, availWidthHwp) {
  const records = [];
  const rowCnt = grid.kids.length;
  const colCnt = Math.max(1, grid.kids[0]?.kids.length ?? 1);
  const cwPt = grid.props.colWidths ?? [];
  const totalPt = cwPt.reduce((s, w) => s + w, 0) || 453;
  const defColPt = totalPt / colCnt;
  const defStroke = grid.props.defaultStroke ?? bank.DEF_STROKE;
  const defBfId = bank.addBorderFill(defStroke);
  const rowHwp = grid.kids.map(
    (row) => row.heightPt != null && row.heightPt > 0 ? Metric.ptToHwp(row.heightPt) : Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT)
  );
  const tblWPt = cwPt.length > 0 ? cwPt.reduce((s, w) => s + w, 0) : totalPt;
  const tblHPt = grid.kids.reduce(
    (s, row) => s + (row.heightPt != null && row.heightPt > 0 ? row.heightPt : DEFAULT_ROW_HEIGHT_PT),
    0
  );
  const tblInstanceId = idGen();
  const tblAlign = grid.props.align ?? "left";
  records.push(
    mkRec(
      TAG_CTRL_HEADER2,
      lv,
      mkTableCtrl(
        Metric.ptToHwp(tblWPt),
        Metric.ptToHwp(tblHPt),
        tblInstanceId,
        tblAlign
      )
    )
  );
  records.push(
    mkRec(TAG_TABLE, lv + 1, mkTableRecord(rowCnt, colCnt, rowHwp, defBfId))
  );
  for (let r = 0; r < grid.kids.length; r++) {
    for (let c = 0; c < grid.kids[r].kids.length; c++) {
      const cell = grid.kids[r].kids[c];
      const wHwp = Metric.ptToHwp(cwPt[c] ?? defColPt);
      const hHwp = rowHwp[r];
      const cp = cell.props;
      const hasPerSide = cp.top || cp.bot || cp.left || cp.right;
      const bfId = hasPerSide ? bank.addBorderFillPerSide(
        cp.left ?? defStroke,
        cp.right ?? defStroke,
        cp.top ?? defStroke,
        cp.bot ?? defStroke,
        cp.bg
      ) : bank.addBorderFill(defStroke, cp.bg);
      const paras = cell.kids.length > 0 ? cell.kids : [{ tag: "para", props: {}, kids: [] }];
      const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : 510;
      const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : 510;
      const padT = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : 141;
      const padB = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : 141;
      records.push(
        mkRec(
          TAG_LIST_HEADER2,
          lv + 1,
          mkCellListHeader(
            paras.length,
            r,
            c,
            cell.rs,
            cell.cs,
            wHwp,
            hHwp,
            bfId,
            padL,
            padR,
            padT,
            padB
          )
        )
      );
      const cellWidthHwp = Metric.ptToHwp(cwPt[c] ?? defColPt);
      for (const para of paras) {
        records.push(
          ...encodePara3(para, bank, lv + 2, idGen(), cellWidthHwp)
        );
      }
    }
  }
  return records;
}
function mkSectionCtrl() {
  return new BufWriter().u32(CTRL_SECD).u32(0).u32(1134).u16(16384).u16(31).zeros(31).build();
}
function buildSectionParagraph(dims, instanceId) {
  const SECD_CTRL_MASK = 1 << 2;
  const nchars = 9;
  const availWidthHwp = Math.max(
    1e3,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr)
  );
  return [
    mkRec(
      TAG_PARA_HEADER2,
      0,
      mkParaHeader(nchars, SECD_CTRL_MASK, 0, 1, 1, instanceId)
    ),
    mkRec(TAG_PARA_TEXT2, 1, mkSecdParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE2, 1, mkParaCharShape([[0, 0]])),
    mkRec(
      TAG_PARA_LINE_SEG,
      1,
      buildDefaultLineSeg(availWidthHwp, 1e3, nchars)
    ),
    mkRec(TAG_CTRL_HEADER2, 1, mkSectionCtrl()),
    mkRec(TAG_PAGE_DEF2, 2, mkPageDef(dims)),
    mkRec(TAG_FOOTNOTE_SHAPE, 2, new Uint8Array(28)),
    mkRec(TAG_FOOTNOTE_SHAPE, 2, new Uint8Array(28))
  ];
}
function flatImgNodes(kids) {
  const result = [];
  for (const kid of kids) {
    if (kid.tag === "img") result.push(kid);
    else if (kid.tag === "link" && Array.isArray(kid.kids))
      result.push(...flatImgNodes(kid.kids));
  }
  return result;
}
function b64Matches(binImg, b64) {
  const a = TextKit.base64Encode(binImg.data).replace(/\s/g, "");
  const b = b64.replace(/\s/g, "");
  return a === b;
}
function buildBodyTextStream(doc, bank, images) {
  const chunks = [];
  const dims = doc.kids[0]?.dims ?? A4;
  let instanceIdCounter = 1;
  const idGen = () => instanceIdCounter++;
  const availWidthHwp = Math.max(
    1e3,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr)
  );
  for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);
  const TABLE_CTRL_MASK = 1 << 11;
  let vertPos = 0;
  for (const sheet of doc.kids) {
    for (const node of sheet.kids) {
      if (node.tag === "para") {
        const para = node;
        const hasPageBreak = para.kids.some(
          (k) => k.tag === "span" && k.kids.some((c) => c.tag === "pb")
        );
        let paraMask = hasPageBreak ? 1 << 2 : 0;
        const hasCourier = (kids) => kids.some(
          (k) => k.tag === "span" && k.props.font?.toLowerCase().includes("courier") || k.tag === "link" && hasCourier(k.kids)
        );
        const isCode = para.props.styleId?.toLowerCase().includes("code") || hasCourier(para.kids);
        if (isCode) {
          const gridNode = {
            tag: "grid",
            props: {
              colWidths: [Metric.hwpToPt(availWidthHwp)],
              defaultStroke: { kind: "solid", pt: 0.5, color: "aaaaaa" }
            },
            kids: [
              {
                tag: "row",
                kids: [
                  {
                    tag: "cell",
                    rs: 1,
                    cs: 1,
                    props: { bg: "f4f4f4" },
                    kids: [para]
                  }
                ]
              }
            ]
          };
          chunks.push(
            mkRec(
              TAG_PARA_HEADER2,
              0,
              mkParaHeader(9, TABLE_CTRL_MASK | paraMask, 0, 1, 1, idGen())
            )
          );
          chunks.push(mkRec(TAG_PARA_TEXT2, 1, mkTableParaText()));
          chunks.push(mkRec(TAG_PARA_CHAR_SHAPE2, 1, mkParaCharShape([[0, 0]])));
          chunks.push(
            mkRec(
              TAG_PARA_LINE_SEG,
              1,
              buildDefaultLineSeg(availWidthHwp, 1e3, 9, void 0, vertPos)
            )
          );
          vertPos += Metric.ptToHwp(20);
          for (const r of encodeGrid4(gridNode, bank, 1, idGen, availWidthHwp))
            chunks.push(r);
          continue;
        }
        const imgNodes = flatImgNodes(para.kids);
        if (imgNodes.length > 0) {
          for (const img of imgNodes) {
            const binImg = images.find((b) => b64Matches(b, img.b64));
            if (binImg) {
              for (const r of encodePicPara(
                img,
                binImg.id,
                bank,
                0,
                idGen,
                availWidthHwp
              )) {
                chunks.push(r);
              }
              vertPos += Metric.ptToHwp(img.h ?? 100);
            }
          }
          const textKids = para.kids.filter(
            (k) => k.tag !== "img" && k.tag !== "link"
          );
          if (textKids.length > 0) {
            const textPara = {
              tag: "para",
              props: para.props,
              kids: textKids
            };
            for (const r of encodePara3(
              textPara,
              bank,
              0,
              idGen(),
              availWidthHwp,
              paraMask,
              vertPos
            )) {
              if (r[0] === (TAG_PARA_HEADER2 & 255)) {
              }
              chunks.push(r);
            }
            const fontHwp_img = textKids.find(
              (k) => k.tag === "span" && k.props?.pt
            )?.props.pt ? Metric.ptToHwp(
              textKids.find(
                (k) => k.tag === "span" && k.props?.pt
              ).props.pt
            ) : 1e3;
            const lineSpacePct_img = Math.round((para.props.lineHeight ?? 1.6) * 100);
            vertPos += Math.round(fontHwp_img * lineSpacePct_img / 100);
          }
        } else {
          for (const r of encodePara3(
            para,
            bank,
            0,
            idGen(),
            availWidthHwp,
            paraMask,
            vertPos
          ))
            chunks.push(r);
          const fontHwp_para = para.kids.find(
            (k) => k.tag === "span" && k.props?.pt
          )?.props.pt ? Metric.ptToHwp(
            para.kids.find(
              (k) => k.tag === "span" && k.props?.pt
            ).props.pt
          ) : 1e3;
          const lineSpacePct_para = para.props.lineHeight ? Math.round(para.props.lineHeight * 100) : 160;
          vertPos += Math.round(fontHwp_para * lineSpacePct_para / 100);
        }
      } else if (node.tag === "grid") {
        chunks.push(
          mkRec(
            TAG_PARA_HEADER2,
            0,
            mkParaHeader(9, TABLE_CTRL_MASK, 0, 1, 1, idGen())
          )
        );
        chunks.push(mkRec(TAG_PARA_TEXT2, 1, mkTableParaText()));
        chunks.push(mkRec(TAG_PARA_CHAR_SHAPE2, 1, mkParaCharShape([[0, 0]])));
        chunks.push(
          mkRec(
            TAG_PARA_LINE_SEG,
            1,
            buildDefaultLineSeg(availWidthHwp, 1e3, 9, void 0, vertPos)
          )
        );
        vertPos += Metric.ptToHwp(20);
        for (const r of encodeGrid4(
          node,
          bank,
          1,
          idGen,
          availWidthHwp
        ))
          chunks.push(r);
      }
    }
  }
  return concatU8(chunks);
}
function buildHwpFileHeader() {
  const SIZE = 256;
  const buf = new Uint8Array(SIZE);
  const dv = new DataView(buf.buffer);
  const sig = "HWP Document File";
  for (let i = 0; i < sig.length; i++) {
    buf[i] = sig.charCodeAt(i);
  }
  dv.setUint32(32, 83886848, true);
  dv.setUint32(36, 1, true);
  if (buf.length !== SIZE) {
    throw new Error(`FileHeader \uD06C\uAE30 \uC624\uB958: ${buf.length} (\uAE30\uB300: ${SIZE})`);
  }
  if (new TextDecoder().decode(buf.subarray(0, sig.length)) !== sig) {
    throw new Error("FileHeader \uC2DC\uADF8\uB2C8\uCC98 \uC624\uB958");
  }
  if (dv.getUint32(32, true) !== 83886848) {
    throw new Error("FileHeader \uBC84\uC804 \uC624\uB958");
  }
  return buf;
}
function buildHwpOle2(fileHeaderData, docInfoData, section0Data, binImages = []) {
  const SS = 512;
  const ENDOFCHAIN = 4294967294;
  const FREESECT = 4294967295;
  const FATSECT = 4294967293;
  if (fileHeaderData.length < 256) {
    throw new Error(
      `FileHeader \uD06C\uAE30 \uBD80\uC871: ${fileHeaderData.length} (\uCD5C\uC18C 256)`
    );
  }
  function padSector(d) {
    const n = Math.ceil(Math.max(d.length, 1) / SS) * SS;
    if (d.length === n) return d;
    const out2 = new Uint8Array(n);
    out2.set(d);
    return out2;
  }
  const fhPad = padSector(fileHeaderData);
  const diPad = padSector(docInfoData);
  const s0Pad = padSector(section0Data);
  const imgPads = binImages.map((img) => padSector(img.data));
  const fhN = fhPad.length / SS;
  const diN = diPad.length / SS;
  const s0N = s0Pad.length / SS;
  const imgNs = imgPads.map((p) => p.length / SS);
  const totalImgN = imgNs.reduce((s, n) => s + n, 0);
  const numDirEntries = 5 + (binImages.length > 0 ? 1 + binImages.length : 0);
  const dirN = Math.max(2, Math.ceil(numDirEntries / 4));
  let fatN = 1;
  for (let iter = 0; iter < 10; iter++) {
    const total = fatN + dirN + fhN + diN + s0N + totalImgN;
    const needed = Math.ceil(total / 128);
    if (needed <= fatN) break;
    fatN = needed;
  }
  const dir1Sec = fatN;
  const fhSec = dir1Sec + dirN;
  const diSec = fhSec + fhN;
  const s0Sec = diSec + diN;
  const imgSecs = [];
  let curSec = s0Sec + s0N;
  for (const n of imgNs) {
    imgSecs.push(curSec);
    curSec += n;
  }
  const totalSec = curSec;
  const fatBuf = new Uint8Array(fatN * SS).fill(255);
  const setFat = (i, v) => {
    fatBuf[i * 4] = v & 255;
    fatBuf[i * 4 + 1] = v >>> 8 & 255;
    fatBuf[i * 4 + 2] = v >>> 16 & 255;
    fatBuf[i * 4 + 3] = v >>> 24 & 255;
  };
  for (let i = 0; i < fatN; i++) setFat(i, FATSECT);
  for (let i = 0; i < dirN; i++)
    setFat(dir1Sec + i, i + 1 < dirN ? dir1Sec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < fhN; i++)
    setFat(fhSec + i, i + 1 < fhN ? fhSec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < diN; i++)
    setFat(diSec + i, i + 1 < diN ? diSec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < s0N; i++)
    setFat(s0Sec + i, i + 1 < s0N ? s0Sec + i + 1 : ENDOFCHAIN);
  for (let ii = 0; ii < imgNs.length; ii++) {
    const start = imgSecs[ii];
    const n = imgNs[ii];
    for (let i = 0; i < n; i++)
      setFat(start + i, i + 1 < n ? start + i + 1 : ENDOFCHAIN);
  }
  const dirBuf = new Uint8Array(dirN * SS);
  const dv = new DataView(dirBuf.buffer);
  function writeDirEntry(idx, name, type, left, right, child, startSec, size) {
    const base = idx * 128;
    const nl = name.length;
    for (let i = 0; i < nl; i++)
      dv.setUint16(base + i * 2, name.charCodeAt(i), true);
    dv.setUint16(base + 64, (nl + 1) * 2, true);
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = 1;
    dv.setInt32(base + 68, left, true);
    dv.setInt32(base + 72, right, true);
    dv.setInt32(base + 76, child, true);
    dv.setUint32(base + 116, startSec >>> 0, true);
    dv.setUint32(base + 120, size >>> 0, true);
  }
  for (let i = 0; i < dirN * 4; i++) {
    const base = i * 128;
    dv.setInt32(base + 68, -1, true);
    dv.setInt32(base + 72, -1, true);
    dv.setInt32(base + 76, -1, true);
  }
  if (binImages.length > 0) {
    writeDirEntry(0, "Root Entry", 5, -1, -1, 1, ENDOFCHAIN, 0);
    writeDirEntry(1, "FileHeader", 2, -1, 2, -1, fhSec, fileHeaderData.length);
    writeDirEntry(2, "DocInfo", 2, -1, 3, -1, diSec, docInfoData.length);
    writeDirEntry(3, "BodyText", 1, -1, 5, 4, ENDOFCHAIN, 0);
    writeDirEntry(4, "Section0", 2, -1, -1, -1, s0Sec, section0Data.length);
    writeDirEntry(5, "BinData", 1, -1, -1, 6, ENDOFCHAIN, 0);
    for (let ii = 0; ii < binImages.length; ii++) {
      const img = binImages[ii];
      const streamName = `BIN${String(img.id).padStart(4, "0")}.${img.ext}`;
      const sibling = ii + 1 < binImages.length ? 7 + ii : -1;
      writeDirEntry(
        6 + ii,
        streamName,
        2,
        -1,
        sibling,
        -1,
        imgSecs[ii],
        img.data.length
      );
    }
  } else {
    writeDirEntry(0, "Root Entry", 5, -1, -1, 1, ENDOFCHAIN, 0);
    writeDirEntry(1, "FileHeader", 2, -1, 2, -1, fhSec, fileHeaderData.length);
    writeDirEntry(2, "DocInfo", 2, -1, 3, -1, diSec, docInfoData.length);
    writeDirEntry(3, "BodyText", 1, -1, -1, 4, ENDOFCHAIN, 0);
    writeDirEntry(4, "Section0", 2, -1, -1, -1, s0Sec, section0Data.length);
  }
  const HWP_CLSID = [
    32,
    233,
    227,
    192,
    70,
    53,
    207,
    17,
    141,
    129,
    0,
    170,
    0,
    56,
    155,
    113
  ];
  for (let i = 0; i < 16; i++) dirBuf[80 + i] = HWP_CLSID[i];
  const hdr = new Uint8Array(SS);
  const hdv = new DataView(hdr.buffer);
  const MAGIC = [208, 207, 17, 224, 161, 177, 26, 225];
  MAGIC.forEach((b, i) => {
    hdr[i] = b;
  });
  hdv.setUint16(24, 62, true);
  hdv.setUint16(26, 3, true);
  hdv.setUint16(28, 254, true);
  hdv.setUint16(30, 9, true);
  hdv.setUint16(32, 6, true);
  hdv.setUint32(40, fatN, true);
  hdv.setUint32(44, dir1Sec, true);
  hdv.setUint32(48, 0, true);
  hdv.setUint32(52, 4096, true);
  hdv.setUint32(56, ENDOFCHAIN, true);
  hdv.setUint32(60, 0, true);
  hdv.setUint32(64, ENDOFCHAIN, true);
  hdv.setUint32(68, 0, true);
  hdv.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) {
    hdv.setUint32(76 + i * 4, i < fatN ? i : FREESECT, true);
  }
  const out = new Uint8Array(SS + totalSec * SS);
  out.set(hdr, 0);
  for (let i = 0; i < fatN; i++)
    out.set(fatBuf.subarray(i * SS, (i + 1) * SS), SS + i * SS);
  for (let i = 0; i < dirN; i++)
    out.set(dirBuf.subarray(i * SS, (i + 1) * SS), SS + (dir1Sec + i) * SS);
  out.set(fhPad, SS + fhSec * SS);
  out.set(diPad, SS + diSec * SS);
  out.set(s0Pad, SS + s0Sec * SS);
  for (let ii = 0; ii < imgPads.length; ii++)
    out.set(imgPads[ii], SS + imgSecs[ii] * SS);
  return out;
}
function concatU8(arrays) {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) {
    out.set(a, off);
    off += a.length;
  }
  return out;
}
function validateOle2Magic(hwp) {
  const OLE_MAGIC = [208, 207, 17, 224, 161, 177, 26, 225];
  return OLE_MAGIC.every((b, i) => hwp[i] === b);
}
var HwpEncoder = class extends BaseEncoder {
  getFormat() {
    return "hwp";
  }
  getAliases() {
    return ["application/vnd.hancom.hwp"];
  }
  async encode(doc) {
    try {
      let registerImg2 = function(img) {
        const key = img.b64.substring(0, 50);
        if (seenB64.has(key)) return;
        seenB64.add(key);
        const raw = TextKit.base64Decode(img.b64);
        const ext = img.mime === "image/png" ? "png" : img.mime === "image/gif" ? "gif" : img.mime === "image/bmp" ? "bmp" : "jpg";
        images.push({ id: binIdCounter++, ext, data: new Uint8Array(raw) });
      }, collectImages3 = function(node) {
        if (node.tag === "para") {
          for (const img of flatImgNodes(node.kids)) registerImg2(img);
        } else if (node.tag === "grid") {
          for (const row of node.kids)
            for (const cell of row.kids)
              for (const para of cell.kids) collectImages3(para);
        }
      };
      var registerImg = registerImg2, collectImages2 = collectImages3;
      const bank = new HwpStyleBank();
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectNode(node, bank);
      }
      const images = [];
      const seenB64 = /* @__PURE__ */ new Set();
      let binIdCounter = 1;
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectImages3(node);
      }
      const docInfoRaw = buildDocInfoStream(bank, images);
      const bodyRaw = buildBodyTextStream(doc, bank, images);
      const docInfoCmp = pako3.deflateRaw(docInfoRaw);
      const bodyCmp = pako3.deflateRaw(bodyRaw);
      const fileHdr = buildHwpFileHeader();
      if (fileHdr.length !== 256) {
        return fail(
          `HwpEncoder: FileHeader \uD06C\uAE30 \uC624\uB958 - ${fileHdr.length} bytes (\uAE30\uB300: 256 bytes)`
        );
      }
      const hwp = buildHwpOle2(fileHdr, docInfoCmp, bodyCmp, images);
      if (!validateOle2Magic(hwp)) {
        return fail("HwpEncoder: OLE2 \uB9E4\uC9C1 \uBC14\uC774\uD2B8 \uC624\uB958");
      }
      if (hwp.length < 512) {
        return fail(
          `HwpEncoder: HWP \uD30C\uC77C \uD06C\uAE30 \uBD80\uC871 - ${hwp.length} bytes (\uCD5C\uC18C 512 bytes)`
        );
      }
      return succeed(hwp);
    } catch (e) {
      return fail(`HwpEncoder: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
};
registry.registerEncoder(new HwpEncoder());

// src/pipeline/Pipeline.ts
var Pipeline = class _Pipeline {
  constructor(raw, srcFmt) {
    this.raw = raw;
    this.srcFmt = srcFmt;
  }
  /** 파일을 열고 포맷을 자동 감지하거나 명시 */
  static open(input, fmt) {
    if (typeof input === "string") {
      return new _Pipeline(new TextEncoder().encode(input), fmt ?? "md");
    }
    return new _Pipeline(input, fmt ?? detectFormat(input));
  }
  /** File/Blob 비동기 입력 */
  static async openAsync(input, fmt) {
    if (input instanceof Uint8Array || typeof input === "string") {
      return _Pipeline.open(input, fmt);
    }
    const buf = await input.arrayBuffer();
    const data = new Uint8Array(buf);
    const detectedFmt = fmt ?? (input instanceof File ? getExt(input.name) : void 0) ?? detectFormat(data);
    return new _Pipeline(data, detectedFmt);
  }
  /** 목표 포맷으로 변환 */
  async to(targetFmt, options) {
    const decoder = registry.getDecoder(this.srcFmt);
    const encoder = registry.getEncoder(targetFmt);
    if (!decoder) return fail(`\uC9C0\uC6D0\uD558\uC9C0 \uC54A\uB294 \uC785\uB825 \uD3EC\uB9F7: ${this.srcFmt}`);
    if (!encoder) return fail(`\uC9C0\uC6D0\uD558\uC9C0 \uC54A\uB294 \uCD9C\uB825 \uD3EC\uB9F7: ${targetFmt}`);
    const docResult = await decoder.decode(this.raw);
    if (!docResult.ok) return docResult;
    const encResult = await encoder.encode(docResult.data, options);
    if (!encResult.ok) return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };
    return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };
  }
  /** DocRoot만 추출 (인코딩 없이) */
  async inspect() {
    const decoder = registry.getDecoder(this.srcFmt);
    if (!decoder) return fail(`\uB514\uCF54\uB354 \uC5C6\uC74C: ${this.srcFmt}`);
    return decoder.decode(this.raw);
  }
};
function detectFormat(data) {
  if (data[0] === 208 && data[1] === 207 && data[2] === 17 && data[3] === 224) return "hwp";
  if (data[0] === 80 && data[1] === 75) {
    const str = new TextDecoder("utf-8", { fatal: false }).decode(data.slice(0, 4096));
    if (str.includes("wordprocessingml")) return "docx";
    if (str.includes("ha-xml")) return "hwpx";
    if (str.includes("hwpml/")) return "hwpx";
    if (str.includes("word/")) return "docx";
    return "hwpx";
  }
  return "md";
}
function getExt(name) {
  const parts = name.split(".");
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : void 0;
}

// src/walk/TreeWalker.ts
function walkNode(node, cb, parent = null, depth = 0) {
  const result = cb(node, parent, depth);
  if (result === "stop") return false;
  if ("kids" in node && Array.isArray(node.kids)) {
    for (const kid of node.kids) {
      if (!walkNode(kid, cb, node, depth + 1)) return false;
    }
  }
  return true;
}
var TreeWalker = class {
  walk(root, cb) {
    walkNode(root, cb);
  }
  findAll(root, predicate) {
    const results = [];
    walkNode(root, (n) => {
      if (predicate(n)) results.push(n);
    });
    return results;
  }
  extractText(root) {
    const parts = [];
    walkNode(root, (n) => {
      if (n.tag === "txt") parts.push(n.content);
      if (n.tag === "br") parts.push("\n");
      if (n.tag === "pb") parts.push("\n\n");
    });
    return parts.join("");
  }
};

// src/walk/tree-ops.ts
function countNodes(root) {
  const counts = {};
  walkNode(root, (n) => {
    counts[n.tag] = (counts[n.tag] ?? 0) + 1;
  });
  return counts;
}
function validateRoot(root) {
  const errors = [];
  if (root.tag !== "root") errors.push('Root node must have tag "root"');
  if (!Array.isArray(root.kids)) errors.push("Root.kids must be an array");
  if (root.kids.length === 0) errors.push("Document has no sheets");
  walkNode(root, (n) => {
    if (n.tag === "cell" && n.kids.length === 0) {
      errors.push("CellNode must have at least one ParaNode child");
    }
    if (n.tag === "grid" && n.kids.length === 0) {
      errors.push("GridNode must have at least one RowNode");
    }
  });
  return errors;
}
export {
  A4,
  A4_LANDSCAPE,
  ArchiveKit,
  BinaryKit,
  DEFAULT_STROKE,
  Metric,
  Pipeline,
  ShieldedParser,
  TextKit,
  TreeWalker,
  XmlKit,
  buildBr,
  buildCell,
  buildGrid,
  buildImg,
  buildPageNum,
  buildPara,
  buildPb,
  buildRoot,
  buildRow,
  buildSheet,
  buildSpan,
  countNodes,
  fail,
  normalizeDims,
  registry,
  safeAlign,
  safeFont,
  safeFontToKr,
  safeHex,
  safeStrokeDocx,
  safeStrokeHwpx,
  succeed,
  validateRoot,
  walkNode
};
//# sourceMappingURL=index.mjs.map