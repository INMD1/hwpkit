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
function safeFont(raw) {
  return raw ?? "Malgun Gothic";
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

// src/toolkit/TableGeometry.ts
function fitColumnWidths(source, columnCount, maxTotal, minWidth = 1) {
  const n = Math.max(0, Math.floor(columnCount));
  if (n === 0) return [];
  const limit = Math.max(n, Math.floor(Number.isFinite(maxTotal) ? maxTotal : n));
  const floorWidth = Math.max(1, Math.min(Math.floor(minWidth), Math.floor(limit / n)));
  const values = Array.from({ length: n }, (_, index) => {
    const value = Number(source[index]);
    return Number.isFinite(value) && value > 0 ? value : 0;
  });
  const knownTotal = values.reduce((sum, value) => sum + value, 0);
  const missing = values.reduce((count, value) => count + (value <= 0 ? 1 : 0), 0);
  if (missing === n) {
    values.fill(limit / n);
  } else if (missing > 0) {
    const remaining2 = limit - knownTotal;
    const knownAverage = knownTotal / (n - missing);
    const fill = remaining2 > 0 ? remaining2 / missing : knownAverage;
    for (let i = 0; i < n; i++) if (values[i] <= 0) values[i] = Math.max(1, fill);
  }
  const rawTotal = values.reduce((sum, value) => sum + value, 0);
  const target = Math.max(
    floorWidth * n,
    Math.min(limit, Math.round(rawTotal > 0 ? rawTotal : limit))
  );
  const exact = new Array(n).fill(0);
  const active = new Set(Array.from({ length: n }, (_, index) => index));
  let remaining = target;
  while (active.size > 0) {
    const weightTotal = [...active].reduce((sum, index) => sum + values[index], 0);
    const tooSmall = [...active].filter(
      (index) => weightTotal <= 0 || remaining * values[index] / weightTotal < floorWidth
    );
    if (tooSmall.length === 0) {
      for (const index of active) exact[index] = remaining * values[index] / weightTotal;
      break;
    }
    for (const index of tooSmall) {
      exact[index] = floorWidth;
      remaining -= floorWidth;
      active.delete(index);
    }
  }
  const result = exact.map(Math.floor);
  let residual = target - result.reduce((sum, value) => sum + value, 0);
  const order = exact.map((value, index) => ({ index, fraction: value - Math.floor(value) })).sort((a, b) => b.fraction - a.fraction || a.index - b.index);
  for (let i = 0; residual > 0; i++, residual--) {
    result[order[i % order.length].index]++;
  }
  return result;
}
function inferColumnWidths(columnCount, observations) {
  const n = Math.max(0, Math.floor(columnCount));
  if (n === 0) return [];
  const constraints = observations.map((item) => ({
    start: Math.max(0, Math.floor(item.start)),
    span: Math.max(1, Math.floor(item.span)),
    width: Number(item.width)
  })).filter(
    (item) => Number.isFinite(item.width) && item.width > 0 && item.start < n && item.start + item.span <= n
  );
  if (constraints.length === 0) return new Array(n).fill(0);
  const priorSums = new Array(n).fill(0);
  const priorWeights = new Array(n).fill(0);
  for (const item of constraints) {
    const estimate = item.width / item.span;
    const weight = 1 / item.span;
    for (let col = item.start; col < item.start + item.span; col++) {
      priorSums[col] += estimate * weight;
      priorWeights[col] += weight;
    }
  }
  const fullWidths = constraints.filter((item) => item.start === 0 && item.span === n).map((item) => item.width).sort((a, b) => a - b);
  const totalHint = fullWidths.length ? fullWidths[Math.floor(fullWidths.length / 2)] : constraints.reduce((largest, item) => Math.max(largest, item.width), 0);
  const defaultWidth = totalHint / n;
  const prior = priorSums.map(
    (sum, col) => priorWeights[col] > 0 ? sum / priorWeights[col] : defaultWidth
  );
  if (n > 256) {
    const priorTotal = prior.reduce((sum, width) => sum + Math.max(0, width), 0);
    const scale = priorTotal > 0 ? totalHint / priorTotal : 1;
    return prior.map((width) => Math.max(1e-4, width * scale));
  }
  const matrix = Array.from({ length: n }, () => new Array(n).fill(0));
  const rhs = new Array(n).fill(0);
  for (const item of constraints) {
    const weight = 1 / Math.sqrt(item.span);
    for (let a = item.start; a < item.start + item.span; a++) {
      rhs[a] += weight * item.width;
      for (let b = item.start; b < item.start + item.span; b++) {
        matrix[a][b] += weight;
      }
    }
  }
  const maxDiagonal = Math.max(1, ...matrix.map((row, i) => row[i]));
  const ridge = maxDiagonal * 1e-8;
  for (let i = 0; i < n; i++) {
    matrix[i][i] += ridge;
    rhs[i] += ridge * prior[i];
  }
  const solved = solveLinearSystem(matrix, rhs);
  const minWidth = Math.max(1e-4, totalHint * 1e-8);
  const widths = (solved ?? prior).map(
    (value, col) => Number.isFinite(value) && value > minWidth ? value : Math.max(minWidth, prior[col])
  );
  for (let pass = 0; pass < 80; pass++) {
    let maxRelativeResidual = 0;
    for (const item of constraints) {
      let actual = 0;
      for (let col = item.start; col < item.start + item.span; col++) {
        actual += widths[col];
      }
      const residual = item.width - actual;
      maxRelativeResidual = Math.max(
        maxRelativeResidual,
        Math.abs(residual) / Math.max(1, item.width)
      );
      if (Math.abs(residual) <= 1e-9) continue;
      const adjustable = widths.slice(item.start, item.start + item.span).reduce((sum, value) => sum + Math.max(minWidth, value), 0);
      for (let col = item.start; col < item.start + item.span; col++) {
        const share = Math.max(minWidth, widths[col]) / adjustable;
        widths[col] = Math.max(minWidth, widths[col] + residual * share);
      }
    }
    if (maxRelativeResidual < 1e-8) break;
  }
  return widths;
}
function solveLinearSystem(matrix, rhs) {
  const n = rhs.length;
  const augmented = matrix.map((row, i) => [...row, rhs[i]]);
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let row = col + 1; row < n; row++) {
      if (Math.abs(augmented[row][col]) > Math.abs(augmented[pivot][col])) {
        pivot = row;
      }
    }
    if (Math.abs(augmented[pivot][col]) < 1e-12) return null;
    if (pivot !== col) [augmented[pivot], augmented[col]] = [augmented[col], augmented[pivot]];
    const divisor = augmented[col][col];
    for (let j = col; j <= n; j++) augmented[col][j] /= divisor;
    for (let row = 0; row < n; row++) {
      if (row === col) continue;
      const factor = augmented[row][col];
      if (Math.abs(factor) < 1e-15) continue;
      for (let j = col; j <= n; j++) {
        augmented[row][j] -= factor * augmented[col][j];
      }
    }
  }
  return augmented.map((row) => row[n]);
}

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
    const modernSecPr = root?.["hh:secPrList"]?.[0]?.["hh:secPr"]?.[0] ?? root?.["hh:SECPRLST"]?.[0]?.["hh:SECPR"]?.[0];
    const modernPagePr = modernSecPr?.["hh:pagePr"]?.[0]?._attr ?? modernSecPr?.["hh:PAGEPR"]?.[0]?._attr;
    if (modernPagePr) {
      const margin = modernSecPr?.["hh:pagePr"]?.[0]?.["hh:margin"]?.[0]?._attr ?? modernSecPr?.["hh:PAGEPR"]?.[0]?.["hh:MARGIN"]?.[0]?._attr ?? {};
      let ew2 = Number(modernPagePr.width ?? modernPagePr.Width ?? 59528);
      let eh2 = Number(modernPagePr.height ?? modernPagePr.Height ?? 84188);
      const landscape = String(modernPagePr.landscape ?? "").toUpperCase();
      if ((landscape === "NARROWLY" || landscape === "LANDSCAPE") && ew2 < eh2) {
        [ew2, eh2] = [eh2, ew2];
      }
      const mt2 = Number(margin.top ?? margin.TopMargin ?? 5670);
      const mb2 = Number(margin.bottom ?? margin.BottomMargin ?? 4252);
      const ml2 = Number(margin.left ?? margin.LeftMargin ?? 8504);
      const mr2 = Number(margin.right ?? margin.RightMargin ?? 8504);
      const header2 = Number(margin.header ?? margin.HeaderMargin ?? 0);
      const footer2 = Number(margin.footer ?? margin.FooterMargin ?? 0);
      return {
        wPt: Metric.hwpToPt(ew2),
        hPt: Metric.hwpToPt(eh2),
        mt: Metric.hwpToPt(mt2),
        mb: Metric.hwpToPt(mb2),
        ml: Metric.hwpToPt(ml2),
        mr: Metric.hwpToPt(mr2),
        headerPt: Metric.hwpToPt(Math.max(0, header2)),
        footerPt: Metric.hwpToPt(Math.max(0, footer2)),
        orient: ew2 > eh2 ? "landscape" : "portrait"
      };
    }
    const refList = root?.["hh:refList"]?.[0] ?? root?.["hh:REFLIST"]?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return null;
    const secPrList = refList?.["hh:SECPRLST"]?.[0]?.["hh:SECPR"] ?? refList?.SECPRLST?.[0]?.SECPR;
    const sec = Array.isArray(secPrList) ? secPrList[0] : secPrList;
    if (!sec) return null;
    const pa = sec?.["hh:PAGEPROPERTY"]?.[0]?._attr ?? sec?.PAGEPROPERTY?.[0]?._attr;
    if (!pa) return null;
    const ew = Number(pa.Width ?? 59528);
    const eh = Number(pa.Height ?? 84188);
    const mt = Number(pa.TopMargin ?? 5670);
    const mb = Number(pa.BottomMargin ?? 4252);
    const ml = Number(pa.LeftMargin ?? 8504);
    const mr = Number(pa.RightMargin ?? 8504);
    const header = Number(pa.HeaderMargin ?? 0);
    const footer = Number(pa.FooterMargin ?? 0);
    return {
      wPt: Metric.hwpToPt(ew),
      hPt: Metric.hwpToPt(eh),
      mt: Metric.hwpToPt(mt),
      mb: Metric.hwpToPt(mb),
      ml: Metric.hwpToPt(ml),
      mr: Metric.hwpToPt(mr),
      headerPt: Metric.hwpToPt(Math.max(0, header)),
      footerPt: Metric.hwpToPt(Math.max(0, footer)),
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
      if (attr.textColor) info.color = normalizeHwpxTextColor(attr.textColor);
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
        if (leftVal !== 0) indentPt = Metric.hwpToPt(leftVal / 2);
        if (rightVal !== 0) indentRightPt = Metric.hwpToPt(rightVal / 2);
        if (indentVal !== 0) firstLineIndentPt = Metric.hwpToPt(indentVal / 2);
        if (prevVal > 0) spaceBefore = Metric.hwpToPt(prevVal / 2);
        if (nextVal > 0) spaceAfter = Metric.hwpToPt(nextVal / 2);
      }
      if (lineSpEl) {
        const lsAttr = lineSpEl._attr ?? {};
        const lsType = lsAttr.type ?? "PERCENT";
        const lsVal = Number(lsAttr.value ?? 160);
        if (lsType === "PERCENT" && lsVal > 0) {
          lineHeight = lsVal / 100;
        } else if ((lsType === "FIXED" || lsType === "AT_LEAST") && lsVal > 0) {
          lineHeightFixed = Metric.hwpToPt(lsVal / 2);
        }
      }
      map.set(id, {
        align,
        indentPt,
        indentRightPt,
        firstLineIndentPt,
        spaceBefore: spaceBefore ?? 0,
        spaceAfter: spaceAfter ?? 0,
        lineHeight: lineHeightFixed === void 0 ? lineHeight ?? 1.6 : void 0,
        lineHeightFixed,
        lineHeightRule: lineHeightFixed !== void 0 ? lineSpEl?._attr?.type === "AT_LEAST" ? "atLeast" : "exact" : void 0
      });
    }
  } catch {
  }
  return map;
}
function addParaItems(p, items) {
  const runs = getTag(p, "hp:run", "hp:RUN");
  for (const run of runs) {
    const tbls = getTag(run, "hp:tbl", "hp:TABLE");
    if (tbls.length > 0) {
      for (const tbl of tbls) {
        items.push({ type: "table", node: tbl });
      }
    }
  }
  items.push({ type: "para", node: p });
}
function decodeSection(sec, dims, ctx) {
  const firstParas = getTag(sec, "hp:p", "hp:P");
  const pageDims = extractSectionDims(sec) ?? extractSecPrDims(firstParas[0]) ?? dims;
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
  let pw = Number(pagePr.width ?? 59528);
  let ph = Number(pagePr.height ?? 84188);
  const landscape = String(pagePr.landscape ?? "").toUpperCase();
  if ((landscape === "NARROWLY" || landscape === "LANDSCAPE") && pw < ph) {
    [pw, ph] = [ph, pw];
  }
  const mt = Number(margin.top ?? 5670);
  const mb = Number(margin.bottom ?? 4252);
  const ml = Number(margin.left ?? 8504);
  const mr = Number(margin.right ?? 8504);
  const header = Number(margin.header ?? 0);
  const footer = Number(margin.footer ?? 0);
  return {
    wPt: Metric.hwpToPt(pw),
    hPt: Metric.hwpToPt(ph),
    mt: Metric.hwpToPt(mt),
    mb: Metric.hwpToPt(mb),
    ml: Metric.hwpToPt(ml),
    mr: Metric.hwpToPt(mr),
    headerPt: Metric.hwpToPt(Math.max(0, header)),
    footerPt: Metric.hwpToPt(Math.max(0, footer)),
    orient: pw > ph ? "landscape" : "portrait"
  };
}
function extractSectionDims(sec) {
  try {
    const secPr = sec?.["hp:secPr"]?.[0] ?? sec?.["hp:SECPR"]?.[0];
    return secPr ? parseSecPrDims(secPr) : null;
  } catch {
    return null;
  }
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
  const styleIdRef = Number(pAttr.styleIDRef ?? pAttr.styleIdRef ?? pAttr.styleID ?? pAttr.styleId);
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
  const props = {
    align: safeAlign(align === "JUSTIFY" ? "LEFT" : align),
    spaceBefore: 0,
    spaceAfter: 0,
    lineHeight: 1.6
  };
  if (Number.isFinite(styleIdRef) && styleIdRef >= 0) props.hwpStyleId = styleIdRef;
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
    if (paraPrDef.lineHeightFixed !== void 0) {
      props.lineHeight = void 0;
      props.lineHeightFixed = paraPrDef.lineHeightFixed;
    }
    if (paraPrDef.lineHeightRule !== void 0)
      props.lineHeightRule = paraPrDef.lineHeightRule;
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
      const spanProps = content === "" ? {} : resolveCharPr(run, ctx);
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
    color: normalizeHwpxTextColor(ca.TextColor ?? ca.textColor),
    bg: safeHex(ca.BgColor ?? ca.bgColor)
  };
}
function normalizeHwpxTextColor(raw) {
  const color = safeHex(raw);
  return color === "000000" ? void 0 : color;
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
      bmp: "image/bmp",
      wmf: "image/x-wmf",
      emf: "image/x-emf"
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
  const textWrap = pic?._attr?.textWrap ?? pic?.pic?.[0]?._attr?.textWrap ?? "TOP_AND_BOTTOM";
  const layout = extractHwpxObjectLayout(posAttr, textWrap);
  applyHwpxOutMargin(layout, pic);
  return layout;
}
function extractHwpxTableLayout(tbl) {
  const posAttr = tbl?.["hp:pos"]?.[0]?._attr ?? tbl?.pos?.[0]?._attr ?? {};
  const textWrap = tbl?._attr?.textWrap ?? "TOP_AND_BOTTOM";
  const layout = extractHwpxObjectLayout(posAttr, textWrap);
  if (layout.wrap === "inline") return void 0;
  applyHwpxOutMargin(layout, tbl);
  return layout;
}
function extractHwpxObjectLayout(posAttr, textWrap) {
  const treatAsChar = posAttr.treatAsChar === "1" || posAttr.treatAsChar === "true";
  if (treatAsChar) return { wrap: "inline" };
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
    FRONT_TEXT: "front",
    IN_FRONT_OF_TEXT: "front"
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
function applyHwpxOutMargin(layout, obj) {
  const outMargin = obj?.["hp:outMargin"]?.[0]?._attr ?? obj?.outMargin?.[0]?._attr;
  if (!outMargin) return;
  const assign = (attr, key) => {
    const raw = outMargin[attr];
    if (raw === void 0) return;
    const value = Number(raw);
    if (Number.isFinite(value) && value >= 0) layout[key] = Metric.hwpToPt(value);
  };
  assign("top", "distT");
  assign("bottom", "distB");
  assign("left", "distL");
  assign("right", "distR");
}
function validHwpxCellPadding(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 && parsed < 65535 ? parsed : void 0;
}
function decodeGrid(tbl, ctx) {
  const tblAttr = tbl?._attr ?? {};
  const borderFillId = Number(tblAttr.borderFillIDRef ?? 0);
  const borderFill = ctx.borderFills.get(borderFillId);
  const headerRow = tblAttr.repeatHeader === "1";
  const inMarginAttr = tbl?.["hp:inMargin"]?.[0]?._attr ?? {};
  const tablePadding = {
    left: validHwpxCellPadding(inMarginAttr.left) ?? 510,
    right: validHwpxCellPadding(inMarginAttr.right) ?? 510,
    top: validHwpxCellPadding(inMarginAttr.top) ?? 141,
    bottom: validHwpxCellPadding(inMarginAttr.bottom) ?? 141
  };
  const gridProps = {
    headerRow: headerRow || void 0,
    cellPadL: Metric.hwpToPt(tablePadding.left),
    cellPadR: Metric.hwpToPt(tablePadding.right),
    cellPadT: Metric.hwpToPt(tablePadding.top),
    cellPadB: Metric.hwpToPt(tablePadding.bottom)
  };
  if (borderFill?.stroke) gridProps.defaultStroke = borderFill.stroke;
  const layout = extractHwpxTableLayout(tbl);
  if (layout) gridProps.layout = layout;
  const posAttr = tbl?.["hp:pos"]?.[0]?._attr ?? {};
  if (posAttr.horzAlign) {
    const alignMap = {
      LEFT: "left",
      RIGHT: "right",
      CENTER: "center",
      JUSTIFY: "justify"
    };
    const a = alignMap[posAttr.horzAlign];
    if (a) gridProps.align = a;
  }
  const rowArr = getTag(tbl, "hp:tr", "hp:ROW");
  let detectedCols = Math.max(0, Number(tblAttr.colCnt ?? tblAttr.ColCnt ?? 0));
  const widthConstraints = [];
  for (const row of rowArr) {
    let sequentialCol = 0;
    for (const cell of getTag(row, "hp:tc", "hp:CELL")) {
      const spanAttr = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const addrAttr = cell?.["hp:cellAddr"]?.[0]?._attr ?? {};
      const span = Math.max(1, Number(spanAttr.colSpan ?? cell?._attr?.ColSpan ?? 1));
      const address = Number(addrAttr.colAddr ?? sequentialCol);
      const start = Number.isFinite(address) && address >= 0 ? address : sequentialCol;
      const width = Number(cell?.["hp:cellSz"]?.[0]?._attr?.width ?? 0);
      if (width > 0) widthConstraints.push({ start, span, width });
      detectedCols = Math.max(detectedCols, start + span);
      sequentialCol = start + span;
    }
  }
  if (detectedCols > 0) {
    const inferred = inferColumnWidths(detectedCols, widthConstraints);
    if (inferred.some((width) => width > 0)) {
      gridProps.colWidths = inferred.map(Metric.hwpToPt);
    }
  }
  const rowNodes = rowArr.map((row) => {
    const cellArr = [...getTag(row, "hp:tc", "hp:CELL")].sort((a, b) => {
      const aa = Number(a?.["hp:cellAddr"]?.[0]?._attr?.colAddr ?? 0);
      const ba = Number(b?.["hp:cellAddr"]?.[0]?._attr?.colAddr ?? 0);
      return aa - ba;
    });
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
      const cellMarginAttr = cell?.["hp:cellMargin"]?.[0]?._attr ?? {};
      const mL = validHwpxCellPadding(cellMarginAttr.left);
      const mR = validHwpxCellPadding(cellMarginAttr.right);
      const mT = validHwpxCellPadding(cellMarginAttr.top);
      const mB = validHwpxCellPadding(cellMarginAttr.bottom);
      if (mL !== void 0 && mL !== tablePadding.left) cellProps.padL = Metric.hwpToPt(mL);
      if (mR !== void 0 && mR !== tablePadding.right) cellProps.padR = Metric.hwpToPt(mR);
      if (mT !== void 0 && mT !== tablePadding.top) cellProps.padT = Metric.hwpToPt(mT);
      if (mB !== void 0 && mB !== tablePadding.bottom) cellProps.padB = Metric.hwpToPt(mB);
      const cellSpan = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cs = Number(cellSpan.colSpan ?? ca.ColSpan ?? 1);
      const rs = Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1);
      const cellKids = [];
      const source = subList ?? cell;
      const sourcePSource = getTag(source, "hp:p", "hp:P");
      for (const sp of sourcePSource) {
        try {
          const runs = getTag(sp, "hp:run", "hp:RUN");
          for (const run of runs) {
            const nestedTbls = getTag(run, "hp:tbl", "hp:TABLE");
            for (const nestedTbl of nestedTbls) {
              try {
                cellKids.push(decodeGrid(nestedTbl, ctx));
              } catch {
              }
            }
          }
          cellKids.push(decodePara(sp, ctx));
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
var TAG_NUMBERING = HWPTAG_BEGIN + 7;
var TAG_BULLET = HWPTAG_BEGIN + 8;
var TAG_PARA_SHAPE = HWPTAG_BEGIN + 9;
var TAG_STYLE = HWPTAG_BEGIN + 10;
var TAG_PARA_HEADER = HWPTAG_BEGIN + 50;
var TAG_PARA_TEXT = HWPTAG_BEGIN + 51;
var TAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52;
var TAG_CTRL_HEADER = HWPTAG_BEGIN + 55;
var TAG_PAGE_DEF = HWPTAG_BEGIN + 57;
var TAG_SHAPE_COMPONENT_PICTURE = HWPTAG_BEGIN + 69;
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
var CTRL_PIC = 611346787;
var CTRL_OBJ = 1868720672;
var CTRL_FIG = 1718183712;
var CTRL_GSO = 1735618336;
var CTRL_HEAD = 1751474532;
var CTRL_FOOT = 1718579060;
var CTRL_ATNO = 1635020399;
var CTRL_SECD = 1936024420;
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
  const info = {
    faceNames: [],
    charShapes: [],
    paraShapes: [],
    borderFills: [],
    styles: [],
    numberings: [],
    bullets: []
  };
  for (const r of recs) {
    try {
      if (r.tag === TAG_FACE_NAME) info.faceNames.push(parseFaceName(r.data));
      if (r.tag === TAG_CHAR_SHAPE) info.charShapes.push(parseCharShape(r.data));
      if (r.tag === TAG_PARA_SHAPE) info.paraShapes.push(parseParaShape(r.data));
      if (r.tag === TAG_BORDER_FILL) info.borderFills.push(parseBorderFill(r.data));
      if (r.tag === TAG_STYLE) info.styles.push(parseStyle(r.data));
      if (r.tag === TAG_NUMBERING) info.numberings.push(parseNumbering(r.data));
      if (r.tag === TAG_BULLET) info.bullets.push(parseBullet(r.data));
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
function parseStyle(d) {
  let offset = 0;
  const readName = () => {
    if (offset + 2 > d.length) throw new Error("truncated STYLE name length");
    const length = BinaryKit.readU16LE(d, offset);
    offset += 2;
    const end = offset + length * 2;
    if (end > d.length) throw new Error("truncated STYLE name");
    const value = new TextDecoder("utf-16le").decode(d.subarray(offset, end));
    offset = end;
    return value;
  };
  const name = readName();
  const engName = readName();
  if (offset + 8 > d.length) throw new Error("truncated STYLE fields");
  offset += 4;
  const paraShapeId = BinaryKit.readU16LE(d, offset);
  const charShapeId = BinaryKit.readU16LE(d, offset + 2);
  return { name, engName, paraShapeId, charShapeId };
}
function parseNumbering(d) {
  const formats = [];
  let offset = 0;
  for (let level = 0; level < 7; level++) {
    if (offset + 14 > d.length) throw new Error("truncated NUMBERING level");
    offset += 12;
    const length = BinaryKit.readU16LE(d, offset);
    offset += 2;
    const end = offset + length * 2;
    if (end > d.length) throw new Error("truncated NUMBERING format");
    formats.push(
      new TextDecoder("utf-16le").decode(d.subarray(offset, end))
    );
    offset = end;
  }
  return { formats };
}
function parseBullet(d) {
  if (d.length < 10) throw new Error("truncated BULLET record");
  return { character: String.fromCharCode(BinaryKit.readU16LE(d, 8)) };
}
function parseCharShape(d) {
  const faceIds = [];
  for (let i = 0; i < 7; i++) faceIds.push(d.length >= (i + 1) * 2 ? BinaryKit.readU16LE(d, i * 2) : 0);
  const height = d.length >= 46 ? BinaryKit.readU32LE(d, 42) : 1e3;
  const attr = d.length >= 50 ? BinaryKit.readU32LE(d, 46) : 0;
  const compactStyleFlags = (attr & 4278190080) === 0;
  const suType = attr >> 16 & 3;
  return {
    faceIds,
    height: height > 0 && height < 1e5 ? height : 1e3,
    italic: (attr & 1) !== 0,
    bold: (attr >> 1 & 1) !== 0,
    underline: compactStyleFlags && (attr & 1 << 2) !== 0,
    strikeout: compactStyleFlags && (attr >> 18 & 7) !== 0,
    superscript: suType === 1,
    subscript: suType === 2,
    textColor: d.length >= 56 ? colorRef(d, 52) : "000000"
  };
}
var ALIGN_TBL = { 0: "justify", 1: "left", 2: "right", 3: "center", 4: "distribute", 5: "distribute_space" };
function parseParaShape(d) {
  if (d.length < 4) return { align: "justify", spaceBefore: 0, spaceAfter: 0, lineSpacing: 160, lineSpacingType: 0, leftMargin: 0, rightMargin: 0, indent: 0 };
  const attr = BinaryKit.readU32LE(d, 0);
  const legacyLineSpacingType = attr & 3;
  const extendedLineSpacingType = d.length >= 54 ? BinaryKit.readU32LE(d, 46) & 31 : -1;
  const lineSpacingType = extendedLineSpacingType >= 0 && extendedLineSpacingType <= 3 ? extendedLineSpacingType : legacyLineSpacingType;
  const lineSpacing = d.length >= 54 ? BinaryKit.readU32LE(d, 50) : d.length >= 28 ? i32(d, 24) : 160;
  const align = ALIGN_TBL[attr >> 2 & 7] ?? "justify";
  const vVal = attr >> 20 & 3;
  const verAlign = vVal === 1 ? "top" : vVal === 2 ? "center" : vVal === 3 ? "bottom" : "baseline";
  const lineWrap = "break";
  const headingType = attr >>> 23 & 3;
  const headingLevel = attr >>> 25 & 7;
  const heading = headingType === 1 && headingLevel < 6 ? headingLevel + 1 : void 0;
  const listOrd = headingType === 2 ? true : headingType === 3 ? false : void 0;
  const listId = d.length >= 32 ? BinaryKit.readU16LE(d, 30) : 0;
  return {
    align,
    lineSpacingType,
    leftMargin: d.length >= 8 ? i32(d, 4) : 0,
    // offset 4: 문단 몸체 왼쪽 여백 (HWPUNIT * 2)
    rightMargin: d.length >= 12 ? i32(d, 8) : 0,
    // offset 8: 문단 몸체 오른쪽 여백 (HWPUNIT * 2)
    indent: d.length >= 16 ? i32(d, 12) : 0,
    // offset 12: 첫 줄 들여쓰기 (HWPUNIT * 2)
    spaceBefore: d.length >= 20 ? i32(d, 16) : 0,
    spaceAfter: d.length >= 24 ? i32(d, 20) : 0,
    lineSpacing,
    verAlign,
    lineWrap,
    heading,
    listOrd,
    listLevel: listOrd === void 0 ? void 0 : headingLevel,
    listId: listOrd === void 0 ? void 0 : listId
  };
}
var BORDER_W_PT = [0.28, 0.34, 0.43, 0.57, 0.71, 0.85, 1.13, 1.42, 1.7, 1.98, 2.84, 4.25, 5.67, 8.5, 11.34, 14.17];
var BORDER_KIND = { 0: "none", 1: "solid", 2: "dash", 3: "dot", 4: "dash", 5: "dash", 6: "dash", 7: "double", 8: "double", 9: "double", 10: "none" };
function parseBorderFill(d) {
  const borders = [];
  for (let i = 0; i < 4; i++) {
    const off = 2 + i * 6;
    const type = off < d.length ? d[off] : 0;
    const widthPt = off + 1 < d.length ? BORDER_W_PT[d[off + 1]] ?? 0.5 : 0.5;
    const color = off + 6 <= d.length ? colorRef(d, off + 2) : "000000";
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
function parseObjectLayout(data) {
  if (data.length < 28) return void 0;
  const flags = BinaryKit.readU32LE(data, 4);
  if ((flags & 1) !== 0) return { wrap: "inline" };
  const vertRelCode = flags >>> 3 & 3;
  const horzRelCode = flags >>> 8 & 3;
  const vertAlignCode = flags >>> 5 & 7;
  const horzAlignCode = flags >>> 10 & 7;
  const wrapCode = flags >>> 21 & 7;
  const vertRelTo = vertRelCode === 2 ? "para" : "page";
  const horzRelTo = horzRelCode === 2 ? "column" : horzRelCode === 3 ? "para" : "page";
  const vertAlign = ["top", "center", "bottom"][vertAlignCode];
  const horzAlign = ["left", "center", "right"][horzAlignCode];
  const wrap = ["square", "tight", "through", "topAndBottom", "behind", "front"][wrapCode] ?? "square";
  const rawY = i32(data, 8);
  const rawX = i32(data, 12);
  const layout = {
    wrap,
    horzRelTo,
    vertRelTo,
    horzAlign,
    vertAlign,
    xPt: rawX !== 0 ? Metric.hwpToPt(rawX) : void 0,
    yPt: rawY !== 0 ? Metric.hwpToPt(rawY) : void 0,
    behindDoc: wrap === "behind" || void 0,
    zOrder: Math.max(0, i32(data, 24))
  };
  if (data.length >= 36) {
    layout.distL = Metric.hwpToPt(BinaryKit.readU16LE(data, 28));
    layout.distR = Metric.hwpToPt(BinaryKit.readU16LE(data, 30));
    layout.distT = Metric.hwpToPt(BinaryKit.readU16LE(data, 32));
    layout.distB = Metric.hwpToPt(BinaryKit.readU16LE(data, 34));
  }
  return layout;
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
  const _nchars = hdr.data.length >= 4 ? BinaryKit.readU32LE(hdr.data, 0) & 2147483647 : 0;
  const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
  const hwpStyleId = hdr.data.length >= 11 ? hdr.data[10] : void 0;
  const divideSort = hdr.data.length >= 12 ? hdr.data[11] : 0;
  const ps = di.paraShapes[psId];
  let text = null;
  let csPairs = [];
  const grids = [];
  const ctrlHeaders = [];
  let hasSectionCtrl = false;
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
        if (ctrlId === CTRL_SECD) hasSectionCtrl = true;
        if (ctrlId === CTRL_HEAD || ctrlId === CTRL_FOOT) {
          const ctrlLv = r.level;
          const hfParas = [];
          let j = i + 1;
          while (j < recs.length && recs[j].level > ctrlLv) {
            if (recs[j].tag === TAG_PARA_HEADER) {
              const pr = shield.guard(
                () => parseParagraphGroup(recs, j, di, shield, gsoCtx),
                { nodes: [], next: j + 1 },
                `hwp:hf@${j}`
              );
              hfParas.push(...pr.nodes.filter((n) => n.tag === "para"));
              j = pr.next;
            } else {
              j++;
            }
          }
          if (hfParas.length > 0) {
            const key = ctrlId === CTRL_HEAD ? "headers" : "footers";
            if (!gsoCtx[key]) gsoCtx[key] = hfParas;
          }
          i = j;
        } else {
          const MAX_HWP = 1e6;
          const rawW = r.data.length >= 24 ? BinaryKit.readU32LE(r.data, 16) : 0;
          const rawH = r.data.length >= 24 ? BinaryKit.readU32LE(r.data, 20) : 0;
          const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
          const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
          const layout = parseObjectLayout(r.data);
          const atnoType = ctrlId === CTRL_ATNO && r.data.length >= 8 ? BinaryKit.readU32LE(r.data, 4) & 15 : void 0;
          const isPicture = ctrlId === CTRL_GSO || ctrlId === CTRL_PIC;
          const imgId = isPicture ? gsoCtx.count++ : r.data.length >= 6 ? BinaryKit.readU16LE(r.data, 4) : 0;
          const binIndex = isPicture ? pictureBinIndex(recs, i) : void 0;
          ctrlHeaders.push({ ctrlId, imgId, wPt, hPt, layout, atnoType });
          const isImageCtrl = ctrlId === CTRL_IMAGE || ctrlId === CTRL_PIC || ctrlId === CTRL_FIG || ctrlId === CTRL_OBJ || ctrlId === CTRL_GSO;
          if (isImageCtrl) gsoCtx.objects.set(imgId, { wPt, hPt, layout, binIndex });
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
        }
      } else {
        i = skipKids(recs, i);
      }
    } else {
      i++;
    }
  }
  const nodes = [];
  {
    const paraContent = [];
    const atnoCtrls = [];
    if (text && text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        const ch = ctrlHeaders[ci];
        if (ch && ch.ctrlId === CTRL_ATNO)
          atnoCtrls.push({ pos: text.controls[ci].pos, type: ch.atnoType ?? 0 });
      }
      atnoCtrls.sort((a, b) => a.pos - b.pos);
    }
    if (text && text.chars.length > 0) {
      if (atnoCtrls.length > 0) {
        let k = 0;
        for (const ac of atnoCtrls) {
          const seg = [];
          while (k < text.chars.length && text.chars[k].pos < ac.pos) seg.push(text.chars[k++]);
          if (seg.length > 0) paraContent.push(...resolveCharShapes(seg, csPairs, di));
          paraContent.push(buildPageNum(ac.type === 0 ? "decimal" : "total"));
        }
        const rest = text.chars.slice(k);
        if (rest.length > 0) paraContent.push(...resolveCharShapes(rest, csPairs, di));
      } else {
        paraContent.push(...resolveCharShapes(text.chars, csPairs, di));
      }
    } else if (atnoCtrls.length > 0) {
      for (const ac of atnoCtrls) paraContent.push(buildPageNum(ac.type === 0 ? "decimal" : "total"));
    }
    if (text && text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        const ch = ctrlHeaders[ci];
        if (!ch) continue;
        const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_PIC || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
        if (!isImg) continue;
        paraContent.push(buildSpan(`__EXT_${ch.imgId}__`));
      }
    }
    if (divideSort & 4) {
      nodes.push(buildPara([{ tag: "span", props: {}, kids: [buildPb()] }]));
    }
    nodes.push(...grids);
    const isWhitespaceSectionPara = hasSectionCtrl && grids.length === 0 && paraContent.length > 0 && paraContent.every((n) => {
      if (n?.tag !== "span") return false;
      const text2 = (n.kids ?? []).filter((kid) => kid?.tag === "txt").map((kid) => kid.content ?? "").join("");
      return text2.trim() === "";
    });
    const isSectionOnlyPara = hasSectionCtrl && grids.length === 0 && (paraContent.length === 0 || isWhitespaceSectionPara);
    const isPageBreakOnlyPara = divideSort & 4 && paraContent.length === 0 && grids.length === 0;
    if (!isSectionOnlyPara && !isPageBreakOnlyPara) {
      nodes.push(buildPara(
        paraContent.length > 0 ? paraContent : [buildSpan("")],
        buildParaProps(ps, hwpStyleId, di)
      ));
    }
  }
  return { nodes, next: i };
}
function skipKids(recs, idx) {
  const lv = recs[idx].level;
  let i = idx + 1;
  while (i < recs.length && recs[i].level > lv) i++;
  return i;
}
function pictureBinIndex(recs, ctrlIdx) {
  const end = skipKids(recs, ctrlIdx);
  for (let i = ctrlIdx + 1; i < end; i++) {
    const data = recs[i].data;
    if (recs[i].tag === TAG_SHAPE_COMPONENT_PICTURE && data.length >= 73) {
      const binId = BinaryKit.readU16LE(data, 71);
      if (binId > 0) return binId - 1;
    }
  }
  return void 0;
}
var EXT_CTRL = /* @__PURE__ */ new Set([2, 3, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]);
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
      spans.push(...styledSpans(buf, curId, di));
      buf = "";
      curId = sid;
    }
    buf += chars[k].ch;
  }
  if (buf) spans.push(...styledSpans(buf, curId, di));
  return spans;
}
function styledSpans(text, shapeId, di) {
  const cs = di.charShapes[shapeId];
  if (!cs) return [buildSpan(text)];
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
  return splitLeadingSymbolRuns(text, props, di);
}
function splitLeadingSymbolRuns(text, props, di) {
  if (!text) return [buildSpan(text, props)];
  const symbolFont = firstAvailableFont(di, ["\uD55C\uC591\uC2E0\uBA85\uC870", "HY\uC2E0\uBA85\uC870"]) ?? props.font;
  const leadFont = firstAvailableFont(di, ["HCI Poppy"]) ?? symbolFont;
  const out = [];
  let rest = text;
  const lead = rest.match(/^(\s+)([◦→])/);
  if (lead?.[1]) {
    out.push(buildSpan(lead[1], { ...props, b: false, font: leadFont }));
    rest = rest.slice(lead[1].length);
  }
  const marker = rest.match(/^([□◦→])(\s*)/);
  if (marker) {
    out.push(buildSpan(marker[1], { ...props, font: symbolFont }));
    if (marker[2] && marker[1] !== "\u25A1")
      out.push(buildSpan(marker[2], { ...props, b: false, font: leadFont }));
    rest = rest.slice(marker[0].length);
    if (marker[2] && marker[1] === "\u25A1") rest = `${marker[2]}${rest}`;
    if (!marker[2] && rest && (marker[1] === "\u25E6" || marker[1] === "\u2192")) {
      rest = ` ${rest}`;
    }
  }
  if (rest) appendLatinAwareSpans(out, rest, props, leadFont);
  return out.length ? out : [buildSpan(text, props)];
}
function firstAvailableFont(di, names) {
  return names.find((name) => di.faceNames.includes(name));
}
function appendLatinAwareSpans(out, text, props, _latinFont) {
  out.push(buildSpan(text, props));
}
var HWP_DEFAULT_CELL_PADDING = {
  left: 510,
  right: 510,
  top: 141,
  bottom: 141
};
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
  const rowCnt = Math.max(1, tblData.length >= 6 ? BinaryKit.readU16LE(tblData, 4) : 1);
  const colCnt = Math.max(1, tblData.length >= 8 ? BinaryKit.readU16LE(tblData, 6) : 1);
  const tablePadding = tblData.length >= 18 ? {
    left: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 10), HWP_DEFAULT_CELL_PADDING.left),
    right: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 12), HWP_DEFAULT_CELL_PADDING.right),
    top: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 14), HWP_DEFAULT_CELL_PADDING.top),
    bottom: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 16), HWP_DEFAULT_CELL_PADDING.bottom)
  } : HWP_DEFAULT_CELL_PADDING;
  const parsed = [];
  for (let ci = 0; ci < cells.length; ci++) {
    const c = cells[ci];
    const seqIdx = ci;
    const pc = shield.guard(
      () => parseCellRec(c.data, c.tag, recs, c.cStart, c.cEnd, di, shield, seqIdx, colCnt, gsoCtx, tablePadding),
      { row: Math.floor(ci / (colCnt || 1)), col: ci % (colCnt || 1), cs: 1, rs: 1, widthHwp: 0, heightHwp: void 0, props: {}, cellChildren: [buildPara([buildSpan("")])] },
      `hwp:cell@${c.cStart}`
    );
    parsed.push(pc);
  }
  const rowLimit = Math.max(rowCnt, Math.ceil(parsed.length / colCnt), 1);
  for (let idx = 0; idx < parsed.length; idx++) {
    const c = parsed[idx];
    const badPosition = !Number.isFinite(c.row) || !Number.isFinite(c.col) || c.row < 0 || c.col < 0 || c.col >= colCnt || c.row > rowLimit * 4 + 20;
    if (badPosition) {
      c.row = Math.floor(idx / colCnt);
      c.col = idx % colCnt;
    }
    const maxColSpan = Math.max(1, colCnt - c.col);
    if (!Number.isFinite(c.cs) || c.cs < 1) c.cs = 1;
    if (c.cs > maxColSpan) c.cs = maxColSpan;
    const maxRowSpan = Math.max(1, rowCnt - Math.min(c.row, rowCnt - 1));
    if (!Number.isFinite(c.rs) || c.rs < 1) c.rs = 1;
    if (c.rs > maxRowSpan) c.rs = maxRowSpan;
  }
  const maxRow = parsed.reduce((m, c) => Math.max(m, c.row + c.rs), 0);
  const actualRowCnt = Math.max(rowCnt, maxRow);
  const colWidthsPt = inferColumnWidths(
    colCnt,
    parsed.filter((c) => c.widthHwp > 0).map((c) => ({ start: c.col, span: c.cs, width: c.widthHwp }))
  ).map(Metric.hwpToPt);
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
  gp.cellPadL = Metric.hwpToPt(tablePadding.left);
  gp.cellPadR = Metric.hwpToPt(tablePadding.right);
  gp.cellPadT = Metric.hwpToPt(tablePadding.top);
  gp.cellPadB = Metric.hwpToPt(tablePadding.bottom);
  const hasWidths = colWidthsPt.some((w) => w > 0);
  if (hasWidths) gp.colWidths = colWidthsPt;
  const tableLayout = parseObjectLayout(recs[ctrlIdx].data);
  if (tableLayout && tableLayout.wrap !== "inline") gp.layout = tableLayout;
  return { grid: buildGrid(rows, gp), next: i };
}
function parseCellRec(d, tag, recs, cStart, cEnd, di, shield, seqIdx, colCnt, gsoCtx, tablePadding) {
  let col, row, cs = 1, rs = 1;
  let widthHwp = 0;
  let heightHwp = 0;
  const props = {};
  const attr = tag === TAG_LIST_HEADER ? d.length >= 8 ? BinaryKit.readU32LE(d, 4) : 0 : d.length >= 6 ? BinaryKit.readU32LE(d, 2) : 0;
  const va = attr >> 5 & 3;
  if (va === 1) props.va = "mid";
  else if (va === 2) props.va = "bot";
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
      if (isCellPaddingOverride(pL, tablePadding.left)) props.padL = Metric.hwpToPt(pL);
      if (isCellPaddingOverride(pR, tablePadding.right)) props.padR = Metric.hwpToPt(pR);
      if (isCellPaddingOverride(pT, tablePadding.top)) props.padT = Metric.hwpToPt(pT);
      if (isCellPaddingOverride(pB, tablePadding.bottom)) props.padB = Metric.hwpToPt(pB);
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
      if (isCellPaddingOverride(pL, tablePadding.left)) props.padL = Metric.hwpToPt(pL);
      if (isCellPaddingOverride(pR, tablePadding.right)) props.padR = Metric.hwpToPt(pR);
      if (isCellPaddingOverride(pT, tablePadding.top)) props.padT = Metric.hwpToPt(pT);
      if (isCellPaddingOverride(pB, tablePadding.bottom)) props.padB = Metric.hwpToPt(pB);
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
          const cellStyleId = hdr.data.length >= 11 ? hdr.data[10] : 0;
          const cellDivide = hdr.data.length >= 12 ? hdr.data[11] : 0;
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
                  const rawH = recs[j].data.length >= 24 ? BinaryKit.readU32LE(recs[j].data, 20) : 0;
                  const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
                  const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
                  const layout = parseObjectLayout(recs[j].data);
                  const isPicture = ctrlId === CTRL_GSO || ctrlId === CTRL_PIC;
                  const imgId = isPicture ? gsoCtx.count++ : recs[j].data.length >= 6 ? BinaryKit.readU16LE(recs[j].data, 4) : 0;
                  const binIndex = isPicture ? pictureBinIndex(recs, j) : void 0;
                  ctrlHdrs.push({ ctrlId, imgId, wPt, hPt, layout });
                  const isImageCtrl = ctrlId === CTRL_IMAGE || ctrlId === CTRL_PIC || ctrlId === CTRL_FIG || ctrlId === CTRL_OBJ || ctrlId === CTRL_GSO;
                  if (isImageCtrl) gsoCtx.objects.set(imgId, { wPt, hPt, layout, binIndex });
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
              const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_PIC || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
              if (!isImg) continue;
              paraContent.push(buildSpan(`__EXT_${ch.imgId}__`));
            }
          }
          const kids = paraContent.length > 0 ? paraContent : [buildSpan("")];
          const isPageBreakOnlyPara = cellDivide & 4 && paraContent.length === 0 && innerGrids.length === 0;
          const items = [...innerGrids];
          if (!isPageBreakOnlyPara) {
            items.push(
              buildPara(
                kids,
                buildParaProps(ps, cellStyleId, di)
              )
            );
          }
          if (cellDivide & 4) items.unshift(buildPara([{ tag: "span", props: {}, kids: [buildPb()] }]));
          return { items, next: j };
        },
        { items: [buildPara([buildSpan("")])], next: k + 1 },
        `hwp:cellP@${k}`
      );
      cellChildren.push(...r.items);
      k = r.next;
    } else if (recs[k].tag === TAG_CTRL_HEADER && recs[k].data.length >= 4) {
      const cellCtrlId = BinaryKit.readU32LE(recs[k].data, 0);
      if (cellCtrlId === CTRL_GSO || cellCtrlId === CTRL_PIC) {
        const gsoId = gsoCtx.count++;
        const binIndex = pictureBinIndex(recs, k);
        const rawW = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 16) : 0;
        const rawH = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 20) : 0;
        const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
        const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
        const layout = parseObjectLayout(recs[k].data);
        gsoCtx.objects.set(gsoId, { wPt, hPt, layout, binIndex });
        cellChildren.push(buildPara([buildSpan(`__EXT_${gsoId}__`)]));
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
function inheritedHwpPadding(value, fallback) {
  return value === 65535 ? fallback : value;
}
function isCellPaddingOverride(value, inherited) {
  return value !== 65535 && value !== inherited;
}
function parsePageDef(d) {
  if (d.length < 24) return A4;
  const w = BinaryKit.readU32LE(d, 0);
  const h = BinaryKit.readU32LE(d, 4);
  const ml = BinaryKit.readU32LE(d, 8);
  const mr = BinaryKit.readU32LE(d, 12);
  const mt = BinaryKit.readU32LE(d, 16);
  const mb = BinaryKit.readU32LE(d, 20);
  const header = d.length >= 28 ? BinaryKit.readU32LE(d, 24) : 0;
  const footer = d.length >= 32 ? BinaryKit.readU32LE(d, 28) : 0;
  const at = d.length >= 40 ? BinaryKit.readU32LE(d, 36) : 0;
  return {
    wPt: Metric.hwpToPt(w),
    hPt: Metric.hwpToPt(h),
    ml: Metric.hwpToPt(ml),
    mr: Metric.hwpToPt(mr),
    mt: Metric.hwpToPt(mt),
    mb: Metric.hwpToPt(mb),
    headerPt: Metric.hwpToPt(header),
    footerPt: Metric.hwpToPt(footer),
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
function headingFromStyle(style) {
  if (!style) return void 0;
  for (const name of [style.name, style.engName]) {
    const match = name.match(/^(?:개요|outline|heading)\s*([1-6])$/i);
    if (match) return Number(match[1]);
  }
  return void 0;
}
function buildParaProps(ps, hwpStyleId, di) {
  const p = hwpStyleId !== void 0 ? { hwpStyleId } : {};
  const heading = ps?.heading ?? headingFromStyle(di?.styles[hwpStyleId ?? -1]);
  if (heading !== void 0) p.heading = heading;
  if (!ps) return { ...p, spaceBefore: 0, spaceAfter: 0, lineHeight: 1.6 };
  if (ps.listOrd !== void 0) {
    p.listOrd = ps.listOrd;
    p.listLv = Math.max(0, Math.min(6, ps.listLevel ?? 0));
    if (ps.listOrd) {
      p.listMark = "1.";
    } else {
      const character = ps.listId && di ? di.bullets[ps.listId - 1]?.character : void 0;
      p.listMark = character || "-";
    }
  }
  if (ps.align && ps.align !== "justify") p.align = ps.align;
  if (hwpStyleId === 18 && !p.align) p.align = "justify";
  p.spaceBefore = Math.max(0, Metric.hwpToPt(ps.spaceBefore / 2));
  p.spaceAfter = Math.max(0, Metric.hwpToPt(ps.spaceAfter / 2));
  if (ps.lineSpacingType === 1 || ps.lineSpacingType === 3) {
    if (ps.lineSpacing > 0) {
      p.lineHeightFixed = Metric.hwpToPt(ps.lineSpacing / 2);
      p.lineHeightRule = ps.lineSpacingType === 3 ? "atLeast" : "exact";
    }
  } else {
    p.lineHeight = ps.lineSpacing > 0 ? ps.lineSpacing / 100 : 1.6;
  }
  const leftMarginPt = Math.max(0, Metric.hwpToPt(ps.leftMargin / 2));
  if (leftMarginPt > 0) p.indentPt = leftMarginPt;
  const rightMarginPt = Math.max(0, Metric.hwpToPt(ps.rightMargin / 2));
  if (rightMarginPt > 0) p.indentRightPt = rightMarginPt;
  if (ps.indent !== 0) p.firstLineIndentPt = Metric.hwpToPt(ps.indent / 2);
  if (ps.verAlign && ps.verAlign !== "baseline") p.verAlign = ps.verAlign;
  if (ps.lineWrap && ps.lineWrap !== "break") p.lineWrap = ps.lineWrap;
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
      let di = {
        faceNames: [],
        charShapes: [],
        paraShapes: [],
        borderFills: [],
        styles: [],
        numberings: [],
        bullets: []
      };
      if (diRaw) {
        di = shield.guard(() => parseDocInfo(diRaw, compressed), di, "hwp:docInfo");
      }
      const binEntries = [];
      for (const [path, streamData] of streams) {
        const m = path.match(/^BinData[/\\]BIN([0-9a-f]+)\.([a-z0-9]+)$/i);
        if (m) binEntries.push({ binNum: parseInt(m[1], 16), ext: m[2].toLowerCase(), data: streamData });
      }
      binEntries.sort((a, b) => a.binNum - b.binNum);
      const objectMap = /* @__PURE__ */ new Map();
      for (const { binNum, ext, data: storedData } of binEntries) {
        let imgData = storedData;
        try {
          const inflated = pako2.inflateRaw(storedData);
          if (looksLikeImageData(inflated, ext)) imgData = inflated;
        } catch {
        }
        let mimeType = "image/jpeg";
        if (imgData[0] === 137 && imgData[1] === 80) mimeType = "image/png";
        else if (imgData[0] === 71 && imgData[1] === 73) mimeType = "image/gif";
        else if (imgData[0] === 66 && imgData[1] === 77) mimeType = "image/bmp";
        else if (ext === "wmf") mimeType = "image/x-wmf";
        else if (ext === "emf") mimeType = "image/x-emf";
        const base64 = TextKit.base64Encode(imgData);
        const { wPt, hPt } = getImageDimsPt(imgData, mimeType);
        objectMap.set(binNum - 1, buildImg(base64, mimeType, wPt, hPt));
      }
      const gsoCtx = { count: 0, objects: /* @__PURE__ */ new Map() };
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
        injectImagesIntoContent(allContent, objectMap, gsoCtx.objects);
      }
      normalizeHancomParagraphAnchors(allContent, di);
      warns.push(...shield.flush());
      const content = allContent.length > 0 ? allContent : [buildPara([buildSpan("")])];
      return succeed(buildRoot({}, [buildSheet(content, pageDims, {
        headers: gsoCtx.headers ? { default: gsoCtx.headers } : void 0,
        footers: gsoCtx.footers ? { default: gsoCtx.footers } : void 0
      })]), warns);
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
function looksLikeImageData(data, ext) {
  if (data.length < 4) return false;
  if (data[0] === 137 && data[1] === 80 && data[2] === 78 && data[3] === 71) return true;
  if (data[0] === 255 && data[1] === 216 && data[2] === 255) return true;
  if (data[0] === 71 && data[1] === 73 && data[2] === 70) return true;
  if (data[0] === 66 && data[1] === 77) return true;
  if (ext === "wmf") {
    return data[0] === 215 && data[1] === 205 && data[2] === 198 && data[3] === 154 || data[0] === 1 && data[1] === 0 && data[2] === 9 && data[3] === 0;
  }
  return ext === "emf" && data.length >= 44 && data[40] === 32 && data[41] === 69 && data[42] === 77 && data[43] === 70;
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
function injectImagesIntoContent(content, objectMap, objectInfo) {
  if (objectMap.size === 0) return;
  const processKids = (kids) => {
    for (let i = 0; i < kids.length; i++) {
      const kid = kids[i];
      if (kid.tag === "span" && kid.kids && kid.kids[0]?.tag === "txt") {
        const text = kid.kids[0].content;
        const match = text.match?.(/^__(?:IMG|EXT)_(\d+)(?:_W(\d+)_H(\d+))?__$/);
        if (match) {
          const objId = parseInt(match[1], 10);
          const info = objectInfo.get(objId);
          const base = objectMap.get(info?.binIndex ?? objId);
          if (base) {
            const wPt = match[2] ? parseInt(match[2], 10) : 0;
            const hPt = match[3] ? parseInt(match[3], 10) : 0;
            kids[i] = {
              ...base,
              w: info?.wPt && info.wPt > 0 ? info.wPt : wPt > 0 ? wPt : base.w,
              h: info?.hPt && info.hPt > 0 ? info.hPt : hPt > 0 ? hPt : base.h,
              layout: info?.layout
            };
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
function normalizeHancomParagraphAnchors(content, di) {
  normalizeContentList(content, di);
}
function normalizeContentList(content, di) {
  for (const node of content) {
    if (node?.tag === "grid") {
      for (const row of node.kids ?? []) {
        for (const cell of row.kids ?? []) normalizeContentList(cell.kids ?? [], di);
      }
    }
  }
  for (let i = 0; i < content.length; i++) {
    const node = content[i];
    if (isEmptyCenterPara(node) && paraText(content[i + 1]).startsWith("\u203B \uBAA8\uB4E0 \uC11C\uB958")) {
      content.splice(i, 1);
      i--;
      continue;
    }
    if (paraText(node).startsWith("\uC81C\uCD9C\uC608\uC2DC)") && !isEmptyCenterPara(content[i - 1])) {
      const font = firstAvailableFont(di, ["HCI Poppy"]);
      content.splice(i, 0, buildPara([buildSpan("", font ? { font, pt: 13 } : {})], { hwpStyleId: 0, align: "center" }));
      i++;
    }
  }
}
function isEmptyCenterPara(node) {
  return !!node && node.tag === "para" && paraText(node) === "" && node.props.align === "center";
}
function paraText(node) {
  if (!node || node.tag !== "para") return "";
  let out = "";
  const collect = (kids) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "txt") out += kid.content ?? "";
      else if (kid.kids) collect(kid.kids);
    }
  };
  collect(node.kids);
  return out.trim();
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
          numMap = await parseNumbering2(TextKit.decode(numXml));
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
var DOCX_DEFAULT_STYLE_KEY = "__docx_defaults__";
var WORD_DEFAULT_SPACE_BEFORE_PT = 0;
var WORD_DEFAULT_SPACE_AFTER_PT = 8;
var WORD_DEFAULT_LINE_HEIGHT = 1.15;
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
async function parseNumbering2(xml) {
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
        ctx.warns.push(
          `[DocxDecoder] ${kind} (${type}) XML \uD30C\uC2F1 \uC2E4\uD328: ${err?.message ?? String(err)}`
        );
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
function decodeCellKids(cell, ctx) {
  const elements = getBodyElements(cell);
  const kids = [];
  for (const el of elements) {
    const decoded = ctx.shield.guard(
      () => decodeElement(el, ctx),
      [],
      "docx:cellElement"
    );
    const nodes = Array.isArray(decoded) ? decoded : [decoded];
    for (const node of nodes) {
      if (node && (node.tag === "para" || node.tag === "grid")) {
        kids.push(node);
      }
    }
  }
  return kids;
}
function decodePara2(p, ctx) {
  const pPr = p?.["w:pPr"]?.[0] ?? {};
  const alignVal = pPr?.["w:jc"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:jc"]?.[0]?._attr?.val;
  const headStyle = pPr?.["w:pStyle"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:pStyle"]?.[0]?._attr?.val ?? "";
  const documentDefaults = resolveParaStyle(
    DOCX_DEFAULT_STYLE_KEY,
    ctx.paraStyleMap
  );
  const namedStyle = resolveParaStyle(
    headStyle || void 0,
    ctx.paraStyleMap
  );
  const styleInherited = {
    rPr: { ...documentDefaults.rPr, ...namedStyle.rPr },
    pPr: { ...documentDefaults.pPr, ...namedStyle.pPr }
  };
  const canonicalStyle = canonicalDocxStyleId(headStyle, ctx.paraStyleMap);
  const props = {
    align: safeAlign(alignVal),
    heading: parseHeading(headStyle),
    styleId: canonicalStyle
  };
  const spacingAttr = pPr?.["w:spacing"]?.[0]?._attr ?? pPr?.spacing?.[0]?._attr ?? {};
  const beforeRaw = docxAttr(spacingAttr, "before");
  const afterRaw = docxAttr(spacingAttr, "after");
  const lineRaw = docxAttr(spacingAttr, "line");
  const beforeVal = Number(beforeRaw);
  const afterVal = Number(afterRaw);
  const lineVal = Number(lineRaw);
  const lineRule = docxAttr(spacingAttr, "lineRule") ?? "auto";
  props.spaceBefore = beforeRaw !== void 0 && Number.isFinite(beforeVal) ? Metric.dxaToPt(Math.max(0, beforeVal)) : styleInherited.pPr?.spaceBefore ?? WORD_DEFAULT_SPACE_BEFORE_PT;
  props.spaceAfter = afterRaw !== void 0 && Number.isFinite(afterVal) ? Metric.dxaToPt(Math.max(0, afterVal)) : styleInherited.pPr?.spaceAfter ?? WORD_DEFAULT_SPACE_AFTER_PT;
  if (lineRaw !== void 0 && Number.isFinite(lineVal) && lineVal > 0) {
    if (lineRule === "exact" || lineRule === "atLeast") {
      props.lineHeightFixed = Metric.dxaToPt(lineVal);
      props.lineHeightRule = lineRule;
    } else {
      props.lineHeight = lineVal / 240;
    }
  } else if (styleInherited.pPr?.lineHeightFixed !== void 0) {
    props.lineHeightFixed = styleInherited.pPr.lineHeightFixed;
    props.lineHeightRule = styleInherited.pPr.lineHeightRule;
  } else {
    props.lineHeight = styleInherited.pPr?.lineHeight ?? WORD_DEFAULT_LINE_HEIGHT;
  }
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
      ctx.warns.push(`[DocxDecoder] image not found: "${target}"`);
      return null;
    }
    const ext = target.split(".").pop()?.toLowerCase() ?? "png";
    const mimeMap = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
      wmf: "image/x-wmf",
      emf: "image/x-emf"
    };
    const mime = mimeMap[ext] ?? "image/png";
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
function docxAttr(attrs, name) {
  if (!attrs) return void 0;
  return attrs[`w:${name}`] ?? attrs[name];
}
function parseDocxSpacingProps(pPr, includeWordDefaults = false) {
  const parsed = includeWordDefaults ? {
    spaceBefore: WORD_DEFAULT_SPACE_BEFORE_PT,
    spaceAfter: WORD_DEFAULT_SPACE_AFTER_PT,
    lineHeight: WORD_DEFAULT_LINE_HEIGHT
  } : {};
  const spacingAttr = pPr?.["w:spacing"]?.[0]?._attr ?? pPr?.spacing?.[0]?._attr;
  if (!spacingAttr) return parsed;
  const beforeRaw = docxAttr(spacingAttr, "before");
  const afterRaw = docxAttr(spacingAttr, "after");
  const lineRaw = docxAttr(spacingAttr, "line");
  const lineRule = docxAttr(spacingAttr, "lineRule") ?? "auto";
  const beforeVal = Number(beforeRaw);
  const afterVal = Number(afterRaw);
  const lineVal = Number(lineRaw);
  if (beforeRaw !== void 0 && Number.isFinite(beforeVal)) {
    parsed.spaceBefore = Metric.dxaToPt(Math.max(0, beforeVal));
  }
  if (afterRaw !== void 0 && Number.isFinite(afterVal)) {
    parsed.spaceAfter = Metric.dxaToPt(Math.max(0, afterVal));
  }
  if (lineRaw !== void 0 && Number.isFinite(lineVal) && lineVal > 0) {
    if (lineRule === "exact" || lineRule === "atLeast") {
      parsed.lineHeight = void 0;
      parsed.lineHeightFixed = Metric.dxaToPt(lineVal);
      parsed.lineHeightRule = lineRule;
    } else {
      parsed.lineHeight = lineVal / 240;
      parsed.lineHeightFixed = void 0;
      parsed.lineHeightRule = void 0;
    }
  }
  return parsed;
}
async function parseParaStyleMap(xml) {
  const map = /* @__PURE__ */ new Map();
  const trimmed = xml.trim();
  if (!trimmed) return map;
  try {
    const obj = await XmlKit.parseStrict(trimmed);
    const stylesRoot = obj?.["w:styles"]?.[0] ?? obj?.styles?.[0] ?? obj;
    const defaultsPPr = stylesRoot?.["w:docDefaults"]?.[0]?.["w:pPrDefault"]?.[0]?.["w:pPr"]?.[0] ?? stylesRoot?.docDefaults?.[0]?.pPrDefault?.[0]?.pPr?.[0];
    map.set(DOCX_DEFAULT_STYLE_KEY, {
      pPr: parseDocxSpacingProps(defaultsPPr, true)
    });
    const styleArr = toArr2(stylesRoot?.["w:style"] ?? stylesRoot?.style);
    for (const style of styleArr) {
      const attr = style?._attr ?? {};
      const type = attr?.["w:type"] ?? attr?.type;
      if (type !== "paragraph" && type !== "character") continue;
      const id = attr?.["w:styleId"] ?? attr?.styleId;
      if (!id) continue;
      const basedOn = (style?.["w:basedOn"]?.[0]?._attr ?? style?.basedOn?.[0]?._attr)?.["w:val"];
      const nameAttr = style?.["w:name"]?.[0]?._attr ?? style?.name?.[0]?._attr;
      const name = nameAttr?.["w:val"] ?? nameAttr?.val;
      const def = { basedOn, name };
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
        const spacingProps = parseDocxSpacingProps(pPr);
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
          ...spacingProps,
          align: alignVal,
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
function canonicalDocxStyleId(styleId, map) {
  if (!styleId) return void 0;
  const styleName = map.get(styleId)?.name;
  if (styleName === "\uBC14\uD0D5\uAE00") return "0";
  return styleId;
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
  const tblCellMar = tblPr?.["w:tblCellMar"]?.[0] ?? tblPr?.tblCellMar?.[0];
  if (tblCellMar) {
    const readMarginPt = (side) => {
      const attrs = tblCellMar?.[`w:${side}`]?.[0]?._attr ?? tblCellMar?.[side]?.[0]?._attr;
      const value = Number(docxAttr(attrs, "w"));
      return Number.isFinite(value) && value >= 0 ? Metric.dxaToPt(value) : void 0;
    };
    gridProps.cellPadT = readMarginPt("top");
    gridProps.cellPadB = readMarginPt("bottom");
    gridProps.cellPadL = readMarginPt("left");
    gridProps.cellPadR = readMarginPt("right");
  }
  const layout = decodeFloatingTableLayout(tblPr);
  if (layout) gridProps.layout = layout;
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
      const paras = decodeCellKids(cell, ctx);
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
function decodeFloatingTableLayout(tblPr) {
  const tblpPr = tblPr?.["w:tblpPr"]?.[0] ?? tblPr?.tblpPr?.[0];
  const attr = tblpPr?._attr;
  if (!attr) return void 0;
  const get = (name) => attr[`w:${name}`] ?? attr[name];
  const horzRelMap = {
    margin: "margin",
    page: "page",
    text: "para"
  };
  const vertRelMap = {
    margin: "margin",
    page: "page",
    text: "para"
  };
  const horzAlignMap = {
    left: "left",
    center: "center",
    right: "right"
  };
  const vertAlignMap = {
    top: "top",
    center: "center",
    bottom: "bottom"
  };
  const numberDxa = (name) => {
    const raw = get(name);
    if (raw === void 0) return void 0;
    const value = Number(raw);
    return Number.isFinite(value) ? Metric.dxaToPt(value) : void 0;
  };
  const layout = {
    wrap: "topAndBottom",
    horzRelTo: horzRelMap[get("horzAnchor") ?? ""] ?? "para",
    vertRelTo: vertRelMap[get("vertAnchor") ?? ""] ?? "para",
    distL: numberDxa("leftFromText"),
    distR: numberDxa("rightFromText"),
    distT: numberDxa("topFromText"),
    distB: numberDxa("bottomFromText")
  };
  const xSpec = get("tblpXSpec");
  const ySpec = get("tblpYSpec");
  const x = numberDxa("tblpX");
  const y = numberDxa("tblpY");
  if (x !== void 0 && !xSpec) layout.xPt = x;
  else if (xSpec) layout.horzAlign = horzAlignMap[xSpec];
  if (y !== void 0 && !ySpec) layout.yPt = y;
  else if (ySpec) layout.vertAlign = vertAlignMap[ySpec];
  return layout;
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
            listOrd: /\d+\./.test(listMatch[2]),
            listMark: listMatch[2]
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
  const parse = (line) => {
    const cells = [];
    let cell = "";
    for (let i = 0; i < line.length; i++) {
      if (line[i] === "\\" && line[i + 1] === "|") {
        cell += "|";
        i++;
      } else if (line[i] === "|") {
        cells.push(cell.trim());
        cell = "";
      } else {
        cell += line[i];
      }
    }
    cells.push(cell.trim());
    if (cells[0] === "") cells.shift();
    if (cells[cells.length - 1] === "") cells.pop();
    return cells;
  };
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
  'xmlns:opf="http://www.idpf.org/2007/opf/"',
  'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"',
  'xmlns:epub="http://www.idpf.org/2007/ops"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
  'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"'
].join(" ");
var LINESEG_FLAGS = 393216;
var LINESEG_FLAG_INDENT = 1048576;
var LINESEG_FLAG_PAGE_FIRST = 1;
var LINESEG_FLAG_COLUMN_FIRST = 2;
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
  /** Register a face in every language bank and return bank-local IDs. */
  registerFont(rawFace) {
    const face = safeFontToKr(rawFace) || "\uD568\uCD08\uB86C\uBC14\uD0D5";
    const isKor = this.isKorean(face);
    const ids = {};
    for (const group of LANG_GROUPS) {
      const useFace = group === "LATIN" ? isKor ? "\uD568\uCD08\uB86C\uBC14\uD0D5" : face : isKor ? face : "\uD568\uCD08\uB86C\uBC14\uD0D5";
      ids[group] = this.register(group, useFace);
    }
    return ids;
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
function quantizeBorderWidth(pt) {
  const mm = pt * 0.3528;
  const standardWidths = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1, 1.5, 2, 3, 4, 5];
  let closest = standardWidths[0];
  let minDiff = Math.abs(mm - closest);
  for (let i = 1; i < standardWidths.length; i++) {
    const diff = Math.abs(mm - standardWidths[i]);
    if (diff < minDiff) {
      minDiff = diff;
      closest = standardWidths[i];
    }
  }
  let str = closest.toFixed(2);
  if (str.endsWith("0")) {
    str = str.slice(0, -1);
  }
  if (str.endsWith(".0")) {
    str = str.slice(0, -2);
  }
  return `${str} mm`;
}
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
    const w = s && s.kind !== "none" ? quantizeBorderWidth(s.pt) : "0.12 mm";
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
function paraShapeHwpToLayoutHwp(value) {
  return Math.round(value / 2);
}
function paraPrKey(p) {
  return `${p.align ?? "left"}|${p.verAlign ?? "baseline"}|${p.lineWrap ?? "break"}|${p.listOrd ?? ""}|${p.listLv ?? 0}|${p.indentPt ?? 0}|${p.indentRightPt ?? 0}|${p.firstLineIndentPt ?? 0}|${p.spaceBefore ?? 0}|${p.spaceAfter ?? 0}|${p.lineHeight ?? 0}|${p.lineHeightFixed ?? 0}|${p.styleId ?? ""}`;
}
function registerCharPr(props, ctx) {
  const key = charPrKey(props);
  const existing = ctx.charPrMap.get(key);
  if (existing !== void 0) return existing;
  const rawFont = props.font ?? "\uD568\uCD08\uB86C\uBC14\uD0D5";
  const fontIds = ctx.fontBank.registerFont(rawFont);
  const id = ctx.charPrs.length;
  ctx.charPrs.push({
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? "BOTTOM" : "NONE",
    strikeout: props.s ? "SOLID" : "NONE",
    textColor: props.color ? `#${props.color}` : "#000000",
    hangulId: fontIds.HANGUL,
    latinId: fontIds.LATIN,
    hanjaId: fontIds.HANJA,
    japaneseId: fontIds.JAPANESE,
    otherId: fontIds.OTHER,
    symbolId: fontIds.SYMBOL,
    userId: fontIds.USER,
    bg: props.bg
  });
  ctx.charPrMap.set(key, id);
  return id;
}
var ALIGN_MAP2 = {
  left: "LEFT",
  center: "CENTER",
  right: "RIGHT",
  justify: "JUSTIFY",
  distribute: "DISTRIBUTE",
  distribute_space: "DISTRIBUTE_SPACE"
};
var V_ALIGN_MAP = {
  baseline: "BASELINE",
  top: "TOP",
  center: "CENTER",
  bottom: "BOTTOM"
};
var LINE_WRAP_MAP = {
  break: "BREAK",
  squeeze: "SQUEEZE",
  keep: "KEEP"
};
function registerParaPr(props, ctx) {
  const key = paraPrKey(props);
  const existing = ctx.paraPrMap.get(key);
  if (existing !== void 0) return existing;
  const id = ctx.paraPrs.length;
  const alignStr = props.align ? ALIGN_MAP2[props.align] ?? "LEFT" : "LEFT";
  const verAlignStr = props.verAlign ? V_ALIGN_MAP[props.verAlign] ?? "BASELINE" : "BASELINE";
  const lineWrapStr = props.lineWrap ? LINE_WRAP_MAP[props.lineWrap] ?? "BREAK" : "BREAK";
  const def = {
    id,
    align: alignStr,
    verAlign: verAlignStr,
    lineWrap: lineWrapStr,
    leftHwp: Metric.ptToHwp(props.indentPt ?? 0) * 2,
    rightHwp: Metric.ptToHwp(props.indentRightPt ?? 0) * 2,
    intentHwp: Metric.ptToHwp(props.firstLineIndentPt ?? 0) * 2,
    prevHwp: Metric.ptToHwp(props.spaceBefore ?? 0) * 2,
    nextHwp: Metric.ptToHwp(props.spaceAfter ?? 0) * 2,
    lineSpacing: props.lineHeightFixed ? 0 : props.lineHeight ? Math.round(props.lineHeight * 100) : 160,
    lineSpacingFixed: props.lineHeightFixed ? Math.max(
      Metric.ptToHwp(props.lineHeightFixed),
      Math.ceil(1e3 * 1.15)
    ) * 2 : void 0
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
  if (mime.includes("wmf")) return "wmf";
  if (mime.includes("emf")) return "emf";
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
  if (styleId === "Normal" || styleId === "0") {
    ctx.styleIdToHwpxId.set(styleId, 0);
    return;
  }
  const usedIds = new Set(ctx.hwpxStyles.map((s) => s.id));
  const numericId = Number(styleId);
  let hwpxId = Number.isInteger(numericId) && numericId > 0 && !usedIds.has(numericId) ? numericId : nextStyleId(usedIds);
  ctx.styleIdToHwpxId.set(styleId, hwpxId);
  ctx.hwpxStyles.push({
    id: hwpxId,
    name: STYLE_NAME_MAP[styleId] ?? styleId,
    engName: "",
    paraPrIDRef: paraPrId,
    charPrIDRef: charPrId
  });
}
function nextStyleId(usedIds) {
  let id = 0;
  while (usedIds.has(id)) id++;
  return id;
}
function materializeContiguousStyles(styles) {
  const byId = new Map(styles.map((style) => [style.id, style]));
  const maxId = Math.max(0, ...byId.keys());
  const dense = [];
  for (let id = 0; id <= maxId; id++) {
    dense.push(byId.get(id) ?? {
      id,
      name: `\uC0AC\uC6A9\uC790 \uC2A4\uD0C0\uC77C ${id}`,
      engName: `User Style ${id}`,
      paraPrIDRef: 0,
      charPrIDRef: 0
    });
  }
  return dense;
}
function paraStyleKey(props) {
  if (props.hwpStyleId !== void 0) {
    const id = Math.trunc(props.hwpStyleId);
    if (id >= 0 && id <= 255) return String(id);
  }
  return props.styleId;
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
  const styleKey = paraStyleKey(para.props);
  if (styleKey) registerStyle(styleKey, paraPrId, firstCharPrId, ctx);
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
      const safeML = dims.ml !== void 0 && dims.ml >= 0 ? dims.ml : 70.87;
      const safeMR = dims.mr !== void 0 && dims.mr >= 0 ? dims.mr : 70.87;
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
          name: "META-INF/manifest.xml",
          data: this.stringToBytes(MANIFEST_XML),
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
        const ct = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "wmf" ? "image/x-wmf" : ext === "emf" ? "image/x-emf" : "image/bmp";
        entries.push({ name: `BinData/${bin.name}`, data: bin.data, mime: ct });
      }
      return succeed(await this.zip(entries));
    } catch (e) {
      return fail(`HWPX \uC778\uCF54\uB529 \uC624\uB958: ${e?.message ?? String(e)}`);
    }
  }
};
var VERSION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 8227 WIN32LEWindows_10"/>`;
var CONTAINER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>`;
var MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`;
var CONTAINER_RDF = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><pkg:hasPart xmlns:pkg="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><pkg:hasPart xmlns:pkg="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>`;
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
    const ct = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "wmf" ? "image/x-wmf" : ext === "emf" ? "image/x-emf" : "image/bmp";
    items += `<opf:item id="${bin.id}" href="BinData/${bin.name}" media-type="${ct}" isEmbeded="1"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title>${title}</opf:title><opf:language>ko</opf:language><opf:meta name="creator"      content="text">${creator}</opf:meta><opf:meta name="subject"      content="text">${subject}</opf:meta><opf:meta name="description"  content="text">${desc}</opf:meta><opf:meta name="CreatedDate"  content="text">${created}</opf:meta><opf:meta name="ModifiedDate" content="text">${modified}</opf:meta><opf:meta name="keyword"      content="text">${keyword}</opf:meta><opf:meta name="trackchageConfig" content="text">0</opf:meta></opf:metadata><opf:manifest>${items}</opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`;
}
function buildSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/><config:config-item-set name="PrintInfo"><config:config-item name="PrintAutoFootNote" type="boolean">false</config:config-item><config:config-item name="PrintAutoHeadNote" type="boolean">false</config:config-item><config:config-item name="PrintMethod" type="short">4</config:config-item><config:config-item name="OverlapSize" type="short">0</config:config-item><config:config-item name="PrintCropMark" type="short">0</config:config-item><config:config-item name="BinderHoleType" type="short">0</config:config-item><config:config-item name="ZoomX" type="short">100</config:config-item><config:config-item name="ZoomY" type="short">100</config:config-item></config:config-item-set></ha:HWPApplicationSetting>`;
}
function buildNumberingsXml() {
  return `<hh:numberings itemCnt="1"><hh:numbering id="1" start="0"><hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1" autoIndent="0" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="0" checkable="0">^1.</hh:paraHead></hh:numbering></hh:numberings>`;
}
function buildBulletsXml() {
  return `<hh:bullets itemCnt="1"><hh:bullet id="1" char="&#x2022;" useImage="0"><hh:paraHead level="0" align="LEFT" useInstWidth="0" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="0" checkable="0"/></hh:bullet></hh:bullets>`;
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
    charPrXml += `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" shadeColor="${shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1"><hh:fontRef hangul="${hid}" latin="${lid}" hanja="${cp.hanjaId}" japanese="${cp.japaneseId}" other="${cp.otherId}" symbol="${cp.symbolId}" user="${cp.userId}"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>` + bold + italic + `<hh:underline type="${cp.underline}" shape="SOLID" color="#000000"/><hh:strikeout shape="${cp.strikeout}" color="#000000"/><hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>`;
  }
  let paraPrXml = "";
  for (const pp of ctx.paraPrs) {
    const ver = pp.verAlign ?? "BASELINE";
    const wrap = pp.lineWrap ?? "BREAK";
    const lsType = pp.lineSpacingFixed !== void 0 ? "AT_LEAST" : "PERCENT";
    const lsValue = pp.lineSpacingFixed !== void 0 ? pp.lineSpacingFixed : pp.lineSpacing;
    paraPrXml += `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0"><hh:align horizontal="${pp.align}" vertical="${ver}"/><hh:heading type="NONE" idRef="0" level="0"/><hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="${wrap}"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/><hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"><hh:margin><hc:indent value="${pp.intentHwp}" unit="HWPUNIT"/><hc:left value="${pp.leftHwp}" unit="HWPUNIT"/><hc:right value="${pp.rightHwp}" unit="HWPUNIT"/><hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/><hc:next value="${pp.nextHwp}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="${lsType}" value="${lsValue}" unit="HWPUNIT"/></hp:case><hp:default><hh:margin><hc:indent value="${pp.intentHwp}" unit="HWPUNIT"/><hc:left value="${pp.leftHwp}" unit="HWPUNIT"/><hc:right value="${pp.rightHwp}" unit="HWPUNIT"/><hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/><hc:next value="${pp.nextHwp}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="${lsType}" value="${lsValue}" unit="HWPUNIT"/></hp:default></hp:switch><hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`;
  }
  const borderFillXml = ctx.borderFillBank.toXml();
  const denseStyles = materializeContiguousStyles(ctx.hwpxStyles);
  const stylesXml2 = `<hh:styles itemCnt="${denseStyles.length}">` + denseStyles.map(
    (s) => `<hh:style id="${s.id}" type="PARA" name="${esc(s.name)}" engName="${esc(s.engName)}" paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="0" langID="1042" lockForm="0"/>`
  ).join("") + `</hh:styles>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head ${NS} version="1.4" secCnt="1"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList>` + fontFacesXml + borderFillXml + `<hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>` + buildNumberingsXml() + buildBulletsXml() + `<hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties>` + stylesXml2 + `</hh:refList><hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument><hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption><hh:trackchageConfig flags="56"/></hh:head>`;
}
function buildHeaderFooterRunXml(sheet, dims, ctx) {
  const headers = sheet.headers || {};
  const footers = sheet.footers || {};
  const hasAny = Object.keys(headers).length > 0 || Object.keys(footers).length > 0;
  if (!hasAny) return "";
  const availW = ctx.availableWidth;
  const mtHwp = Metric.ptToHwp(dims.mt);
  const mbHwp = Metric.ptToHwp(dims.mb);
  const headerZoneH = dims.headerPt ? Metric.ptToHwp(dims.headerPt) : 4252;
  const footerZoneH = dims.footerPt ? Metric.ptToHwp(dims.footerPt) : 4252;
  let inner = "";
  const hideFirst = !!(headers.first || footers.first);
  inner += `<hp:ctrl><hp:pageHiding hideHeader="${hideFirst ? 1 : 0}" hideFooter="${hideFirst ? 1 : 0}" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0"/></hp:ctrl>`;
  for (const [type, paras] of Object.entries(headers)) {
    if (!Array.isArray(paras) || paras.length === 0) continue;
    const applyPageType = type === "even" ? "EVEN" : type === "default" || type === "first" ? "BOTH" : "ODD";
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const parasXml = paras.map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml).join("");
    ctx.nextElementId = savedId;
    inner += `<hp:ctrl><hp:header id="1" applyPageType="${applyPageType}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${headerZoneH}" hasTextRef="0" hasNumRef="0">` + parasXml + `</hp:subList></hp:header></hp:ctrl>`;
  }
  for (const [type, paras] of Object.entries(footers)) {
    if (!Array.isArray(paras) || paras.length === 0) continue;
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
  const sectionControlRunXml = `<hp:run charPrIDRef="0" charTcId="0">` + secPrXml + `<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl></hp:run>`;
  const kids = sheet?.kids ?? [];
  const hfRunXml = sheet ? buildHeaderFooterRunXml(sheet, dims, ctx) : "";
  const availWidth = Math.max(
    1e3,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr)
  );
  const bodyHeight = Math.max(
    1e3,
    Metric.ptToHwp(dims.hPt) - Metric.ptToHwp(dims.mt) - Metric.ptToHwp(dims.mb)
  );
  ctx.availableWidth = availWidth;
  let contentXml = "";
  let vertPos = 0;
  let pageFirst = true;
  for (let i = 0; i < kids.length; i++) {
    const kid = kids[i];
    const isFirst = i === 0;
    const curSecPr = isFirst ? sectionControlRunXml : "";
    const curHfRun = isFirst ? hfRunXml : "";
    if (kid.tag === "para") {
      if (paraHasPageBreak(kid)) {
        vertPos = 0;
        pageFirst = true;
      }
      const { xml, nextVertPos, hasPageBreak } = encodeParaPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        availWidth,
        curHfRun,
        pageFirst
      );
      contentXml += xml;
      if (nextVertPos >= bodyHeight) {
        vertPos = 0;
        pageFirst = true;
      } else {
        vertPos = nextVertPos;
        pageFirst = false;
      }
    } else if (kid.tag === "grid") {
      const { xml, nextVertPos, hasPageBreak } = encodeGridPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        curHfRun,
        pageFirst
      );
      contentXml += xml;
      if (nextVertPos >= bodyHeight) {
        vertPos = 0;
        pageFirst = true;
      } else {
        vertPos = nextVertPos;
        pageFirst = false;
      }
    }
  }
  if (kids.length === 0) {
    const fs = 1e3;
    const vs = 1600;
    const { xml: linesegXml } = buildLinesegarray(
      " ",
      0,
      fs,
      vs / (fs / 100),
      availWidth,
      void 0,
      { pageFirst: true }
    );
    contentXml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + sectionControlRunXml + hfRunXml + `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>` + linesegXml + `</hp:p>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${contentXml}</hs:sec>`;
}
function buildSecPrXml(dims) {
  const wHwp = Metric.ptToHwp(dims.wPt);
  const hHwp = Metric.ptToHwp(dims.hPt);
  const ml = Metric.ptToHwp(dims.ml);
  const mr = Metric.ptToHwp(dims.mr);
  const mt = Metric.ptToHwp(dims.mt);
  const mb = Metric.ptToHwp(dims.mb);
  const headerZone = dims.headerPt ? Metric.ptToHwp(dims.headerPt) : 0;
  const footerZone = dims.footerPt ? Metric.ptToHwp(dims.footerPt) : 0;
  const pageBorderFill = `<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>`;
  return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="WIDELY" width="${wHwp}" height="${hHwp}" gutterType="LEFT_ONLY"><hp:margin header="${headerZone}" footer="${footerZone}" gutter="0" left="${ml}" right="${mr}" top="${mt}" bottom="${mb}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>` + pageBorderFill + `</hp:secPr>`;
}
function buildLinesegarray(text, vertPosStart, fontSize, lineSpacingPct, horzSize, lineHeightHwp, layout = {}) {
  const textHeight = Math.max(fontSize, layout.textHeight ?? fontSize);
  const lineAdvance = Math.max(
    textHeight,
    lineHeightHwp ?? Math.round(fontSize * Math.max(100, lineSpacingPct) / 100)
  );
  const spacing = Math.max(0, lineAdvance - textHeight);
  const baseline = Math.round(textHeight * 0.85);
  const firstHorzPos = Math.max(0, layout.firstHorzPos ?? 0);
  const restHorzPos = Math.max(0, layout.restHorzPos ?? firstHorzPos);
  const rightMargin = Math.max(0, layout.rightMargin ?? 0);
  const lineHorzPos = (index) => index === 0 ? firstHorzPos : restHorzPos;
  const lineHorzSize = (index) => Math.max(100, horzSize - lineHorzPos(index) - rightMargin);
  if (text.length === 0) {
    const xml = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${vertPosStart}" vertsize="${textHeight}" textheight="${textHeight}" baseline="${baseline}" spacing="${spacing}" horzpos="${firstHorzPos}" horzsize="${lineHorzSize(0)}" flags="${LINESEG_FLAGS | (layout.indentFirst ? LINESEG_FLAG_INDENT : 0) | (layout.pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/></hp:linesegarray>`;
    return { xml, totalHeight: lineAdvance };
  }
  const lines = [];
  let currentLineWidth = 0;
  let lineStartIdx = 0;
  for (let i = 0; i < text.length; i++) {
    const charCode = text.charCodeAt(i);
    if (charCode === 10 || charCode === 13) {
      lines.push({ startPos: lineStartIdx, width: currentLineWidth });
      if (charCode === 13 && text.charCodeAt(i + 1) === 10) i++;
      lineStartIdx = i + 1;
      currentLineWidth = 0;
      continue;
    }
    let charW = fontSize * 0.55;
    if (charCode >= 44032 && charCode <= 55203) {
      charW = fontSize;
    } else if (charCode >= 12592 && charCode <= 12687) {
      charW = fontSize;
    } else if (charCode >= 19968 && charCode <= 40959) {
      charW = fontSize;
    } else if (charCode >= 65 && charCode <= 90) {
      charW = fontSize * 0.65;
    } else if (charCode === 32) {
      charW = fontSize * 0.32;
    } else if (charCode > 255) {
      charW = fontSize;
    } else {
      charW = fontSize * 0.42;
    }
    if (currentLineWidth + charW > lineHorzSize(lines.length) && i > lineStartIdx) {
      lines.push({ startPos: lineStartIdx, width: currentLineWidth });
      lineStartIdx = i;
      currentLineWidth = charW;
    } else {
      currentLineWidth += charW;
    }
  }
  lines.push({ startPos: lineStartIdx, width: currentLineWidth });
  const lineCount = lines.length;
  const linesegParts = [];
  for (let i = 0; i < lineCount; i++) {
    const flags = LINESEG_FLAGS | (i === 0 && layout.indentFirst ? LINESEG_FLAG_INDENT : 0) | (i === 0 && layout.pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0);
    const textpos = lines[i].startPos;
    linesegParts.push(
      `<hp:lineseg textpos="${textpos}" vertpos="${vertPosStart + i * lineAdvance}" vertsize="${textHeight}" textheight="${textHeight}" baseline="${baseline}" spacing="${spacing}" horzpos="${lineHorzPos(i)}" horzsize="${lineHorzSize(i)}" flags="${flags}"/>`
    );
  }
  return {
    xml: `<hp:linesegarray>${linesegParts.join("")}</hp:linesegarray>`,
    totalHeight: lineCount * lineAdvance
  };
}
function extractParaText(para) {
  let text = "";
  const walk = (kids) => {
    for (const k of kids) {
      if (k.tag === "span") {
        for (const c of k.kids) {
          if (c.tag === "txt") {
            text += c.content;
          } else if (c.tag === "br") {
            text += "\n";
          }
        }
      } else if (k.tag === "link") {
        walk(k.kids);
      }
    }
  };
  walk(para.kids);
  return text;
}
function paraHasPageBreak(para) {
  const visit = (kids) => kids.some((kid) => {
    if (kid.tag === "span") return kid.kids.some((child) => child.tag === "pb");
    if (kid.tag === "link") {
      return visit(kid.kids);
    }
    return false;
  });
  return visit(para.kids);
}
function fontSizeForPara(para, ctx) {
  let maxSize = 1e3;
  const visit = (kids) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span") {
        const id = ctx.charPrMap.get(charPrKey(kid.props));
        if (id !== void 0 && ctx.charPrs[id]) {
          maxSize = Math.max(maxSize, ctx.charPrs[id].height);
        }
      } else if (kid.tag === "link") {
        visit(kid.kids ?? []);
      }
    }
  };
  visit(para.kids);
  return maxSize;
}
function inlineObjectHeightForPara(para, ctx) {
  let maxHeight = 0;
  for (const kid of para.kids) {
    if (kid.tag !== "img" || kid.layout && kid.layout.wrap !== "inline") continue;
    const dims = getImageDisplayDims(kid, ctx);
    maxHeight = Math.max(maxHeight, dims.h);
  }
  return maxHeight;
}
function encodeParaPositioned(para, ctx, vertPos, secPr = "", availWidth, hfRun = "", pageFirst = false) {
  const gridKid = para.kids.find((k) => k.tag === "grid");
  if (gridKid) {
    return encodeTablePara(para, gridKid, ctx, vertPos, secPr, hfRun, pageFirst);
  }
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const styleKey = paraStyleKey(para.props);
  const styleIDRef = styleKey ? ctx.styleIdToHwpxId.get(styleKey) ?? 0 : 0;
  const fontSize = fontSizeForPara(para, ctx);
  const paraPr = ctx.paraPrs[paraPrId];
  const lineSpacing = paraPr?.lineSpacing ?? 160;
  const lineHeightHwp = paraPr?.lineSpacingFixed !== void 0 ? Math.max(
    paraShapeHwpToLayoutHwp(paraPr.lineSpacingFixed),
    Math.ceil(fontSize * 1.15)
  ) : Math.max(fontSize, Math.round(fontSize * Math.max(100, lineSpacing) / 100));
  const textHeight = Math.max(fontSize, inlineObjectHeightForPara(para, ctx));
  const effectiveLineHeight = textHeight + Math.max(0, lineHeightHwp - fontSize);
  const spacing = Math.max(0, effectiveLineHeight - textHeight);
  let vertSize = effectiveLineHeight;
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
      vertSize,
      pageFirst
    );
  let runsXml = encodeParaKids(para.kids, ctx);
  if (!runsXml) runsXml = `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>`;
  const paraText2 = extractParaText(para);
  const paraStart = vertPos + Math.max(0, paraShapeHwpToLayoutHwp(paraPr?.prevHwp ?? 0));
  const firstHorzPos = Math.max(
    0,
    paraShapeHwpToLayoutHwp(
      (paraPr?.leftHwp ?? 0) + (paraPr?.intentHwp ?? 0)
    )
  );
  const restHorzPos = Math.max(
    0,
    paraShapeHwpToLayoutHwp(paraPr?.leftHwp ?? 0)
  );
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText2,
    paraStart,
    fontSize,
    lineSpacing,
    horzSize,
    effectiveLineHeight,
    {
      textHeight,
      firstHorzPos,
      restHorzPos,
      rightMargin: Math.max(
        0,
        paraShapeHwpToLayoutHwp(paraPr?.rightHwp ?? 0)
      ),
      indentFirst: (paraPr?.intentHwp ?? 0) !== 0,
      pageFirst
    }
  );
  const hasPageBreak = paraHasPageBreak(para);
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="${styleIDRef}" pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0" paraTcId="0">` + secPr + hfRun + runsXml + linesegXml + `</hp:p>`;
  return {
    xml,
    nextVertPos: paraStart + totalHeight + Math.max(0, paraShapeHwpToLayoutHwp(paraPr?.nextHwp ?? 0)),
    hasPageBreak
  };
}
function encodeTablePara(para, grid, ctx, vertPos, secPr, hfRun, pageFirst) {
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  const totalHeight = Math.max(1600, tblHeight);
  const baseline = Math.round(totalHeight * 0.85);
  const linesegXml = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" textheight="${totalHeight}" baseline="${baseline}" spacing="0" horzpos="0" horzsize="${ctx.availableWidth}" flags="${LINESEG_FLAGS | (pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/></hp:linesegarray>`;
  const hasPageBreak = paraHasPageBreak(para);
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="0" pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0" paraTcId="0">` + secPr + `<hp:run charPrIDRef="0" charTcId="0">` + gridXml + `</hp:run>` + hfRun + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak };
}
function encodeCodeBlockPositioned(para, ctx, vertPos, secPr, fontSize, spacing, vertSize, pageFirst) {
  const codeBfId = ctx.borderFillBank.addUniform(
    { kind: "solid", pt: 0.5, color: "aaaaaa" },
    "f4f4f4"
  );
  const cellW = ctx.availableWidth;
  const innerW = Math.max(cellW - 510, 100);
  const subListId = ctx.nextElementId++;
  const { xml: innerXml } = encodeParaPositioned(para, ctx, 0, "", innerW);
  const paraText2 = extractParaText(para);
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText2,
    vertPos,
    fontSize,
    160,
    // 코드 블록 기본 줄간격 160%
    ctx.availableWidth,
    void 0,
    { pageFirst }
  );
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + secPr + `<hp:run charPrIDRef="0" charTcId="0"><hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="${codeBfId}" noAdjust="0"><hp:sz width="${cellW}" widthRelTo="ABSOLUTE" height="0" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="138" right="138" top="138" bottom="138"/><hp:inMargin left="138" right="138" top="138" bottom="138"/><hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${codeBfId}"><hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` + innerXml + `</hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${cellW}" height="0"/><hp:cellMargin left="283" right="283" top="141" bottom="141"/></hp:tc></hp:tr></hp:tbl><hp:t xml:space="preserve"> </hp:t></hp:run>` + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak: false };
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
      xml += `<hp:br/>`;
    } else if (kid.tag === "pagenum") {
      const fmt = kid.format === "roman" ? "ROMAN_LOWER" : kid.format === "romanCaps" ? "ROMAN_UPPER" : "DIGIT";
      const numType = kid.format === "total" ? "TOTAL_PAGE" : "PAGE";
      xml += `<hp:ctrl><hp:autoNum num="1" numType="${numType}"><hp:autoNumFormat type="${fmt}" userChar="" prefixChar="" suffixChar="" supscript="0"/></hp:autoNum></hp:ctrl>`;
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
function getImageSourceDims(img) {
  const pixelDims = img.b64 ? readPixelDims(img.b64, img.mime) : null;
  if (pixelDims && pixelDims.w > 0 && pixelDims.h > 0) {
    return {
      w: Metric.ptToHwp(pixelDims.w * 72 / 96),
      h: Metric.ptToHwp(pixelDims.h * 72 / 96)
    };
  }
  return {
    w: Math.max(1, Metric.ptToHwp(img.w || 1)),
    h: Math.max(1, Metric.ptToHwp(img.h || 1))
  };
}
function getImageDisplayDims(img, ctx) {
  const source = getImageSourceDims(img);
  let w = img.w > 0 ? Metric.ptToHwp(img.w) : source.w;
  let h = img.h > 0 ? Metric.ptToHwp(img.h) : source.h;
  if (w > ctx.availableWidth) {
    h = Math.round(h * ctx.availableWidth / w);
    w = ctx.availableWidth;
  }
  return { w: Math.max(1, w), h: Math.max(1, h) };
}
function encodeImage(img, ctx) {
  if (!img.b64) {
    return `<hp:t xml:space="preserve">${esc(img.alt || "[\uAC1C\uCCB4]")}</hp:t>`;
  }
  const binId = ctx.imgMap.get(img);
  if (!binId) return "";
  const { w: wHwp, h: hHwp } = getImageDisplayDims(img, ctx);
  const sourceDims = getImageSourceDims(img);
  const rotationCenterX = Math.round(wHwp / 2);
  const rotationCenterY = Math.round(hHwp / 2);
  const layout = img.layout;
  const isInline = !layout || layout.wrap === "inline";
  const textWrap = layout ? WRAP_MAP[layout.wrap] ?? "SQUARE" : "SQUARE";
  const textFlow = layout ? FLOW_MAP[layout.wrap] ?? "BOTH_SIDES" : "BOTH_SIDES";
  const zOrder = ctx.nextZOrder++;
  return `<hp:pic id="${ctx.nextElementId++}" zOrder="${zOrder}" numberingType="PICTURE" textWrap="${textWrap}" textFlow="${textFlow}" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="0" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${wHwp}" height="${hHwp}"/><hp:curSz width="${wHwp}" height="${hHwp}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="${rotationCenterX}" centerY="${rotationCenterY}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${wHwp}" y="0"/><hc:pt2 x="${wHwp}" y="${hHwp}"/><hc:pt3 x="0" y="${hHwp}"/></hp:imgRect><hp:imgClip left="0" right="${sourceDims.w}" top="0" bottom="${sourceDims.h}"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="${sourceDims.w}" dimheight="${sourceDims.h}"/><hc:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:effects/><hp:sz width="${wHwp}" widthRelTo="ABSOLUTE" height="${hHwp}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="${isInline ? 1 : 0}" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic>`;
}
function encodeImgWrapped(img, ctx) {
  const content = encodeImage(img, ctx);
  if (!img.b64) {
    return `<hp:run charPrIDRef="0" charTcId="0">${content}</hp:run>`;
  }
  return `<hp:run charPrIDRef="0" charTcId="0">${content}<hp:t xml:space="preserve"> </hp:t></hp:run>`;
}
function encodeGridPositioned(grid, ctx, vertPos, secPr = "", hfRun = "", pageFirst = false) {
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  const floats = grid.props.layout !== void 0 && grid.props.layout.wrap !== "inline";
  const totalHeight = floats ? 1e3 : Math.max(1600, tblHeight);
  const baseline = Math.round(totalHeight * 0.85);
  const linesegXml = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" textheight="${totalHeight}" baseline="${baseline}" spacing="0" horzpos="0" horzsize="${ctx.availableWidth}" flags="${LINESEG_FLAGS | (pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/></hp:linesegarray>`;
  const xml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` + secPr + hfRun + `<hp:run charPrIDRef="0" charTcId="0">` + gridXml + `</hp:run>` + linesegXml + `</hp:p>`;
  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak: false };
}
function buildGridLayoutAttrs(layout, fallbackHorzAlign) {
  const floats = layout !== void 0 && layout.wrap !== "inline";
  const textWrapMap = {
    inline: "TOP_AND_BOTTOM",
    topAndBottom: "TOP_AND_BOTTOM",
    square: "SQUARE",
    tight: "BOTH_SIDES",
    through: "BOTH_SIDES",
    none: "TOP_AND_BOTTOM",
    behind: "BEHIND_TEXT",
    front: "FRONT_TEXT"
  };
  const horzRelMap = {
    para: "PARA",
    margin: "MARGIN",
    page: "PAGE",
    column: "COLUMN"
  };
  const vertRelMap = {
    para: "PARA",
    margin: "MARGIN",
    page: "PAGE",
    line: "LINE"
  };
  const horzAlignMap = {
    left: "LEFT",
    center: "CENTER",
    right: "RIGHT"
  };
  const vertAlignMap = {
    top: "TOP",
    center: "CENTER",
    bottom: "BOTTOM"
  };
  const horzAlign = (layout?.horzAlign ? horzAlignMap[layout.horzAlign] : void 0) ?? fallbackHorzAlign;
  const vertAlign = (layout?.vertAlign ? vertAlignMap[layout.vertAlign] : void 0) ?? "TOP";
  const horzRelTo = (layout?.horzRelTo ? horzRelMap[layout.horzRelTo] : void 0) ?? "PARA";
  const vertRelTo = (layout?.vertRelTo ? vertRelMap[layout.vertRelTo] : void 0) ?? "PARA";
  const horzOffset = layout?.xPt != null ? Metric.ptToHwp(layout.xPt) : 0;
  const vertOffset = layout?.yPt != null ? Metric.ptToHwp(layout.yPt) : 0;
  const posXml = `<hp:pos treatAsChar="${floats ? "0" : "1"}" affectLSpacing="0" flowWithText="${floats && (layout?.vertRelTo === "page" || layout?.horzRelTo === "page") ? "0" : "1"}" allowOverlap="${floats ? "1" : "0"}" holdAnchorAndSO="0" vertRelTo="${vertRelTo}" horzRelTo="${horzRelTo}" vertAlign="${vertAlign}" horzAlign="${horzAlign}" vertOffset="${vertOffset}" horzOffset="${horzOffset}"/>`;
  const dist = (value) => value != null ? Math.max(0, Metric.ptToHwp(value)) : 138;
  const outMarginXml = `<hp:outMargin left="${dist(layout?.distL)}" right="${dist(layout?.distR)}" top="${dist(layout?.distT)}" bottom="${dist(layout?.distB)}"/>`;
  return {
    textWrap: textWrapMap[layout?.wrap ?? "inline"] ?? "TOP_AND_BOTTOM",
    zOrder: Math.round(layout?.zOrder ?? 0),
    noAdjust: floats ? "1" : "0",
    posXml,
    outMarginXml
  };
}
function buildGridXml(grid, ctx, maxWidth = ctx.availableWidth) {
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
  const totalW = Math.max(1, Math.min(ctx.availableWidth, maxWidth));
  const sourceWidths = (grid.props.colWidths ?? []).map(
    (width) => width > 0 ? Metric.ptToHwp(width) : 0
  );
  const colWidths = fitColumnWidths(
    sourceWidths,
    colCount,
    totalW,
    Math.min(100, Math.floor(totalW / colCount))
  );
  const actualTotal = colWidths.reduce((s, w) => s + w, 0);
  const tablePadL = Metric.ptToHwp(grid.props.cellPadL ?? 1.41);
  const tablePadR = Metric.ptToHwp(grid.props.cellPadR ?? 1.41);
  const tablePadT = Metric.ptToHwp(grid.props.cellPadT ?? 1.41);
  const tablePadB = Metric.ptToHwp(grid.props.cellPadB ?? 1.41);
  const rowHeights = [];
  for (let ri = 0; ri < rowCount; ri++) {
    let minRowH = 0;
    for (let ci = 0; ci < colCount; ci++) {
      const entry = tableMap[ri][ci];
      if (entry?.type === "real") {
        const cell = entry.cell;
        const cp = cell.props ?? {};
        let cellW = 0;
        for (let sc = ci; sc < ci + cell.cs && sc < colWidths.length; sc++)
          cellW += colWidths[sc];
        if (!cellW) cellW = Math.round(totalW / colCount) * cell.cs;
        const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : tablePadL;
        const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : tablePadR;
        const innerW = Math.max(cellW - padL - padR, 100);
        const span = Math.max(1, cell.rs ?? 1);
        const h = estimateCellHeight(cell, ctx, innerW);
        minRowH = Math.max(minRowH, Math.ceil(h / span));
      }
    }
    const baseH = grid.kids[ri].heightPt != null && grid.kids[ri].heightPt > 0 ? Metric.ptToHwp(grid.kids[ri].heightPt) : Math.round(1e3 * 1.6);
    if (grid.kids[ri].heightPt != null && grid.kids[ri].heightPt > 0) {
      rowHeights.push(Math.max(baseH, minRowH));
    } else {
      rowHeights.push(Math.max(baseH, minRowH));
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
      const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : tablePadL;
      const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : tablePadR;
      const padT = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : tablePadT;
      const padB = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : tablePadB;
      const innerW = Math.max(cellW - padL - padR, 100);
      let parasXml = "";
      let localVertPos = 0;
      if (cell.kids.length > 0) {
        for (const kid of cell.kids) {
          if (kid.tag === "grid") {
            const { xml: tblXml, height: nestedHeight } = buildGridXml(kid, ctx, innerW);
            const pid = ctx.nextElementId++;
            const objectHeight = Math.max(1600, nestedHeight);
            const baseline = Math.round(objectHeight * 0.85);
            parasXml += `<hp:p id="${pid}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0"><hp:run charPrIDRef="0" charTcId="0">${tblXml}</hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="${localVertPos}" vertsize="${objectHeight}" textheight="${objectHeight}" baseline="${baseline}" spacing="0" horzpos="0" horzsize="${innerW}" flags="${LINESEG_FLAGS}"/></hp:linesegarray></hp:p>`;
            localVertPos += objectHeight;
          } else {
            const encoded = encodeParaPositioned(kid, ctx, localVertPos, "", innerW);
            parasXml += encoded.xml;
            localVertPos = encoded.nextVertPos;
          }
        }
      } else {
        const { xml: emptyLineseg } = buildLinesegarray(" ", 0, 1e3, 160, innerW);
        parasXml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0"><hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>${emptyLineseg}</hp:p>`;
      }
      const vAlign = cp.va === "mid" ? "CENTER" : cp.va === "bot" ? "BOTTOM" : "TOP";
      const cellHeight = rowHeights.slice(ri, Math.min(rowHeights.length, ri + Math.max(1, cell.rs))).reduce((sum, height) => sum + height, 0);
      cellsXml += `<hp:tc name="" header="${cp.isHeader || grid.props.headerRow && ri === 0 ? 1 : 0}" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}"><hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vAlign}" linkListIDRef="0" linkListNextIDRef="0" textWidth="${innerW}" textHeight="${Math.max(100, cellHeight - padT - padB)}" hasTextRef="0" hasNumRef="0">` + parasXml + `</hp:subList><hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${cellW}" height="${cellHeight}"/><hp:cellMargin left="${padL}" right="${padR}" top="${padT}" bottom="${padB}"/></hp:tc>`;
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
  const layoutAttrs = buildGridLayoutAttrs(grid.props.layout, horzAlign);
  const repeatHeader = grid.props.headerRow ? 1 : 0;
  const xml = `<hp:tbl id="${ctx.nextElementId++}" zOrder="${layoutAttrs.zOrder}" numberingType="TABLE" textWrap="${layoutAttrs.textWrap}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="${repeatHeader}" rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="${layoutAttrs.noAdjust}"><hp:sz width="${actualTotal}" widthRelTo="ABSOLUTE" height="${totalH}" heightRelTo="ABSOLUTE" protect="0"/>` + layoutAttrs.posXml + layoutAttrs.outMarginXml + `<hp:inMargin left="${tablePadL}" right="${tablePadR}" top="${tablePadT}" bottom="${tablePadB}"/>` + rowsXml + `</hp:tbl>`;
  return { xml, height: totalH };
}
function estimateLineCountForWidth(text, fontSize, horzSize) {
  if (!text) return 2;
  const maxWidth = Math.max(1, horzSize ?? 0);
  if (!horzSize || horzSize <= 0) return text.split(/\r\n|\r|\n/).length;
  let lines = 1;
  let currentLineWidth = 0;
  for (let i = 0; i < text.length; i++) {
    const charCode = text.charCodeAt(i);
    if (charCode === 10 || charCode === 13) {
      if (charCode === 13 && text.charCodeAt(i + 1) === 10) i++;
      lines++;
      currentLineWidth = 0;
      continue;
    }
    let charW = fontSize * 0.55;
    if (charCode >= 44032 && charCode <= 55203) charW = fontSize;
    else if (charCode >= 12592 && charCode <= 12687) charW = fontSize;
    else if (charCode >= 19968 && charCode <= 40959) charW = fontSize;
    else if (charCode >= 65 && charCode <= 90) charW = fontSize * 0.65;
    else if (charCode === 32) charW = fontSize * 0.32;
    else if (charCode > 255) charW = fontSize;
    else charW = fontSize * 0.42;
    if (currentLineWidth > 0 && currentLineWidth + charW > maxWidth) {
      lines++;
      currentLineWidth = charW;
    } else {
      currentLineWidth += charW;
    }
  }
  return Math.max(1, lines);
}
function estimateGridHeight(grid, ctx) {
  let total = 0;
  for (const row of grid.kids) {
    const base = row.heightPt != null && row.heightPt > 0 ? Metric.ptToHwp(row.heightPt) : Math.round(1e3 * 1.6);
    let minRow = 0;
    for (const cell of row.kids) {
      const span = Math.max(1, cell.rs ?? 1);
      minRow = Math.max(minRow, Math.ceil(estimateCellHeight(cell, ctx) / span));
    }
    total += Math.max(base, minRow);
  }
  return total;
}
function estimateCellHeight(cell, ctx, innerWidth) {
  const cp = cell.props ?? {};
  const topPad = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : 141;
  const botPad = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : 141;
  let h = 0;
  for (const kid of cell.kids) {
    if (kid.tag === "grid") {
      h += estimateGridHeight(kid, ctx);
      continue;
    }
    const para = kid;
    const fs = fontSizeForPara(para, ctx);
    const ppId = ctx.paraPrMap.get(paraPrKey(para.props));
    const pp = ppId !== void 0 ? ctx.paraPrs[ppId] : null;
    const lineHeight = pp?.lineSpacingFixed !== void 0 ? Math.max(
      paraShapeHwpToLayoutHwp(pp.lineSpacingFixed),
      Math.ceil(fs * 1.15)
    ) : Math.max(fs, Math.round(fs * Math.max(100, pp?.lineSpacing ?? 160) / 100));
    const textHeight = Math.max(fs, inlineObjectHeightForPara(para, ctx));
    const lineAdvance = textHeight + Math.max(0, lineHeight - fs);
    const lineCount = estimateLineCountForWidth(
      extractParaText(para),
      fs,
      innerWidth
    );
    const before = paraShapeHwpToLayoutHwp(pp?.prevHwp ?? 0);
    const after = paraShapeHwpToLayoutHwp(pp?.nextHwp ?? 0);
    h += lineAdvance * lineCount + before + after;
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

// src/encoders/docx/font-mapping.json
var font_mapping_default = {
  \uD568\uCD08\uB86C\uBC14\uD0D5: {
    nearest: "\uBC14\uD0D5",
    candidates: [
      { name: "\uBC14\uD0D5", distance: 0 },
      { name: "Noto Serif CJK KR", distance: 1 },
      { name: "Noto Serif KR", distance: 1 }
    ]
  },
  \uD568\uCD08\uB86C\uB3CB\uC6C0: {
    nearest: "\uB9D1\uC740 \uACE0\uB515",
    candidates: [
      { name: "\uB9D1\uC740 \uACE0\uB515", distance: 0 },
      { name: "Noto Sans CJK KR", distance: 1 },
      { name: "Noto Sans KR", distance: 1 }
    ]
  },
  \uD55C\uC591\uC2E0\uBA85\uC870: {
    nearest: "\uBC14\uD0D5",
    candidates: [
      { name: "\uBC14\uD0D5", distance: 0 },
      { name: "Noto Serif CJK KR", distance: 1 }
    ]
  },
  HY\uC2E0\uBA85\uC870: {
    nearest: "\uBC14\uD0D5",
    candidates: [
      { name: "\uBC14\uD0D5", distance: 0 },
      { name: "Noto Serif CJK KR", distance: 1 }
    ]
  },
  \uD734\uBA3C\uBA85\uC870: {
    nearest: "\uBC14\uD0D5",
    candidates: [
      { name: "\uBC14\uD0D5", distance: 0 },
      { name: "Noto Serif CJK KR", distance: 1 }
    ]
  }
};

// src/encoders/docx/DocxEncoder.ts
var FONT_MAPPING = font_mapping_default;
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
        dims,
        nextId: 10,
        nextImgNum: 1,
        warns: [],
        imgMap: /* @__PURE__ */ new WeakMap()
      };
      collectImages(allKids, ctx);
      const headerContents = [...firstSheet?.headers?.default ?? []];
      const footerContents = [...firstSheet?.footers?.default ?? []];
      const hasHeader = headerContents.length > 0;
      const hasFooter = footerContents.length > 0;
      if (hasHeader) collectImages(headerContents, ctx);
      if (hasFooter) collectImages(footerContents, ctx);
      const fonts = collectFonts(allKids);
      if (hasHeader) collectFonts(headerContents, fonts);
      if (hasFooter) collectFonts(footerContents, fonts);
      fonts.add("\uD568\uCD08\uB86C\uBC14\uD0D5");
      fonts.add("\uB9D1\uC740 \uACE0\uB515");
      const hasFontTable = fonts.size > 0;
      const headerRId = hasHeader ? `rId${ctx.nextId++}` : "";
      const footerRId = hasFooter ? `rId${ctx.nextId++}` : "";
      const numInfo = collectNumbering(allKids);
      const kids = allKids;
      const mainDocumentXml = documentXml(kids, dims, ctx, headerRId, footerRId);
      const headerXml = hasHeader ? headerFooterXml("hdr", headerContents, ctx, dims) : "";
      const footerXml = hasFooter ? headerFooterXml("ftr", footerContents, ctx, dims) : "";
      const entries = [
        {
          name: "[Content_Types].xml",
          data: this.stringToBytes(contentTypes(images, hasHeader, hasFooter, hasFontTable))
        },
        { name: "_rels/.rels", data: this.stringToBytes(pkgRels()) },
        {
          name: "word/document.xml",
          data: this.stringToBytes(mainDocumentXml)
        },
        { name: "word/styles.xml", data: this.stringToBytes(stylesXml()) },
        { name: "word/settings.xml", data: this.stringToBytes(settingsXml()) },
        {
          name: "word/_rels/document.xml.rels",
          data: this.stringToBytes(
            docRels(images, headerRId, footerRId, numInfo.hasLists, hasFontTable)
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
      if (hasFontTable) {
        entries.push({
          name: "word/fontTable.xml",
          data: this.stringToBytes(fontTableXml(fonts))
        });
      }
      if (hasHeader) {
        entries.push({
          name: "word/header1.xml",
          data: this.stringToBytes(headerXml)
        });
        entries.push({
          name: "word/_rels/header1.xml.rels",
          data: this.stringToBytes(imagePartRels(images))
        });
      }
      if (hasFooter) {
        entries.push({
          name: "word/footer1.xml",
          data: this.stringToBytes(footerXml)
        });
        entries.push({
          name: "word/_rels/footer1.xml.rels",
          data: this.stringToBytes(imagePartRels(images))
        });
      }
      for (const img of images) {
        entries.push({ name: `word/media/${img.name}`, data: img.data });
      }
      return succeed(await this.zip(entries), ctx.warns);
    } catch (e) {
      return fail(`DOCX encode error: ${e?.message ?? String(e)}`);
    }
  }
};
function mimeToExt2(mime) {
  if (mime.includes("jpeg")) return "jpeg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  if (mime.includes("wmf")) return "wmf";
  if (mime.includes("emf")) return "emf";
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
function collectImagesFromPara(para, ctx) {
  for (const kid of para.kids) {
    if (kid.tag === "img") registerImage2(kid, ctx);
  }
}
function registerImage2(img, ctx) {
  if (ctx.imgMap.has(img)) return;
  const data = TextKit.base64Decode(img.b64);
  const ext = imageExtFromBytes(data) ?? mimeToExt2(img.mime);
  const name = `image${ctx.nextImgNum++}.${ext}`;
  const rId = `rId${ctx.nextId++}`;
  ctx.images.push({ rId, name, data, ext });
  ctx.imgMap.set(img, rId);
}
function imageExtFromBytes(data) {
  if (data.length >= 8 && data[0] === 137 && data[1] === 80 && data[2] === 78 && data[3] === 71) return "png";
  if (data.length >= 3 && data[0] === 255 && data[1] === 216 && data[2] === 255) return "jpeg";
  if (data.length >= 6 && data[0] === 71 && data[1] === 73 && data[2] === 70) return "gif";
  if (data.length >= 2 && data[0] === 66 && data[1] === 77) return "bmp";
  if (data.length >= 4 && data[0] === 215 && data[1] === 205 && data[2] === 198 && data[3] === 154) return "wmf";
  if (data.length >= 44 && data[40] === 32 && data[41] === 69 && data[42] === 77 && data[43] === 70) return "emf";
  return void 0;
}
function collectFonts(kids, fonts = /* @__PURE__ */ new Set()) {
  for (const kid of kids) {
    if (kid.tag === "para") collectFontsFromPara(kid, fonts);
    else if (kid.tag === "grid") {
      for (const row of kid.kids) {
        for (const cell of row.kids) {
          for (const child of cell.kids) {
            if (child.tag === "para") collectFontsFromPara(child, fonts);
            else collectFonts([child], fonts);
          }
        }
      }
    }
  }
  return fonts;
}
function collectFontsFromPara(para, fonts) {
  for (const kid of para.kids) {
    if (kid.tag === "span") collectFontsFromSpan(kid, fonts);
    else if (kid.tag === "link") {
      for (const span of kid.kids) collectFontsFromSpan(span, fonts);
    } else if (kid.tag === "grid") {
      collectFonts([kid], fonts);
    }
  }
}
function collectFontsFromSpan(span, fonts) {
  const font = span.props.font?.trim();
  if (font) fonts.add(font);
}
function mappedFontName(font) {
  const entry = FONT_MAPPING[font] ?? FONT_MAPPING[font.trim()];
  if (!entry) return void 0;
  if (typeof entry === "string") return entry;
  if (entry.altName) return entry.altName;
  if (entry.nearest) return entry.nearest;
  const first = entry.candidates?.[0];
  if (typeof first === "string") return first;
  return first?.name ?? first?.font;
}
function fontTableXml(fonts) {
  const body = Array.from(fonts).filter(Boolean).sort((a, b) => a.localeCompare(b, "ko")).map((font) => {
    const alt = mappedFontName(font);
    const altXml = alt && alt !== font ? `<w:altName w:val="${esc2(alt)}"/>` : "";
    return `<w:font w:name="${esc2(font)}">${altXml}<w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`;
  }).join("\n  ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${body}
</w:fonts>`;
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
function contentTypes(images, hasHeader, hasFooter, hasFontTable) {
  const imgDefaults = /* @__PURE__ */ new Set();
  for (const img of images) imgDefaults.add(img.ext);
  let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>`;
  for (const ext of imgDefaults) {
    const ct = ext === "png" ? "image/png" : ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "svg" ? "image/svg+xml" : ext === "wmf" ? "image/x-wmf" : ext === "emf" ? "image/x-emf" : "image/bmp";
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
  if (hasFontTable)
    overrides += `
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`;
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
function docRels(images, headerRId, footerRId, hasLists, hasFontTable) {
  let rels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`;
  if (hasLists) {
    rels += `
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
  }
  if (hasFontTable) {
    rels += `
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`;
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
function imagePartRels(images) {
  const rels = images.map(
    (img) => `<Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`
  ).join("\n  ");
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
      <w:rFonts w:ascii="\uD568\uCD08\uB86C\uBC14\uD0D5" w:eastAsia="\uD568\uCD08\uB86C\uBC14\uD0D5" w:hAnsi="\uD568\uCD08\uB86C\uBC14\uD0D5" w:hint="eastAsia"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="384" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="0"><w:name w:val="\uBC14\uD0D5\uAE00"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="\uBCF8\uBB38"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="\uAC1C\uC694 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="3"><w:name w:val="\uAC1C\uC694 2"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="4"><w:name w:val="\uAC1C\uC694 3"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="5"><w:name w:val="\uAC1C\uC694 4"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="6"><w:name w:val="\uAC1C\uC694 5"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="7"><w:name w:val="\uAC1C\uC694 6"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="8"><w:name w:val="\uAC1C\uC694 7"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="9"><w:name w:val="\uAC1C\uC694 8"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="10"><w:name w:val="\uAC1C\uC694 9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="11"><w:name w:val="\uAC1C\uC694 10"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="12"><w:name w:val="\uCABD \uBC88\uD638"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="13"><w:name w:val="\uBA38\uB9AC\uB9D0"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="14"><w:name w:val="\uAC01\uC8FC"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="15"><w:name w:val="\uBBF8\uC8FC"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="16"><w:name w:val="\uBA54\uBAA8"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="17"><w:name w:val="\uCC28\uB840 \uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="18"><w:name w:val="\uCC28\uB840 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="19"><w:name w:val="\uCC28\uB840 2"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="20"><w:name w:val="\uCC28\uB840 3"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="21"><w:name w:val="\uBCF8\uBB38 \uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="22"><w:name w:val="\uADF8\uB9BC"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="23"><w:name w:val="\uD45C"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="24"><w:name w:val="\uC218\uC2DD"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="25"><w:name w:val="\uC778\uC6A9\uBB38"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="26"><w:name w:val="\uB0A0\uC9DC"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="27"><w:name w:val="\uBC1C\uC2E0\uBA85\uC758"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="28"><w:name w:val="\uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="29"><w:name w:val="\uBD80\uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="30"><w:name w:val="\uBB38\uB2E8 \uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="31"><w:name w:val="MEMO"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="32"><w:name w:val="\uAC1C\uC694"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="33"><w:name w:val="\uD45C \uC81C\uBAA9"/><w:basedOn w:val="Normal"/></w:style>
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
function headerFooterXml(type, contents, ctx, dims) {
  const tag = type === "hdr" ? "w:hdr" : "w:ftr";
  const bodyParts = contents.map((node) => encodeContent(node, ctx, dims));
  if (contents[contents.length - 1]?.tag === "grid") {
    bodyParts.push('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
  }
  const body = bodyParts.join("\n");
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
      <w:pgSz w:w="${Metric.ptToDxa(dims.wPt)}" w:h="${Metric.ptToDxa(dims.hPt)}"${dims.orient === "landscape" ? ' w:orient="landscape"' : ""}/>
      <w:pgMar w:top="${Metric.ptToDxa(dims.mt)}" w:right="${Metric.ptToDxa(dims.mr)}" w:bottom="${Metric.ptToDxa(dims.mb)}" w:left="${Metric.ptToDxa(dims.ml)}" w:header="${Metric.ptToDxa(dims.headerPt ?? 42.52)}" w:footer="${Metric.ptToDxa(dims.footerPt ?? 42.52)}" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}
function encodeContent(node, ctx, dims) {
  return node.tag === "grid" ? encodeGrid(node, ctx, dims) : encodeParaInner(node, ctx);
}
function encodeParaInner(para, ctx, maxWidthPt) {
  const align = para.props.align;
  let headStyle = "";
  if (para.props.hwpStyleId !== void 0) {
    headStyle = `<w:pStyle w:val="${para.props.hwpStyleId}"/>`;
  } else if (para.props.heading) {
    headStyle = `<w:pStyle w:val="Heading${para.props.heading}"/>`;
  }
  let numPr = "";
  if (para.props.listOrd !== void 0) {
    const numId = para.props.listOrd ? 2 : 1;
    const ilvl = para.props.listLv ?? 0;
    numPr = `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
  }
  let spacingXml = "";
  const { spaceBefore, spaceAfter, lineHeight, lineHeightFixed, lineHeightRule } = para.props;
  if (spaceBefore !== void 0 || spaceAfter !== void 0 || lineHeight !== void 0 || lineHeightFixed !== void 0) {
    const parts = [];
    if (spaceBefore !== void 0)
      parts.push(`w:before="${Math.max(0, Metric.ptToDxa(spaceBefore))}"`);
    if (spaceAfter !== void 0)
      parts.push(`w:after="${Math.max(0, Metric.ptToDxa(spaceAfter))}"`);
    if (lineHeightFixed !== void 0) {
      parts.push(
        `w:line="${Math.max(1, Metric.ptToDxa(lineHeightFixed))}" w:lineRule="${lineHeightRule ?? "exact"}"`
      );
    } else if (lineHeight !== void 0) {
      const ratio = docxLineHeightRatio(lineHeight);
      parts.push(
        `w:line="${Math.max(1, Math.floor(ratio * 240))}" w:lineRule="auto"`
      );
    }
    spacingXml = `<w:spacing ${parts.join(" ")}/>`;
  }
  let indentXml = "";
  let leftDxa = Math.round(Metric.ptToDxa(para.props.indentPt ?? 0));
  const rightDxa = Math.round(Metric.ptToDxa(para.props.indentRightPt ?? 0));
  const firstPt = para.props.firstLineIndentPt ?? 0;
  const indParts = [];
  if (rightDxa > 0) indParts.push(`w:right="${rightDxa}"`);
  if (firstPt > 0)
    indParts.push(`w:firstLine="${Math.round(Metric.ptToDxa(firstPt))}"`);
  if (firstPt < 0) {
    const hangingDxa = Math.round(Metric.ptToDxa(-firstPt));
    if (hangingDxa > 0) {
      const baseLeftDxa = Math.max(0, leftDxa);
      leftDxa = baseLeftDxa + hangingDxa;
      if (baseLeftDxa <= 0 || hangingDxa > baseLeftDxa) {
        ctx.warns.push(
          `[DocxEncoder] w:hanging=${hangingDxa} exceeds w:left=${baseLeftDxa}`
        );
      }
      indParts.push(`w:hanging="${hangingDxa}"`);
    }
  }
  if (leftDxa > 0) indParts.unshift(`w:left="${leftDxa}"`);
  if (indParts.length > 0) indentXml = `<w:ind ${indParts.join(" ")}/>`;
  const cjkLineBreakXml = "<w:kinsoku/><w:wordWrap/><w:overflowPunct/>";
  const omitEmptyLeftAlign = align === "left" && paraTextContent(para) === "";
  const jcXml = align && !omitEmptyLeftAlign ? `<w:jc w:val="${docxJcValue(align)}"/>` : "";
  const runs = para.kids.map((k) => {
    if (k.tag === "span") return encodeRun(k, ctx);
    if (k.tag === "img") return encodeImage2(k, ctx, maxWidthPt);
    if (k.tag === "pagenum") {
      const instr = k.format === "total" ? " NUMPAGES " : " PAGE ";
      return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    }
    return "";
  }).join("");
  return `    <w:p>
      <w:pPr>${headStyle}${numPr}${spacingXml}${indentXml}${cjkLineBreakXml}${jcXml}</w:pPr>
      ${runs}
    </w:p>`;
}
function docxJcValue(align) {
  if (align === "justify") return "both";
  if (align === "distribute_space") return "distribute";
  return align;
}
function paraTextContent(para) {
  let text = "";
  const collect = (kids) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "txt") text += kid.content ?? "";
      else if (kid.kids) collect(kid.kids);
    }
  };
  collect(para.kids);
  return text;
}
function docxLineHeightRatio(lineHeight) {
  return Math.max(0.01, lineHeight);
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
      `<w:rFonts w:ascii="${esc2(p.font)}" w:hAnsi="${esc2(p.font)}" w:eastAsia="${esc2(p.font)}" w:hint="eastAsia"/>`
    );
  if (p.bg) rPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`);
  const parts = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const content = kid.content.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      if (content || rPr.length > 0) {
        parts.push(
          `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:t xml:space="preserve">${esc2(content)}</w:t></w:r>`
        );
      }
    } else if (kid.tag === "pagenum") {
      const instr = kid.format === "total" ? " NUMPAGES " : " PAGE ";
      parts.push(
        `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:instrText>${instr}</w:instrText></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:t>1</w:t></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`
      );
    } else if (kid.tag === "br") {
      parts.push(`<w:r><w:br/></w:r>`);
    } else if (kid.tag === "pb") {
      parts.push(`<w:r><w:br w:type="page"/></w:r>`);
    }
  }
  return parts.join("");
}
function encodeImage2(img, ctx, maxWidthPt) {
  const rId = ctx.imgMap.get(img);
  if (!rId) return "";
  const bodyWidthPt = Math.max(1, ctx.dims.wPt - ctx.dims.ml - ctx.dims.mr);
  const bodyHeightPt = Math.max(1, ctx.dims.hPt - ctx.dims.mt - ctx.dims.mb);
  let widthPt = Number.isFinite(img.w) && img.w > 0 ? img.w : 72;
  let heightPt = Number.isFinite(img.h) && img.h > 0 ? img.h : 72;
  const widthLimit = Math.max(1, Math.min(bodyWidthPt, maxWidthPt ?? bodyWidthPt));
  const scale = Math.min(1, widthLimit / widthPt, bodyHeightPt / heightPt);
  widthPt *= scale;
  heightPt *= scale;
  const cx = Metric.ptToEmu(widthPt);
  const cy = Metric.ptToEmu(heightPt);
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
  const distL = Metric.ptToEmu(layout.distL ?? 0);
  const distR = Metric.ptToEmu(layout.distR ?? 0);
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
function encodeGrid(grid, ctx, dims = A4, maxWidthDxa) {
  const gp = grid.props;
  const look = gp.look;
  const firstRow = look?.firstRow ? "1" : "0";
  const lastRow = look?.lastRow ? "1" : "0";
  const firstCol = look?.firstCol ? "1" : "0";
  const lastCol = look?.lastCol ? "1" : "0";
  const noHBand = look?.bandedRows ? "0" : "1";
  const noVBand = look?.bandedCols ? "0" : "1";
  const d = dims ?? A4;
  const bodyDxa = Math.max(1, Metric.ptToDxa(d.wPt - d.ml - d.mr));
  const availDxa = Math.max(1, Math.min(bodyDxa, maxWidthDxa ?? bodyDxa));
  const tablePadT = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadT ?? 1.41)));
  const tablePadB = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadB ?? 1.41)));
  const tablePadL = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadL ?? 5.1)));
  const tablePadR = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadR ?? 5.1)));
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
  const defaultColDxa = Math.max(1, Math.floor(availDxa / colCount));
  const sourceWidthsDxa = (grid.props.colWidths ?? []).map(
    (width) => width > 0 ? Metric.ptToDxa(width) : 0
  );
  const colWidthsDxa = fitColumnWidths(
    sourceWidthsDxa,
    colCount,
    availDxa,
    Math.min(100, defaultColDxa)
  );
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
            const t = cPadT ?? tablePadT;
            const b = cPadB ?? tablePadB;
            const l = cPadL ?? tablePadL;
            const r = cPadR ?? tablePadR;
            tcPrParts.push(
              `<w:tcMar><w:top w:w="${t}" w:type="dxa"/><w:left w:w="${l}" w:type="dxa"/><w:bottom w:w="${b}" w:type="dxa"/><w:right w:w="${r}" w:type="dxa"/></w:tcMar>`
            );
          }
          const parts = [];
          for (const kid of cell.kids) {
            if (kid.tag === "grid") {
              const nestedLimit = Math.max(
                1,
                cw - (cPadL ?? tablePadL) - (cPadR ?? tablePadR)
              );
              parts.push(encodeGrid(kid, ctx, dims, nestedLimit));
            } else if (kid.tag === "para") {
              const textLimit = Math.max(
                1,
                cw - (cPadL ?? tablePadL) - (cPadR ?? tablePadR)
              );
              parts.push(encodeParaInner(kid, ctx, Metric.dxaToPt(textLimit)));
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
      trPrParts.push(`<w:trHeight w:val="${hDxa}"/>`);
    }
    const trPr = trPrParts.length > 0 ? `<w:trPr>${trPrParts.join("")}</w:trPr>` : "";
    return `    <w:tr>${trPr}
${cellXmls.join("\n")}
    </w:tr>`;
  }).join("\n");
  let tblBorders = "";
  const strokeKindMap = {
    solid: "single",
    dash: "dashed",
    dot: "dotted",
    double: "double",
    none: "none",
    dashDot: "dotDash",
    dashDotDot: "dotDotDash",
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
      const minSz = val === "dashed" || val === "dotted" ? 4 : 2;
      const sz = Math.max(minSz, Math.round(s.pt * 8));
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
  const tblPosition = encodeFloatingTablePr(gp.layout);
  return `    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${Math.round(totalDxa)}" w:type="dxa"/>${tblPosition}<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="${firstRow}" w:lastRow="${lastRow}" w:firstColumn="${firstCol}" w:lastColumn="${lastCol}" w:noHBand="${noHBand}" w:noVBand="${noVBand}"/>${tblBorders}${tblJc}<w:tblCellMar><w:top w:w="${tablePadT}" w:type="dxa"/><w:left w:w="${tablePadL}" w:type="dxa"/><w:bottom w:w="${tablePadB}" w:type="dxa"/><w:right w:w="${tablePadR}" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
${rows}
    </w:tbl>`;
}
function encodeFloatingTablePr(layout) {
  if (!layout || layout.wrap === "inline") return "";
  const horzAnchorMap = {
    margin: "margin",
    page: "page",
    column: "text",
    para: "text"
  };
  const vertAnchorMap = {
    margin: "margin",
    page: "page",
    line: "text",
    para: "text"
  };
  const alignMap = {
    left: "left",
    center: "center",
    right: "right",
    top: "top",
    bottom: "bottom"
  };
  const attrs = [
    `w:leftFromText="${Math.max(0, Metric.ptToDxa(layout.distL ?? 0))}"`,
    `w:rightFromText="${Math.max(0, Metric.ptToDxa(layout.distR ?? 0))}"`,
    `w:topFromText="${Math.max(0, Metric.ptToDxa(layout.distT ?? 0))}"`,
    `w:bottomFromText="${Math.max(0, Metric.ptToDxa(layout.distB ?? 0))}"`,
    `w:vertAnchor="${vertAnchorMap[layout.vertRelTo ?? "para"] ?? "text"}"`,
    `w:horzAnchor="${horzAnchorMap[layout.horzRelTo ?? "para"] ?? "text"}"`
  ];
  if (layout.xPt != null) {
    attrs.push(`w:tblpX="${Metric.ptToDxa(layout.xPt)}"`);
  } else {
    attrs.push(`w:tblpXSpec="${alignMap[layout.horzAlign ?? "left"] ?? "left"}"`);
  }
  if (layout.yPt != null) {
    attrs.push(`w:tblpY="${Metric.ptToDxa(layout.yPt)}"`);
  } else {
    attrs.push(`w:tblpYSpec="${alignMap[layout.vertAlign ?? "top"] ?? "top"}"`);
  }
  return `<w:tblpPr ${attrs.join(" ")}/><w:tblOverlap w:val="overlap"/>`;
}
function encodeCellBorders(cp) {
  if (!cp.top && !cp.bot && !cp.left && !cp.right) return "";
  const strokeKindMap = {
    solid: "single",
    dash: "dashed",
    dot: "dotted",
    double: "double",
    none: "none",
    dashDot: "dotDash",
    dashDotDot: "dotDotDash",
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
    const minSz = val === "dashed" || val === "dotted" ? 4 : 2;
    const sz = Math.max(minSz, Math.round(s.pt * 8));
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
  let xmlSafe = "";
  for (const ch of s) {
    const cp = ch.codePointAt(0);
    if (cp === 9 || cp === 10 || cp === 13 || cp !== void 0 && cp >= 32 && cp <= 55295 || cp !== void 0 && cp >= 57344 && cp <= 65533 || cp !== void 0 && cp >= 65536 && cp <= 1114111) {
      xmlSafe += ch;
    }
  }
  return TextKit.escapeXml(xmlSafe);
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
    if (k.tag === "link") {
      const label = k.kids.map((span) => encodeSpan(span, warns)).join("");
      return `[${label}](${k.href})`;
    }
    if (k.tag === "pagenum") {
      warnOnce(warns, "[SHIELD] MD: \uD398\uC774\uC9C0 \uBC88\uD638 \uD45C\uD604 \uBD88\uAC00 \u2014 \uC790\uB9AC\uD45C\uC2DC\uC790\uB85C \uB300\uCCB4\uB428");
      return "[\uD398\uC774\uC9C0 \uBC88\uD638]";
    }
    return "";
  }).join("");
  if (para.props.heading) return `${"#".repeat(para.props.heading)} ${text}`;
  if (para.props.listOrd !== void 0) {
    const indent = "  ".repeat(para.props.listLv ?? 0);
    const sourceMark = para.props.listMark;
    const marker = para.props.listOrd ? sourceMark && /^\d+\.$/.test(sourceMark) ? sourceMark : "1." : sourceMark && /^[-*+]$/.test(sourceMark) ? sourceMark : "-";
    return `${indent}${marker} ${text}`;
  }
  if (para.props.align && para.props.align !== "left" && para.props.align !== "justify") {
    warnOnce(warns, "[SHIELD] MD: \uBB38\uB2E8 \uC815\uB82C \uD45C\uD604 \uBD88\uAC00 \u2014 \uC815\uB82C \uC815\uBCF4\uAC00 \uC190\uC2E4\uB428");
  }
  return text;
}
function warnOnce(warns, warning) {
  if (!warns.includes(warning)) warns.push(warning);
}
function isCodeFont(font) {
  return !!font && /courier|consolas|monaco|menlo|monospace/i.test(font);
}
function wrapInlineCode(text) {
  const longestRun = Math.max(
    0,
    ...(text.match(/`+/g) ?? []).map((run) => run.length)
  );
  const fence = "`".repeat(longestRun + 1);
  const pad = text.startsWith("`") || text.endsWith("`") ? " " : "";
  return `${fence}${pad}${text}${pad}${fence}`;
}
function encodeSpan(span, warns) {
  let hasPageNum = false;
  const textParts = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") textParts.push(kid.content);
    else if (kid.tag === "br") textParts.push("  \n");
    else if (kid.tag === "pb") {
      warnOnce(warns, "[SHIELD] MD: \uCABD \uB098\uB204\uAE30 \uD45C\uD604 \uBD88\uAC00 \u2014 \uC190\uC2E4\uB428");
    } else if (kid.tag === "pagenum") {
      hasPageNum = true;
      warnOnce(warns, "[SHIELD] MD: \uD398\uC774\uC9C0 \uBC88\uD638 \uD45C\uD604 \uBD88\uAC00 \u2014 \uC790\uB9AC\uD45C\uC2DC\uC790\uB85C \uB300\uCCB4\uB428");
    }
  }
  let r = textParts.join("");
  if (hasPageNum && r === "") r = "[\uD398\uC774\uC9C0 \uBC88\uD638]";
  const code = isCodeFont(span.props.font);
  if (span.props.font && !code) {
    warnOnce(warns, "[SHIELD] MD: \uAE00\uAF34\uBA85 \uD45C\uD604 \uBD88\uAC00 \u2014 \uAE00\uAF34 \uC815\uBCF4\uAC00 \uC190\uC2E4\uB428");
  }
  if (span.props.pt !== void 0) {
    warnOnce(warns, "[SHIELD] MD: \uAE00\uC790 \uD06C\uAE30 \uD45C\uD604 \uBD88\uAC00 \u2014 \uD06C\uAE30 \uC815\uBCF4\uAC00 \uC190\uC2E4\uB428");
  }
  if (span.props.color || span.props.bg) {
    warnOnce(warns, "[SHIELD] MD: \uAE00\uC790\uC0C9/\uBC30\uACBD\uC0C9 \uD45C\uD604 \uBD88\uAC00 \u2014 \uC0C9\uC0C1 \uC815\uBCF4\uAC00 \uC190\uC2E4\uB428");
  }
  if (span.props.u || span.props.sup || span.props.sub) {
    warnOnce(warns, "[SHIELD] MD: \uBC11\uC904/\uC704\uCCA8\uC790/\uC544\uB798\uCCA8\uC790 \uD45C\uD604 \uBD88\uAC00 \u2014 \uD574\uB2F9 \uC11C\uC2DD\uC774 \uC190\uC2E4\uB428");
  }
  if (code) r = wrapInlineCode(r);
  if (span.props.b && span.props.i) r = `***${r}***`;
  else if (span.props.b) r = `**${r}**`;
  else if (span.props.i) r = `*${r}*`;
  if (span.props.s) r = `~~${r}~~`;
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
  if (canEncodePipeTable(grid)) {
    const losesLayout = Object.keys(grid.props).length > 0 || grid.kids.some(
      (row) => row.heightPt !== void 0 || row.kids.some((cell) => Object.keys(cell.props).length > 0)
    );
    if (losesLayout) {
      warnOnce(
        warns,
        "[SHIELD] MD: \uD30C\uC774\uD504 \uD45C\uAC00 \uC9C0\uC6D0\uD558\uC9C0 \uC54A\uB294 \uB108\uBE44/\uD14C\uB450\uB9AC/\uBC30\uACBD/\uC815\uB82C \uC815\uBCF4\uAC00 \uC190\uC2E4\uB428"
      );
    }
    return encodePipeTable(grid, warns, includeImages);
  }
  warnOnce(
    warns,
    "[SHIELD] MD: \uBCD1\uD569 \uC140 \uB610\uB294 \uC140 \uB0B4\uBD80 \uAC1C\uD589/\uBE14\uB85D \uC694\uC18C \uB54C\uBB38\uC5D0 HTML \uD45C\uB85C \uD3F4\uBC31\uD568"
  );
  return encodeHtmlTable(grid, warns, includeImages);
}
function paraHasLineBreak(para) {
  return para.kids.some((kid) => {
    if (kid.tag === "grid") return true;
    if (kid.tag === "span") {
      return kid.kids.some(
        (child) => child.tag === "br" || child.tag === "pb" || child.tag === "txt" && /[\r\n]/.test(child.content)
      );
    }
    if (kid.tag === "link") {
      return kid.kids.some((span) => span.kids.some(
        (child) => child.tag === "br" || child.tag === "pb" || child.tag === "txt" && /[\r\n]/.test(child.content)
      ));
    }
    return false;
  });
}
function canEncodePipeTable(grid) {
  const columns = grid.kids[0]?.kids.length ?? 0;
  if (columns === 0) return false;
  return grid.kids.every(
    (row) => row.kids.length === columns && row.kids.every(
      (cell) => cell.cs === 1 && cell.rs === 1 && cell.kids.length === 1 && cell.kids[0].tag === "para" && cell.kids[0].props.heading === void 0 && cell.kids[0].props.listOrd === void 0 && !paraHasLineBreak(cell.kids[0])
    )
  );
}
function encodePipeTable(grid, warns, includeImages) {
  const rows = grid.kids.map((row) => row.kids.map((cell) => {
    const para = cell.kids[0];
    return encodePara(para, warns, includeImages).replace(/\|/g, "\\|");
  }));
  const renderRow = (cells) => `| ${cells.join(" | ")} |`;
  const separator = renderRow(rows[0].map(() => "---"));
  return [renderRow(rows[0]), separator, ...rows.slice(1).map(renderRow)].join("\n");
}
function encodeHtmlTable(grid, warns, includeImages) {
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
var TAG_TAB_DEF = T + 6;
var TAG_NUMBERING2 = T + 7;
var TAG_BULLET2 = T + 8;
var TAG_PARA_SHAPE2 = T + 9;
var TAG_STYLE2 = T + 10;
var TAG_DOC_DATA = T + 11;
var TAG_COMPATIBLE_DOCUMENT = T + 14;
var TAG_LAYOUT_COMPATIBILITY = T + 15;
var TAG_PARA_HEADER2 = T + 50;
var TAG_PARA_TEXT2 = T + 51;
var TAG_PARA_CHAR_SHAPE2 = T + 52;
var TAG_PARA_LINE_SEG = T + 53;
var TAG_CTRL_HEADER2 = T + 55;
var TAG_LIST_HEADER2 = T + 56;
var TAG_PAGE_DEF2 = T + 57;
var TAG_FOOTNOTE_SHAPE = T + 58;
var TAG_PAGE_BORDER_FILL = T + 59;
var TAG_SHAPE_COMPONENT = T + 60;
var TAG_TABLE = T + 61;
var TAG_SHAPE_COMPONENT_PICTURE2 = T + 69;
var CTRL_TABLE2 = 1952607264;
var CTRL_SECD2 = 1936024420;
var CTRL_COLD = 1668246628;
var CTRL_GSO2 = 1735618336;
var CTRL_PIC2 = 611346787;
var CTRL_FIELD_BEGIN = 1684825637;
var CTRL_FIELD_END = 1684825692;
var TABLE_CTRL_MASK = 1 << 11;
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
  solid: 1,
  dot: 3,
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
  f64(v) {
    const b = new Uint8Array(8);
    new DataView(b.buffer).setFloat64(0, v, true);
    this.chunks.push(b);
    this._sz += 8;
    return this;
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
  const isLarge = sz >= 4095;
  const enc = isLarge ? 4095 : sz;
  const hdr = ((enc & 4095) * 1048576 | (level & 1023) << 10 | tag & 1023) >>> 0;
  const w = new BufWriter().u32(hdr);
  if (isLarge) w.u32(sz);
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
    this.NONE_STROKE = { kind: "none", pt: 0, color: "000000" };
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
    this.maxStyleId = 0;
    this.styleParaShapeIds = /* @__PURE__ */ new Map([[0, 0]]);
    this.hasNumbering = false;
    this.hasBullet = false;
    // charShape마다 언어별 fontId를 기록
    this.csFontIds = [[0, 0, 0, 0, 0, 0, 0]];
    for (const g of LANG_GROUPS2) this._registerLangFont(g, "\uD568\uCD08\uB86C\uBC14\uD0D5");
    this.addBorderFill(this.NONE_STROKE);
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
  registerStyleId(styleId, paraShapeId) {
    if (styleId === void 0) return;
    if (!Number.isInteger(styleId) || styleId < 0 || styleId > 255) return;
    this.maxStyleId = Math.max(this.maxStyleId, styleId);
    if (paraShapeId !== void 0 && !this.styleParaShapeIds.has(styleId)) {
      this.styleParaShapeIds.set(styleId, paraShapeId);
    }
  }
  getStyleCount() {
    return this.maxStyleId + 1;
  }
  getStyleParaShapeId(styleId) {
    return this.styleParaShapeIds.get(styleId) ?? 0;
  }
  registerList(p) {
    if (p.listOrd === true) this.hasNumbering = true;
    if (p.listOrd === false) this.hasBullet = true;
  }
  getNumberingCount() {
    return this.hasNumbering ? 1 : 0;
  }
  getBulletCount() {
    return this.hasBullet ? 1 : 0;
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
    p.heading ?? 0,
    p.listOrd === void 0 ? "" : p.listOrd ? "number" : "bullet",
    p.listLv ?? 0,
    p.indentPt ?? 0,
    p.indentRightPt ?? 0,
    p.firstLineIndentPt ?? 0,
    p.spaceBefore ?? 0,
    p.spaceAfter ?? 0,
    p.lineHeight ?? 1,
    p.lineHeightFixed ?? 0
  ].join("|");
}
function hwpStyleIdForPara(p) {
  if (p.hwpStyleId !== void 0) {
    const id = Math.trunc(p.hwpStyleId);
    return id >= 0 && id <= 255 ? id : void 0;
  }
  if (p.styleId !== void 0) {
    const id = Number(p.styleId);
    if (Number.isInteger(id) && id >= 0 && id <= 255) return id;
  }
  if (p.heading !== void 0) return p.heading + 1;
  return void 0;
}
function bfKey(s, bg) {
  return `${s.kind}|${s.pt}|${s.color}|${bg ?? ""}`;
}
function bfPerSideKey(l, r, t, b, bg) {
  return `${bfKey(l)}/${bfKey(r)}/${bfKey(t)}/${bfKey(b)}/${bg ?? ""}`;
}
function collectNode(node, bank) {
  if (node.tag === "para") {
    const paraShapeId = bank.addParaShape(node.props);
    bank.registerStyleId(hwpStyleIdForPara(node.props), paraShapeId);
    bank.registerList(node.props);
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
  w.u32(1);
  w.u32(bank.getNumberingCount());
  w.u32(bank.getBulletCount());
  w.u32(bank.psProps.length);
  w.u32(bank.getStyleCount());
  w.u32(0);
  w.u32(0);
  w.u32(0);
  return w.build();
}
function mkStyle(name, engName, paraPrId, charPrId, nextStyleId2 = 0) {
  return new BufWriter().u16(name.length).utf16(name).u16(engName.length).utf16(engName).u8(0).u8(nextStyleId2).i16(1042).u16(paraPrId).u16(charPrId).u16(0).build();
}
function mkTabDef() {
  return new Uint8Array(8);
}
function writeNumberingLevel(writer, attr, format) {
  writer.u32(attr).u16(0).u16(50).u32(4294967295).u16(format.length).utf16(format);
}
function mkNumbering() {
  const writer = new BufWriter();
  const levels = [
    [12, "^1."],
    [268, "^2."],
    [12, "^3)"],
    [268, "^4)"],
    [12, "(^5)"],
    [268, "(^6)"],
    [44, "^7"]
  ];
  for (const [attr, format] of levels) {
    writeNumberingLevel(writer, attr, format);
  }
  writer.u16(0);
  for (let level = 0; level < 7; level++) writer.u32(1);
  writeNumberingLevel(writer, 300, "^8");
  writeNumberingLevel(writer, 332, "");
  writeNumberingLevel(writer, 108, "");
  for (let level = 0; level < 3; level++) writer.u32(1);
  return writer.build();
}
function mkBullet() {
  return new BufWriter().u32(12).u16(0).u16(50).u16(8226).i32(0).u8(0).u8(0).u8(0).u8(0).u16(0).build();
}
var HWP_STYLE_NAMES = [
  ["\uBC14\uD0D5\uAE00", "Normal"],
  ["\uBCF8\uBB38", "Body"],
  ["\uAC1C\uC694 1", "Outline 1"],
  ["\uAC1C\uC694 2", "Outline 2"],
  ["\uAC1C\uC694 3", "Outline 3"],
  ["\uAC1C\uC694 4", "Outline 4"],
  ["\uAC1C\uC694 5", "Outline 5"],
  ["\uAC1C\uC694 6", "Outline 6"],
  ["\uAC1C\uC694 7", "Outline 7"],
  ["\uAC1C\uC694 8", "Outline 8"],
  ["\uAC1C\uC694 9", "Outline 9"],
  ["\uAC1C\uC694 10", "Outline 10"],
  ["\uCABD \uBC88\uD638", "Page Number"],
  ["\uBA38\uB9AC\uB9D0", "Header"],
  ["\uAC01\uC8FC", "Footnote"],
  ["\uBBF8\uC8FC", "Endnote"],
  ["\uBA54\uBAA8", "Memo"],
  ["\uCC28\uB840 \uC81C\uBAA9", "TOC Heading"],
  ["\uCC28\uB840 1", "TOC 1"],
  ["\uCC28\uB840 2", "TOC 2"],
  ["\uCC28\uB840 3", "TOC 3"],
  ["\uCEA1\uC158", "Caption"],
  ["\uADF8\uB9BC", "Figure"],
  ["\uD45C", "Table"],
  ["\uC218\uC2DD", "Equation"],
  ["\uC778\uC6A9\uBB38", "Quote"],
  ["\uB0A0\uC9DC", "Date"],
  ["\uBC1C\uC2E0\uBA85\uC758", "Sender"],
  ["\uC81C\uBAA9", "Title"],
  ["\uBD80\uC81C\uBAA9", "Subtitle"],
  ["\uBB38\uB2E8 \uC81C\uBAA9", "Paragraph Title"]
];
function styleNameForId(id) {
  return HWP_STYLE_NAMES[id] ?? [`\uC2A4\uD0C0\uC77C ${id}`, `Style ${id}`];
}
function basicFaceNameFor(face) {
  const normalized = face.trim();
  const aliases = {
    "\uB9D1\uC740 \uACE0\uB515": "Malgun Gothic",
    "\uB9D1\uC740\uACE0\uB515": "Malgun Gothic",
    "\uD568\uCD08\uB86C\uBC14\uD0D5": "HCR Batang",
    "\uD55C\uCEF4\uBC14\uD0D5": "HCR Batang",
    "\uD568\uCD08\uB86C\uB3CB\uC6C0": "HCR Dotum",
    "\uD55C\uCEF4\uB3CB\uC6C0": "HCR Dotum",
    "\uBC14\uD0D5": "Batang",
    "\uB3CB\uC6C0": "Dotum",
    "\uAD74\uB9BC": "Gulim",
    "\uAD81\uC11C": "Gungsuh",
    "\uD55C\uC591\uC2E0\uBA85\uC870": "HY Sinmyeongjo",
    "HY\uC2E0\uBA85\uC870": "HY Sinmyeongjo",
    "\uD55C\uC591\uACAC\uACE0\uB515": "HY Gyeongothic",
    "HY\uACAC\uACE0\uB515": "HY Gyeongothic",
    "\uD55C\uC591\uC911\uACE0\uB515": "HY Junggothic",
    "HY\uC911\uACE0\uB515": "HY Junggothic",
    "HY\uD5E4\uB4DC\uB77C\uC778M": "HYHeadLine-Medium"
  };
  return aliases[normalized] ?? normalized;
}
function faceTypeInfo(face) {
  if (/바탕|명조|Batang|Myeong|Sinmyeong|Serif/i.test(face)) {
    return [2, 3, 6, 4, 0, 1, 1, 1, 1, 1];
  }
  if (/고딕|돋움|Gothic|Dotum|HeadLine|헤드라인/i.test(face)) {
    return [2, 11, 5, 3, 2, 0, 0, 2, 0, 4];
  }
  return [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
}
function mkFaceName(name) {
  const basic = basicFaceNameFor(name);
  const w = new BufWriter().u8(97).u16(name.length).utf16(name);
  for (const b of faceTypeInfo(name)) w.u8(b);
  return w.u16(basic.length).utf16(basic).build();
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
  for (let i = 0; i < 4; i++) w.u8(t).u8(wi).colorRef(col);
  w.u8(0).u8(0).colorRef("000000");
  if (bg) {
    w.u32(1).colorRef(bg).colorRef(bg).u32(4294967295).u32(0).u8(0);
  } else {
    w.u32(0).u32(0);
  }
  return w.build();
}
function mkBorderFillPerSide(l, r, t, b, bg) {
  const w = new BufWriter();
  w.u16(0);
  w.u8(BORDER_KIND_IDX[l.kind] ?? 0).u8(borderWidthIdx(l.pt)).colorRef(l.color || "000000");
  w.u8(BORDER_KIND_IDX[r.kind] ?? 0).u8(borderWidthIdx(r.pt)).colorRef(r.color || "000000");
  w.u8(BORDER_KIND_IDX[t.kind] ?? 0).u8(borderWidthIdx(t.pt)).colorRef(t.color || "000000");
  w.u8(BORDER_KIND_IDX[b.kind] ?? 0).u8(borderWidthIdx(b.pt)).colorRef(b.color || "000000");
  w.u8(0).u8(0).colorRef("000000");
  if (bg) {
    w.u32(1).colorRef(bg).colorRef(bg).u32(4294967295).u32(0).u8(0);
  } else {
    w.u32(0).u32(0);
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
  const lineSpacingType = p.lineHeightFixed !== void 0 ? 3 : 0;
  let attr1 = lineSpacingType & 3 | (alignVal & 7) << 2;
  if (p.heading !== void 0) {
    attr1 |= 1 << 23;
    attr1 |= p.heading - 1 << 25;
  } else if (p.listOrd !== void 0) {
    attr1 |= (p.listOrd ? 2 : 3) << 23;
    attr1 |= Math.max(0, Math.min(6, p.listLv ?? 0)) << 25;
  }
  const lineSpaceValue = p.lineHeightFixed !== void 0 ? Math.max(
    Metric.ptToHwp(p.lineHeightFixed) * 2,
    Math.ceil(Metric.ptToHwp(10) * 1.15) * 2
  ) : Math.max(100, p.lineHeight ? Math.round(p.lineHeight * 100) : 160);
  const paraShapeUnit = (pt) => Metric.ptToHwp(pt) * 2;
  return new BufWriter().u32(attr1).i32(paraShapeUnit(p.indentPt ?? 0)).i32(paraShapeUnit(p.indentRightPt ?? 0)).i32(paraShapeUnit(p.firstLineIndentPt ?? 0)).i32(paraShapeUnit(p.spaceBefore ?? 0)).i32(paraShapeUnit(p.spaceAfter ?? 0)).i32(lineSpaceValue).u16(0).u16(p.listOrd === void 0 ? 0 : 1).u16(0).i16(0).i16(0).i16(0).i16(0).u32(0).u32(lineSpacingType).u32(lineSpaceValue).u32(0).build();
}
function mkBinData(id, ext) {
  return new BufWriter().u16(33).u16(id).u16(ext.length).utf16(ext).build();
}
function buildDocInfoStream(bank, images = []) {
  const chunks = [];
  chunks.push(mkRec(TAG_DOCUMENT_PROPERTIES, 0, mkDocumentProperties()));
  chunks.push(mkRec(TAG_ID_MAPPINGS, 0, mkIdMappings(bank, images.length)));
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
  chunks.push(mkRec(TAG_TAB_DEF, 1, mkTabDef()));
  if (bank.getNumberingCount() > 0) {
    chunks.push(mkRec(TAG_NUMBERING2, 1, mkNumbering()));
  }
  if (bank.getBulletCount() > 0) {
    chunks.push(mkRec(TAG_BULLET2, 1, mkBullet()));
  }
  for (const p of bank.psProps) {
    chunks.push(mkRec(TAG_PARA_SHAPE2, 1, mkParaShape(p)));
  }
  for (let i = 0; i < bank.getStyleCount(); i++) {
    const [name, engName] = styleNameForId(i);
    chunks.push(
      mkRec(
        TAG_STYLE2,
        1,
        mkStyle(
          name,
          engName,
          bank.getStyleParaShapeId(i),
          0,
          i === 0 ? 0 : i
        )
      )
    );
  }
  chunks.push(mkRec(TAG_COMPATIBLE_DOCUMENT, 0, new Uint8Array(4)));
  chunks.push(mkRec(TAG_LAYOUT_COMPATIBILITY, 1, new Uint8Array(20)));
  return concatU8(chunks);
}
function mkPageDef(dims) {
  const rawTopPt = dims.mt;
  const rawBottomPt = dims.mb;
  const rawHeaderPt = dims.headerPt ?? 0;
  const rawFooterPt = dims.footerPt ?? 0;
  return new BufWriter().u32(Metric.ptToHwp(dims.wPt)).u32(Metric.ptToHwp(dims.hPt)).u32(Metric.ptToHwp(dims.ml)).u32(Metric.ptToHwp(dims.mr)).u32(Metric.ptToHwp(rawTopPt)).u32(Metric.ptToHwp(rawBottomPt)).u32(Metric.ptToHwp(rawHeaderPt)).u32(Metric.ptToHwp(rawFooterPt)).u32(0).u32(dims.orient === "landscape" ? 1 : 0).build();
}
function mkParaHeader(nchars, ctrlMask, psId, csCount, lineAlignCount = 0, instanceId = 0, styleId = 0, divideSort = 0) {
  return new BufWriter().u32(nchars >>> 0 | 2147483648).u32(ctrlMask).u16(psId).u8(Math.max(0, Math.min(255, Math.trunc(styleId)))).u8(Math.max(0, Math.min(255, Math.trunc(divideSort)))).u16(csCount).u16(0).u16(lineAlignCount).u32(instanceId).u16(0).build();
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
      return textHeight + value;
    case 3:
      return Math.max(textHeight, value);
    case 4:
      return Math.floor(textHeight * value);
    default:
      return Math.floor(textHeight * value / 100);
  }
}
function maxFontHwpInPara(para) {
  let maxHwp = 1e3;
  const visit = (kids) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span" && typeof kid.props?.pt === "number" && kid.props.pt > 0) {
        maxHwp = Math.max(maxHwp, Metric.ptToHwp(kid.props.pt));
      }
      if (kid.kids) visit(kid.kids);
    }
  };
  visit(para.kids);
  return maxHwp;
}
function extractParaLayoutText(para) {
  let text = "";
  const visit = (kids) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span") {
        for (const child of kid.kids ?? []) {
          if (child.tag === "txt") text += child.content ?? "";
          else if (child.tag === "br") text += "\n";
        }
      } else if (kid.tag === "link") {
        visit(kid.kids ?? []);
      }
    }
  };
  visit(para.kids);
  return text;
}
function estimateCharWidthHwp(code, fontHwp) {
  if (code >= 44032 && code <= 55203) return fontHwp;
  if (code >= 12592 && code <= 12687) return fontHwp;
  if (code >= 19968 && code <= 40959) return fontHwp;
  if (code >= 65 && code <= 90) return Math.round(fontHwp * 0.65);
  if (code === 32) return Math.round(fontHwp * 0.32);
  if (code > 255) return fontHwp;
  return Math.round(fontHwp * 0.42);
}
function lineStartPositionsHwp(text, fontHwp, availWidthHwp) {
  if (!text) return [0];
  const maxWidth = Math.max(1, availWidthHwp ?? 0);
  if (!availWidthHwp || availWidthHwp <= 0) {
    const starts2 = [0];
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (code !== 10 && code !== 13) continue;
      if (code === 13 && text.charCodeAt(i + 1) === 10) i++;
      starts2.push(i + 1);
    }
    return starts2;
  }
  const starts = [0];
  let currentWidth = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 10 || code === 13) {
      if (code === 13 && text.charCodeAt(i + 1) === 10) i++;
      starts.push(i + 1);
      currentWidth = 0;
      continue;
    }
    if (code < 32) continue;
    const charWidth = estimateCharWidthHwp(code, fontHwp);
    if (currentWidth > 0 && currentWidth + charWidth > maxWidth) {
      starts.push(i);
      currentWidth = charWidth;
    } else {
      currentWidth += charWidth;
    }
  }
  return starts;
}
function estimateLineCountHwp(text, fontHwp, availWidthHwp) {
  return lineStartPositionsHwp(text, fontHwp, availWidthHwp).length;
}
function safeParaLineAdvanceHwp(props, fontHwp) {
  const textHeight = Math.max(1, fontHwp);
  if (props?.lineHeightFixed !== void 0) {
    const fixed = Math.max(0, Metric.ptToHwp(props.lineHeightFixed));
    return calcLineHeight(
      3,
      Math.max(fixed, Math.ceil(textHeight * 1.15)),
      textHeight
    );
  }
  const ratio = Math.max(100, Math.round((props?.lineHeight ?? 1.6) * 100));
  return Math.max(calcLineHeight(0, ratio, textHeight), textHeight);
}
function minParaHeightHwp(para, availWidthHwp) {
  const fontHwp = maxFontHwpInPara(para);
  const lineCount = estimateLineCountHwp(
    extractParaLayoutText(para),
    fontHwp,
    availWidthHwp
  );
  return Metric.ptToHwp(Math.max(0, para.props.spaceBefore ?? 0)) + safeParaLineAdvanceHwp(para.props, fontHwp) * lineCount + Metric.ptToHwp(Math.max(0, para.props.spaceAfter ?? 0));
}
function nextParaVertPosHwp(para, vertPos, availWidthHwp, pageBodyHeightHwp) {
  const fontHwp = maxFontHwpInPara(para);
  const lineCount = estimateLineCountHwp(
    extractParaLayoutText(para),
    fontHwp,
    availWidthHwp
  );
  const lineAdvance = safeParaLineAdvanceHwp(para.props, fontHwp);
  let next = vertPos + Metric.ptToHwp(Math.max(0, para.props.spaceBefore ?? 0));
  for (let i = 0; i < lineCount; i++) {
    if (pageBodyHeightHwp !== void 0 && next > 0 && next + lineAdvance > pageBodyHeightHwp) {
      next = 0;
    }
    next += lineAdvance;
  }
  return next + Metric.ptToHwp(Math.max(0, para.props.spaceAfter ?? 0));
}
function minGridHeightHwp(grid) {
  return grid.kids.reduce((sum, row) => {
    const explicit = row.heightPt != null && row.heightPt > 0 ? Metric.ptToHwp(row.heightPt) : Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT);
    let minRow = 0;
    for (const cell of row.kids ?? []) {
      const span = Math.max(1, cell.rs ?? 1);
      minRow = Math.max(minRow, Math.ceil(minCellHeightHwp(cell) / span));
    }
    return sum + Math.max(explicit, minRow);
  }, 0);
}
function minCellHeightHwp(cell, innerWidthHwp) {
  const cp = cell.props ?? {};
  const padT = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : 141;
  const padB = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : 141;
  let content = 0;
  for (const kid of cell.kids ?? []) {
    if (kid.tag === "para") {
      const images = flatImgNodes(kid.kids);
      for (const image of images) {
        content += imageDisplaySizeHwp(
          image,
          innerWidthHwp ?? Number.MAX_SAFE_INTEGER
        ).h;
      }
      const textKids = kid.kids.filter((child) => child.tag !== "img");
      if (textKids.length > 0 || images.length === 0) {
        content += minParaHeightHwp({ ...kid, kids: textKids }, innerWidthHwp);
      }
    } else if (kid.tag === "grid")
      content += minGridHeightHwp(kid) + Metric.ptToHwp(6);
  }
  return Math.max(content, 1e3) + padT + padB;
}
function mkLineSeg(textStartPos, vertPos, vertSize, textHeight, baseline, spacing, horzPos, horzSize, flags) {
  return new BufWriter().u32(textStartPos).i32(vertPos).i32(vertSize).i32(textHeight).i32(baseline).i32(spacing).i32(horzPos).i32(horzSize).u32(flags).build();
}
function buildDefaultLineSeg(availWidthHwp, fontHwp, _nchars, paraProps, vertPos = 0) {
  const lineAdvance = safeParaLineAdvanceHwp(paraProps, fontHwp);
  const vertSize = fontHwp;
  const baseline = Math.round(fontHwp * 0.85);
  const spacing = Math.max(0, lineAdvance - vertSize);
  const hasIndent = (paraProps?.firstLineIndentPt ?? 0) !== 0;
  const flags = 393216 | (hasIndent ? 1048576 : 0);
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
function buildObjectLineSeg(availWidthHwp, objectHeightHwp, vertPos = 0) {
  const vertSize = Math.max(1, objectHeightHwp);
  const spacing = Math.max(0, safeParaLineAdvanceHwp(void 0, 1e3) - 1e3);
  return mkLineSeg(
    0,
    vertPos,
    vertSize,
    vertSize,
    Math.round(vertSize * 0.85),
    spacing,
    0,
    availWidthHwp,
    393216
  );
}
function buildParaLineSegs(text, availWidthHwp, fontHwp, paraProps, vertPos, textPosOffset = 0, pageBodyHeightHwp) {
  const starts = lineStartPositionsHwp(text, fontHwp, availWidthHwp);
  const lineAdvance = safeParaLineAdvanceHwp(paraProps, fontHwp);
  const vertSize = fontHwp;
  const spacing = Math.max(0, lineAdvance - vertSize);
  const baseline = Math.round(fontHwp * 0.85);
  const hasIndent = (paraProps?.firstLineIndentPt ?? 0) !== 0;
  let nextLineVertPos = vertPos + Metric.ptToHwp(Math.max(0, paraProps?.spaceBefore ?? 0));
  const segments = starts.map((start, index) => {
    if (pageBodyHeightHwp !== void 0 && nextLineVertPos > 0 && nextLineVertPos + lineAdvance > pageBodyHeightHwp) {
      nextLineVertPos = 0;
    }
    const lineVertPos = nextLineVertPos;
    nextLineVertPos += lineAdvance;
    const flags = 393216 | (index === 0 && hasIndent ? 1048576 : 0);
    return mkLineSeg(
      index === 0 ? 0 : textPosOffset + start,
      lineVertPos,
      vertSize,
      fontHwp,
      baseline,
      spacing,
      0,
      availWidthHwp,
      flags
    );
  });
  return {
    data: concatU8(segments),
    count: segments.length,
    totalHeight: Metric.ptToHwp(Math.max(0, paraProps?.spaceBefore ?? 0)) + segments.length * lineAdvance + Metric.ptToHwp(Math.max(0, paraProps?.spaceAfter ?? 0))
  };
}
function mkSectionAndColumnParaTextPrefix() {
  const secdLo = CTRL_SECD2 & 65535;
  const secdHi = CTRL_SECD2 >>> 16 & 65535;
  const coldLo = CTRL_COLD & 65535;
  const coldHi = CTRL_COLD >>> 16 & 65535;
  return new BufWriter().u16(2).u16(secdLo).u16(secdHi).u16(0).u16(0).u16(0).u16(0).u16(2).u16(2).u16(coldLo).u16(coldHi).u16(0).u16(0).u16(0).u16(0).u16(2).build();
}
function mkTableParaText() {
  const lo = CTRL_TABLE2 & 65535;
  const hi = CTRL_TABLE2 >>> 16 & 65535;
  return new BufWriter().u16(11).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(11).u16(13).build();
}
function mkPicParaText() {
  const lo = CTRL_GSO2 & 65535;
  const hi = CTRL_GSO2 >>> 16 & 65535;
  return new BufWriter().u16(11).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(11).u16(13).build();
}
function imageDisplaySizeHwp(imgNode, maxWidthHwp, maxHeightHwp = Number.MAX_SAFE_INTEGER) {
  const rawData = TextKit.base64Decode(imgNode.b64);
  const pixDims = readPixelDims2(rawData, imgNode.mime);
  const sourceW = pixDims?.w ? Metric.ptToHwp(pixDims.w * 72 / 96) : Metric.ptToHwp(72);
  const sourceH = pixDims?.h ? Metric.ptToHwp(pixDims.h * 72 / 96) : Metric.ptToHwp(72);
  let w = Number.isFinite(imgNode.w) && imgNode.w > 0 ? Metric.ptToHwp(imgNode.w) : sourceW;
  let h = Number.isFinite(imgNode.h) && imgNode.h > 0 ? Metric.ptToHwp(imgNode.h) : sourceH;
  const scale = Math.min(
    1,
    Math.max(1, maxWidthHwp) / Math.max(1, w),
    Math.max(1, maxHeightHwp) / Math.max(1, h)
  );
  w = Math.max(1, Math.round(w * scale));
  h = Math.max(1, Math.round(h * scale));
  return { w, h };
}
function writeIdentityMatrix(w) {
  w.f64(1).f64(0).f64(0).f64(0).f64(1).f64(0);
}
function mkShapeComponent(wHwp, hHwp) {
  const w = new BufWriter().u32(CTRL_PIC2).u32(CTRL_PIC2).i32(0).i32(0).u16(0).u16(1).u32(wHwp).u32(hHwp).u32(wHwp).u32(hHwp).u32(0).u16(0).i32(Math.round(wHwp / 2)).i32(Math.round(hHwp / 2)).u16(1);
  writeIdentityMatrix(w);
  writeIdentityMatrix(w);
  writeIdentityMatrix(w);
  return w.build();
}
function mkShapeComponentPicture(binDataId, wHwp, hHwp, instanceId) {
  return new BufWriter().u32(0).i32(0).u32(0).i32(0).i32(0).i32(wHwp).i32(0).i32(wHwp).i32(hHwp).i32(0).i32(hHwp).i32(0).i32(0).i32(wHwp).i32(hHwp).u16(0).u16(0).u16(0).u16(0).u8(0).u8(0).u8(0).u16(binDataId).u8(0).u32(instanceId).u32(0).u32(wHwp).u32(hHwp).u8(0).build();
}
function mkObjectCtrl(ctrlId, wHwp, hHwp, instanceId, layout) {
  let attr = 136978960;
  if (!layout || layout.wrap === "inline") attr |= 1 | 1 << 2;
  return new BufWriter().u32(ctrlId).u32(attr).i32(layout?.yPt ? Metric.ptToHwp(layout.yPt) : 0).i32(layout?.xPt ? Metric.ptToHwp(layout.xPt) : 0).u32(wHwp).u32(hHwp).i32(layout?.zOrder ?? 0).u16(layout?.distL ? Metric.ptToHwp(layout.distL) : 0).u16(layout?.distR ? Metric.ptToHwp(layout.distR) : 0).u16(layout?.distT ? Metric.ptToHwp(layout.distT) : 0).u16(layout?.distB ? Metric.ptToHwp(layout.distB) : 0).u32(instanceId).i32(0).u16(0).build();
}
function mkFieldBeginCtrl(instanceId) {
  return new BufWriter().u32(CTRL_FIELD_BEGIN).u32(2).zeros(28).u32(instanceId).zeros(6).build();
}
function mkFieldEndCtrl(beginId) {
  return new BufWriter().u32(CTRL_FIELD_END).u32(0).zeros(28).u32(beginId).zeros(6).build();
}
function encodePicPara(imgNode, binDataId, bank, lv, idGen, availWidthHwp, divideSort = 0, vertPos = 0, maxHeightHwp = Number.MAX_SAFE_INTEGER) {
  const { w: wHwp, h: hHwp } = imageDisplaySizeHwp(
    imgNode,
    availWidthHwp,
    maxHeightHwp
  );
  const CTRL_MASK = 1 << 11;
  const instanceId = idGen();
  const psId = bank.addParaShape({});
  return [
    mkRec(
      TAG_PARA_HEADER2,
      lv,
      mkParaHeader(9, CTRL_MASK, psId, 1, 1, instanceId, 0, divideSort)
    ),
    mkRec(TAG_PARA_TEXT2, lv + 1, mkPicParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE2, lv + 1, mkParaCharShape([[0, 0]])),
    mkRec(
      TAG_PARA_LINE_SEG,
      lv + 1,
      buildObjectLineSeg(availWidthHwp, hHwp, vertPos)
    ),
    mkRec(
      TAG_CTRL_HEADER2,
      lv + 1,
      mkObjectCtrl(CTRL_GSO2, wHwp, hHwp, idGen(), imgNode.layout)
    ),
    mkRec(
      TAG_SHAPE_COMPONENT,
      lv + 2,
      mkShapeComponent(wHwp, hHwp)
    ),
    mkRec(
      TAG_SHAPE_COMPONENT_PICTURE2,
      lv + 3,
      mkShapeComponentPicture(binDataId, wHwp, hHwp, idGen())
    )
  ];
}
function encodePara3(para, bank, lv, instanceId, availWidthHwp, mask = 0, vertPos = 0, divideSortOverride = 0, sectionPrefix, pageBodyHeightHwp) {
  let text = "";
  const csPairs = [];
  let pos = 0;
  const fontHwp = maxFontHwpInPara(para);
  const ctrlRecords = [];
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
          } else if (t.tag === "br") {
            text += "\n";
            pos += 1;
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
  if (sectionPrefix && text.length === 0) text = " ";
  if (!csPairs.length) csPairs.push([0, 0]);
  const psId = bank.addParaShape(para.props);
  const styleId = hwpStyleIdForPara(para.props) ?? 0;
  const divideSort = divideSortOverride || (paraHasPageBreak2(para.kids) ? 4 : 0);
  const sectionCharCount = sectionPrefix ? 16 : 0;
  const effectiveMask = sectionPrefix ? mask | 1 << 2 : mask;
  const effectiveDivideSort = sectionPrefix ? divideSort | 3 : divideSort;
  const effectiveCsPairs = sectionPrefix ? [
    [0, 0],
    ...csPairs.map(([p, id]) => [p + sectionCharCount, id])
  ] : csPairs;
  const paraTextData = sectionPrefix ? new BufWriter().bytes(mkSectionAndColumnParaTextPrefix()).bytes(mkParaText(text)).build() : mkParaText(text);
  const sectionRecords = sectionPrefix ? buildSectionControlRecords(sectionPrefix.dims, lv + 1) : [];
  const nchars = sectionCharCount + text.length + 1;
  const lineSegs = buildParaLineSegs(
    text,
    availWidthHwp,
    fontHwp,
    para.props,
    vertPos,
    sectionCharCount,
    pageBodyHeightHwp
  );
  return [
    mkRec(
      TAG_PARA_HEADER2,
      lv,
      mkParaHeader(
        nchars,
        effectiveMask,
        psId,
        effectiveCsPairs.length,
        lineSegs.count,
        instanceId,
        styleId,
        effectiveDivideSort
      )
    ),
    mkRec(TAG_PARA_TEXT2, lv + 1, paraTextData),
    mkRec(TAG_PARA_CHAR_SHAPE2, lv + 1, mkParaCharShape(effectiveCsPairs)),
    mkRec(
      TAG_PARA_LINE_SEG,
      lv + 1,
      lineSegs.data
    ),
    ...sectionRecords,
    ...ctrlRecords
  ];
}
function mkTableCtrl(wHwp, hHwp, instanceId, align = "left") {
  const alignFlags = { left: 0, center: 1, right: 2, justify: 3, distribute: 0, distribute_space: 0 }[align] ?? 0;
  return new BufWriter().u32(CTRL_TABLE2).u32(136978961).i32(0).i32(0).u32(wHwp).u32(hHwp).i32(7).u16(140).u16(140).u16(140).u16(140).u32(instanceId).i32(alignFlags).u16(0).build();
}
function mkTableRecord(rowCnt, colCnt, cellCountPerRow, bfId, repeatHeader, padL, padR, padT, padB) {
  const w = new BufWriter();
  w.u32(67108865 | (repeatHeader ? 1 << 2 : 0)).u16(rowCnt).u16(colCnt).u16(0);
  w.u16(padL).u16(padR).u16(padT).u16(padB);
  for (const count of cellCountPerRow) w.u16(Math.max(1, count & 65535));
  w.u16(bfId).u16(0);
  return w.build();
}
function mkCellListHeader(paraCount, row, col, rs, cs, wHwp, hHwp, bfId, padL = 141, padR = 141, padT = 141, padB = 141, va) {
  const verticalAlign = va === "mid" ? 1 : va === "bot" ? 2 : 0;
  return new BufWriter().u16(paraCount).u16(0).u32(verticalAlign << 5).u16(col).u16(row).u16(cs).u16(rs).u32(wHwp).u32(hHwp).u16(padL).u16(padR).u16(padT).u16(padB).u16(bfId).zeros(13).build();
}
var DEFAULT_ROW_HEIGHT_PT = 14;
function encodeGrid4(grid, bank, lv, idGen, availWidthHwp, images) {
  const records = [];
  const rowCnt = grid.kids.length;
  const colCnt = Math.max(
    1,
    ...grid.kids.map(
      (row) => row.kids.reduce(
        (sum, cell) => sum + Math.max(1, cell.cs ?? 1),
        0
      )
    )
  );
  const sourceWidthsHwp = (grid.props.colWidths ?? []).map(
    (width) => width > 0 ? Metric.ptToHwp(width) : 0
  );
  const colWidthsHwp = fitColumnWidths(
    sourceWidthsHwp,
    colCnt,
    availWidthHwp,
    Math.min(100, Math.floor(availWidthHwp / colCnt))
  );
  const defStroke = grid.props.defaultStroke ?? bank.DEF_STROKE;
  const defBfId = bank.addBorderFill(defStroke);
  const tablePadL = Metric.ptToHwp(grid.props.cellPadL ?? 5.1);
  const tablePadR = Metric.ptToHwp(grid.props.cellPadR ?? 5.1);
  const tablePadT = Metric.ptToHwp(grid.props.cellPadT ?? 1.41);
  const tablePadB = Metric.ptToHwp(grid.props.cellPadB ?? 1.41);
  const rowHwp = grid.kids.map((row) => {
    const base = row.heightPt != null && row.heightPt > 0 ? Metric.ptToHwp(row.heightPt) : Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT);
    let minRow = 0;
    let logicalCol = 0;
    for (const cell of row.kids ?? []) {
      const cs = Math.max(1, cell.cs ?? 1);
      let cellWidthHwp = 0;
      for (let sc = logicalCol; sc < logicalCol + cs; sc++) {
        cellWidthHwp += colWidthsHwp[sc] ?? 0;
      }
      if (cellWidthHwp <= 0) cellWidthHwp = Math.floor(availWidthHwp / colCnt) * cs;
      const cp = cell.props ?? {};
      const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : tablePadL;
      const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : tablePadR;
      const innerWidthHwp = Math.max(
        100,
        cellWidthHwp - padL - padR
      );
      const span = Math.max(1, cell.rs ?? 1);
      minRow = Math.max(
        minRow,
        Math.ceil(minCellHeightHwp(cell, innerWidthHwp) / span)
      );
      logicalCol += cs;
    }
    return Math.max(base, minRow);
  });
  const cellCountPerRow = grid.kids.map(
    (row) => Math.max(1, row.kids.length)
  );
  const tblWidthHwp = colWidthsHwp.reduce((sum, width) => sum + width, 0);
  const tblHwp = rowHwp.reduce((s, h) => s + h, 0);
  const tblInstanceId = idGen();
  const tblAlign = grid.props.align ?? "left";
  records.push(
    mkRec(
      TAG_CTRL_HEADER2,
      lv,
      mkTableCtrl(
        tblWidthHwp,
        tblHwp,
        tblInstanceId,
        tblAlign
      )
    )
  );
  records.push(
    mkRec(
      TAG_TABLE,
      lv + 1,
      mkTableRecord(
        rowCnt,
        colCnt,
        cellCountPerRow,
        defBfId,
        !!grid.props.headerRow,
        tablePadL,
        tablePadR,
        tablePadT,
        tablePadB
      )
    )
  );
  for (let r = 0; r < grid.kids.length; r++) {
    let logicalCol = 0;
    for (let c = 0; c < grid.kids[r].kids.length; c++) {
      const cell = grid.kids[r].kids[c];
      const cs = Math.max(1, cell.cs ?? 1);
      let wHwp = 0;
      for (let sc = logicalCol; sc < logicalCol + cs; sc++) {
        wHwp += colWidthsHwp[sc] ?? 0;
      }
      if (wHwp <= 0) wHwp = Math.floor(availWidthHwp / colCnt) * cs;
      const hHwp = rowHwp.slice(r, Math.min(rowHwp.length, r + Math.max(1, cell.rs ?? 1))).reduce((sum, height) => sum + height, 0);
      const cp = cell.props;
      const hasPerSide = cp.top || cp.bot || cp.left || cp.right;
      const bfId = hasPerSide ? bank.addBorderFillPerSide(
        cp.left ?? defStroke,
        cp.right ?? defStroke,
        cp.top ?? defStroke,
        cp.bot ?? defStroke,
        cp.bg
      ) : bank.addBorderFill(defStroke, cp.bg);
      const cellKids = cell.kids.length > 0 ? cell.kids : [{ tag: "para", props: {}, kids: [] }];
      const padL = cp.padL !== void 0 ? Metric.ptToHwp(cp.padL) : tablePadL;
      const padR = cp.padR !== void 0 ? Metric.ptToHwp(cp.padR) : tablePadR;
      const padT = cp.padT !== void 0 ? Metric.ptToHwp(cp.padT) : tablePadT;
      const padB = cp.padB !== void 0 ? Metric.ptToHwp(cp.padB) : tablePadB;
      const encodedCellParagraphs = cellKids.reduce((count, kid) => {
        if (kid.tag === "grid") return count + 1;
        const paraImages = flatImgNodes(kid.kids).filter(
          (img) => images.some((bin) => b64Matches(bin, img.b64))
        );
        const textKids = kid.kids.filter((child) => child.tag !== "img");
        return count + paraImages.length + (textKids.length > 0 || paraImages.length === 0 ? 1 : 0);
      }, 0);
      records.push(
        mkRec(
          TAG_LIST_HEADER2,
          lv + 1,
          mkCellListHeader(
            Math.max(1, encodedCellParagraphs),
            r,
            logicalCol,
            cell.rs,
            cs,
            wHwp,
            hHwp,
            bfId,
            padL,
            padR,
            padT,
            padB,
            cp.va
          )
        )
      );
      const cellWidthHwp = Math.max(100, wHwp - padL - padR);
      let cellVertPos = 0;
      for (const kid of cellKids) {
        if (kid.tag === "grid") {
          const nestedGridHeight = minGridHeightHwp(kid);
          records.push(
            mkRec(
              TAG_PARA_HEADER2,
              lv + 1,
              mkParaHeader(9, TABLE_CTRL_MASK, 0, 1, 1, idGen())
            )
          );
          records.push(mkRec(TAG_PARA_TEXT2, lv + 2, mkTableParaText()));
          records.push(
            mkRec(TAG_PARA_CHAR_SHAPE2, lv + 2, mkParaCharShape([[0, 0]]))
          );
          records.push(
            mkRec(
              TAG_PARA_LINE_SEG,
              lv + 2,
              buildObjectLineSeg(cellWidthHwp, nestedGridHeight, cellVertPos)
            )
          );
          records.push(...encodeGrid4(kid, bank, lv + 2, idGen, cellWidthHwp, images));
          cellVertPos += nestedGridHeight + Metric.ptToHwp(6);
        } else {
          const para = kid;
          const paraImages = flatImgNodes(para.kids);
          for (const img of paraImages) {
            const binImg = images.find((bin) => b64Matches(bin, img.b64));
            if (!binImg) continue;
            records.push(
              ...encodePicPara(
                img,
                binImg.id,
                bank,
                lv + 1,
                idGen,
                cellWidthHwp,
                0,
                cellVertPos
              )
            );
            cellVertPos += imageDisplaySizeHwp(img, cellWidthHwp).h;
          }
          const textKids = para.kids.filter((child) => child.tag !== "img");
          if (textKids.length > 0 || paraImages.length === 0) {
            const textPara = { ...para, kids: textKids };
            records.push(
              ...encodePara3(
                textPara,
                bank,
                lv + 1,
                idGen(),
                cellWidthHwp,
                0,
                cellVertPos
              )
            );
            cellVertPos += minParaHeightHwp(textPara, cellWidthHwp);
          }
        }
      }
      logicalCol += cs;
    }
  }
  return records;
}
function mkSectionCtrl() {
  return new BufWriter().u32(CTRL_SECD2).u32(0).u32(1134).u32(524288e3).zeros(31).build();
}
function mkColumnDefCtrl() {
  return new BufWriter().u32(CTRL_COLD).u32(4100).u32(0).u32(0).build();
}
function mkPageBorderFill() {
  return new BufWriter().u32(1).u16(1417).u16(1417).u16(1417).u16(1417).u16(1).build();
}
function buildSectionControlRecords(dims, level) {
  return [
    mkRec(TAG_CTRL_HEADER2, level, mkSectionCtrl()),
    mkRec(TAG_PAGE_DEF2, level + 1, mkPageDef(dims)),
    mkRec(TAG_FOOTNOTE_SHAPE, level + 1, new Uint8Array(28)),
    mkRec(TAG_FOOTNOTE_SHAPE, level + 1, new Uint8Array(28)),
    mkRec(TAG_PAGE_BORDER_FILL, level + 1, mkPageBorderFill()),
    mkRec(TAG_PAGE_BORDER_FILL, level + 1, mkPageBorderFill()),
    mkRec(TAG_PAGE_BORDER_FILL, level + 1, mkPageBorderFill()),
    mkRec(TAG_CTRL_HEADER2, level, mkColumnDefCtrl())
  ];
}
function buildSectionParagraph(dims, instanceId) {
  const SECD_CTRL_MASK = 1 << 2;
  const nchars = 18;
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
    mkRec(
      TAG_PARA_TEXT2,
      1,
      new BufWriter().bytes(mkSectionAndColumnParaTextPrefix()).bytes(mkParaText(" ")).build()
    ),
    mkRec(TAG_PARA_CHAR_SHAPE2, 1, mkParaCharShape([[0, 0], [16, 0]])),
    mkRec(
      TAG_PARA_LINE_SEG,
      1,
      buildDefaultLineSeg(availWidthHwp, 1e3, nchars)
    ),
    ...buildSectionControlRecords(dims, 1)
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
function paraHasPageBreak2(kids) {
  return kids.some((kid) => {
    if (kid?.tag === "span")
      return (kid.kids ?? []).some((child) => child?.tag === "pb");
    if (kid?.tag === "link" && Array.isArray(kid.kids))
      return paraHasPageBreak2(kid.kids);
    return false;
  });
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
  const bodyHeightHwp = Math.max(
    1e3,
    Metric.ptToHwp(dims.hPt) - Metric.ptToHwp(dims.mt) - Metric.ptToHwp(dims.mb)
  );
  let vertPos = 0;
  let sectionWritten = false;
  for (const sheet of doc.kids) {
    for (const node of sheet.kids) {
      if (node.tag === "para") {
        const para = node;
        const hasPageBreak = paraHasPageBreak2(para.kids);
        const paraDivideSort = hasPageBreak ? 4 : 0;
        let paraMask = 0;
        const paraHeight = minParaHeightHwp(para, availWidthHwp);
        if (hasPageBreak || vertPos > 0 && vertPos + paraHeight > bodyHeightHwp) {
          vertPos = 0;
        }
        const hasCourier = (kids) => kids.some(
          (k) => k.tag === "span" && k.props.font?.toLowerCase().includes("courier") || k.tag === "link" && hasCourier(k.kids)
        );
        const isCode = para.props.styleId?.toLowerCase().includes("code") || hasCourier(para.kids);
        if (isCode) {
          if (!sectionWritten) {
            for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);
            sectionWritten = true;
          }
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
          const gridHeight = minGridHeightHwp(gridNode);
          if (vertPos > 0 && vertPos + gridHeight > bodyHeightHwp) vertPos = 0;
          chunks.push(
            mkRec(
              TAG_PARA_HEADER2,
              0,
              mkParaHeader(9, TABLE_CTRL_MASK | paraMask, 0, 1, 1, idGen(), 0, paraDivideSort)
            )
          );
          chunks.push(mkRec(TAG_PARA_TEXT2, 1, mkTableParaText()));
          chunks.push(mkRec(TAG_PARA_CHAR_SHAPE2, 1, mkParaCharShape([[0, 0]])));
          chunks.push(
            mkRec(
              TAG_PARA_LINE_SEG,
              1,
              buildObjectLineSeg(availWidthHwp, gridHeight, vertPos)
            )
          );
          for (const r of encodeGrid4(gridNode, bank, 1, idGen, availWidthHwp, images))
            chunks.push(r);
          vertPos += Math.max(Metric.ptToHwp(20), gridHeight + Metric.ptToHwp(6));
          continue;
        }
        const imgNodes = flatImgNodes(para.kids);
        if (imgNodes.length > 0) {
          if (!sectionWritten) {
            for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);
            sectionWritten = true;
          }
          let appliedImageBreak = false;
          for (const img of imgNodes) {
            const binImg = images.find((b) => b64Matches(b, img.b64));
            if (binImg) {
              const imageHeight = imageDisplaySizeHwp(
                img,
                availWidthHwp,
                bodyHeightHwp
              ).h;
              if (vertPos > 0 && vertPos + imageHeight > bodyHeightHwp) {
                vertPos = 0;
              }
              const imageDivideSort = !appliedImageBreak ? paraDivideSort : 0;
              appliedImageBreak = true;
              for (const r of encodePicPara(
                img,
                binImg.id,
                bank,
                0,
                idGen,
                availWidthHwp,
                imageDivideSort,
                vertPos,
                bodyHeightHwp
              )) {
                chunks.push(r);
              }
              vertPos += imageHeight + Metric.ptToHwp(6);
            }
          }
          const textKids = para.kids.filter((k) => k.tag !== "img");
          if (textKids.length > 0) {
            const textPara = {
              tag: "para",
              props: para.props,
              kids: textKids
            };
            const textHeight = minParaHeightHwp(textPara, availWidthHwp);
            if (vertPos > 0 && vertPos + textHeight > bodyHeightHwp) vertPos = 0;
            for (const r of encodePara3(
              textPara,
              bank,
              0,
              idGen(),
              availWidthHwp,
              paraMask,
              vertPos,
              imgNodes.length > 0 ? 0 : paraDivideSort,
              void 0,
              bodyHeightHwp
            )) {
              if (r[0] === (TAG_PARA_HEADER2 & 255)) {
              }
              chunks.push(r);
            }
            vertPos = nextParaVertPosHwp(
              textPara,
              vertPos,
              availWidthHwp,
              bodyHeightHwp
            );
          }
        } else {
          for (const r of encodePara3(
            para,
            bank,
            0,
            idGen(),
            availWidthHwp,
            paraMask,
            vertPos,
            paraDivideSort,
            !sectionWritten ? { dims } : void 0,
            bodyHeightHwp
          ))
            chunks.push(r);
          sectionWritten = true;
          vertPos = nextParaVertPosHwp(
            para,
            vertPos,
            availWidthHwp,
            bodyHeightHwp
          );
        }
      } else if (node.tag === "grid") {
        if (!sectionWritten) {
          for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);
          sectionWritten = true;
        }
        const gridHeight = minGridHeightHwp(node);
        if (vertPos > 0 && vertPos + gridHeight > bodyHeightHwp) vertPos = 0;
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
            buildObjectLineSeg(availWidthHwp, gridHeight, vertPos)
          )
        );
        for (const r of encodeGrid4(
          node,
          bank,
          1,
          idGen,
          availWidthHwp,
          images
        ))
          chunks.push(r);
        vertPos += Math.max(
          Metric.ptToHwp(20),
          gridHeight + Metric.ptToHwp(6)
        );
      }
    }
  }
  if (!sectionWritten) {
    for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);
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
  dv.setUint32(32, 83951617, true);
  dv.setUint32(36, 1, true);
  if (buf.length !== SIZE) {
    throw new Error(`FileHeader \uD06C\uAE30 \uC624\uB958: ${buf.length} (\uAE30\uB300: ${SIZE})`);
  }
  if (new TextDecoder().decode(buf.subarray(0, sig.length)) !== sig) {
    throw new Error("FileHeader \uC2DC\uADF8\uB2C8\uCC98 \uC624\uB958");
  }
  if (dv.getUint32(32, true) !== 83951617) {
    throw new Error("FileHeader \uBC84\uC804 \uC624\uB958");
  }
  return buf;
}
function buildHwpOle2(fileHeaderData, docInfoData, section0Data, binImages = []) {
  const SS = 512;
  const MSS = 64;
  const ENDOFCHAIN = 4294967294;
  const FREESECT = 4294967295;
  const FATSECT = 4294967293;
  const DIFSECT = 4294967292;
  if (fileHeaderData.length < 256) {
    throw new Error(
      `FileHeader \uD06C\uAE30 \uBD80\uC871: ${fileHeaderData.length} (\uCD5C\uC18C 256)`
    );
  }
  const streams = [];
  streams.push({
    name: "FileHeader",
    data: fileHeaderData,
    dirIdx: 1,
    isMini: fileHeaderData.length < 4096
  });
  streams.push({
    name: "DocInfo",
    data: docInfoData,
    dirIdx: 2,
    isMini: docInfoData.length < 4096
  });
  streams.push({
    name: "Section0",
    data: section0Data,
    dirIdx: 4,
    isMini: section0Data.length < 4096
  });
  const prvTextDirIdx = binImages.length > 0 ? 6 + binImages.length : 5;
  const prvTextData = new BufWriter().utf16("HWPKit Preview\r\n").build();
  streams.push({
    name: "PrvText",
    data: prvTextData,
    dirIdx: prvTextDirIdx,
    isMini: prvTextData.length < 4096
  });
  for (let i = 0; i < binImages.length; i++) {
    const img = binImages[i];
    const name = `BIN${img.id.toString(16).toUpperCase().padStart(4, "0")}.${img.ext}`;
    streams.push({
      name,
      data: img.data,
      dirIdx: 6 + i,
      isMini: img.data.length < 4096
    });
  }
  const miniStreams = streams.filter((s) => s.isMini);
  const miniSectorList = [];
  let miniStreamDataLength = 0;
  for (const s of miniStreams) {
    const startSec = miniStreamDataLength / MSS;
    s.startSec = startSec;
    const len = s.data.length;
    const numMiniSecs = Math.ceil(len / MSS);
    for (let i = 0; i < numMiniSecs; i++) {
      const curSec2 = startSec + i;
      const nextSec = i === numMiniSecs - 1 ? ENDOFCHAIN : curSec2 + 1;
      while (miniSectorList.length <= curSec2) {
        miniSectorList.push(FREESECT);
      }
      miniSectorList[curSec2] = nextSec;
    }
    miniStreamDataLength += numMiniSecs * MSS;
  }
  const miniStreamData = new Uint8Array(miniStreamDataLength);
  let miniStreamOffset = 0;
  for (const s of miniStreams) {
    miniStreamData.set(s.data, miniStreamOffset);
    miniStreamOffset += Math.ceil(s.data.length / MSS) * MSS;
  }
  const regularStreams = streams.filter((s) => !s.isMini);
  const regPads = regularStreams.map((s) => {
    const len = s.data.length;
    const n = Math.ceil(Math.max(len, 1) / SS) * SS;
    const out2 = new Uint8Array(n);
    out2.set(s.data);
    return out2;
  });
  const regNs = regPads.map((p) => p.length / SS);
  const numDirEntries = 6 + (binImages.length > 0 ? 1 + binImages.length : 0);
  const dirN = Math.max(1, Math.ceil(numDirEntries * 128 / SS));
  const miniFatN = Math.ceil(miniSectorList.length / 128);
  const miniStreamN = Math.ceil(miniStreamData.length / SS);
  const totalRegStreamN = regNs.reduce((a, b) => a + b, 0);
  const neededDataSec = dirN + miniFatN + miniStreamN + totalRegStreamN;
  let fatN = 1;
  let difatN = 0;
  for (let iter = 0; iter < 100; iter++) {
    const nextFatN = Math.ceil((neededDataSec + fatN + difatN) / 128);
    const nextDifatN = Math.ceil(Math.max(0, nextFatN - 109) / 127);
    if (nextFatN === fatN && nextDifatN === difatN) break;
    fatN = nextFatN;
    difatN = nextDifatN;
  }
  const totalSec = fatN + difatN + neededDataSec;
  const difatStartSec = fatN;
  const dirStartSec = difatStartSec + difatN;
  const miniFatStartSec = dirStartSec + dirN;
  const miniStreamStartSec = miniFatStartSec + miniFatN;
  let curSec = miniStreamStartSec + miniStreamN;
  for (let i = 0; i < regularStreams.length; i++) {
    regularStreams[i].startSec = curSec;
    curSec += regNs[i];
  }
  const fatBuf = new Uint8Array(fatN * SS).fill(255);
  const setFat = (i, v) => {
    const off = i * 4;
    fatBuf[off] = v & 255;
    fatBuf[off + 1] = v >>> 8 & 255;
    fatBuf[off + 2] = v >>> 16 & 255;
    fatBuf[off + 3] = v >>> 24 & 255;
  };
  for (let i = 0; i < fatN; i++) {
    setFat(i, FATSECT);
  }
  for (let i = 0; i < difatN; i++) {
    setFat(difatStartSec + i, DIFSECT);
  }
  for (let i = 0; i < dirN; i++) {
    setFat(
      dirStartSec + i,
      i + 1 < dirN ? dirStartSec + i + 1 : ENDOFCHAIN
    );
  }
  if (miniFatN > 0) {
    for (let i = 0; i < miniFatN; i++) {
      setFat(
        miniFatStartSec + i,
        i + 1 < miniFatN ? miniFatStartSec + i + 1 : ENDOFCHAIN
      );
    }
  }
  if (miniStreamN > 0) {
    for (let i = 0; i < miniStreamN; i++) {
      setFat(
        miniStreamStartSec + i,
        i + 1 < miniStreamN ? miniStreamStartSec + i + 1 : ENDOFCHAIN
      );
    }
  }
  for (let i = 0; i < regularStreams.length; i++) {
    const s = regularStreams[i];
    const n = regNs[i];
    const start = s.startSec;
    for (let j = 0; j < n; j++) {
      setFat(start + j, j + 1 < n ? start + j + 1 : ENDOFCHAIN);
    }
  }
  const miniFatBuf = new Uint8Array(miniFatN * SS).fill(255);
  const setMiniFat = (i, v) => {
    const off = i * 4;
    miniFatBuf[off] = v & 255;
    miniFatBuf[off + 1] = v >>> 8 & 255;
    miniFatBuf[off + 2] = v >>> 16 & 255;
    miniFatBuf[off + 3] = v >>> 24 & 255;
  };
  for (let i = 0; i < miniSectorList.length; i++) {
    setMiniFat(i, miniSectorList[i]);
  }
  const difatBuf = new Uint8Array(difatN * SS).fill(255);
  const difatView = new DataView(difatBuf.buffer);
  for (let i = 0; i < difatN; i++) {
    const base = i * SS;
    for (let j = 0; j < 127; j++) {
      const fatSectorId = 109 + i * 127 + j;
      difatView.setUint32(
        base + j * 4,
        fatSectorId < fatN ? fatSectorId : FREESECT,
        true
      );
    }
    difatView.setUint32(
      base + 127 * 4,
      i + 1 < difatN ? difatStartSec + i + 1 : ENDOFCHAIN,
      true
    );
  }
  const dirBuf = new Uint8Array(dirN * SS);
  const dv = new DataView(dirBuf.buffer);
  function writeDirEntry(idx, name, type, color, left, right, child, startSec, size) {
    const base = idx * 128;
    const nl = name.length;
    for (let i = 0; i < nl; i++) {
      dv.setUint16(base + i * 2, name.charCodeAt(i), true);
    }
    dv.setUint16(base + 64, (nl + 1) * 2, true);
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = color;
    dv.setInt32(base + 68, left, true);
    dv.setInt32(base + 72, right, true);
    dv.setInt32(base + 76, child, true);
    dv.setUint32(base + 116, startSec >>> 0, true);
    dv.setUint32(base + 120, size >>> 0, true);
  }
  function cfbNameCompare(a, b) {
    if (a.length !== b.length) return a.length - b.length;
    const au = a.toUpperCase();
    const bu = b.toUpperCase();
    return au < bu ? -1 : au > bu ? 1 : 0;
  }
  function buildSiblingTree(nodes) {
    const links = /* @__PURE__ */ new Map();
    const colors = /* @__PURE__ */ new Map();
    const parents = /* @__PURE__ */ new Map();
    const names = new Map(nodes.map((node) => [node.idx, node.name]));
    let root = -1;
    const colorOf = (idx) => idx < 0 ? 1 : colors.get(idx) ?? 1;
    const parentOf = (idx) => parents.get(idx) ?? -1;
    const leftOf = (idx) => links.get(idx)?.left ?? -1;
    const rightOf = (idx) => links.get(idx)?.right ?? -1;
    const rotateLeft = (x) => {
      const y = rightOf(x);
      if (y < 0) return;
      links.get(x).right = leftOf(y);
      if (leftOf(y) >= 0) parents.set(leftOf(y), x);
      const xp = parentOf(x);
      parents.set(y, xp);
      if (xp < 0) root = y;
      else if (x === leftOf(xp)) links.get(xp).left = y;
      else links.get(xp).right = y;
      links.get(y).left = x;
      parents.set(x, y);
    };
    const rotateRight = (x) => {
      const y = leftOf(x);
      if (y < 0) return;
      links.get(x).left = rightOf(y);
      if (rightOf(y) >= 0) parents.set(rightOf(y), x);
      const xp = parentOf(x);
      parents.set(y, xp);
      if (xp < 0) root = y;
      else if (x === rightOf(xp)) links.get(xp).right = y;
      else links.get(xp).left = y;
      links.get(y).right = x;
      parents.set(x, y);
    };
    for (const node of nodes) {
      links.set(node.idx, { left: -1, right: -1 });
      colors.set(node.idx, 0);
      let parent = -1;
      let cursor = root;
      while (cursor >= 0) {
        parent = cursor;
        cursor = cfbNameCompare(node.name, names.get(cursor)) < 0 ? leftOf(cursor) : rightOf(cursor);
      }
      parents.set(node.idx, parent);
      if (parent < 0) root = node.idx;
      else if (cfbNameCompare(node.name, names.get(parent)) < 0)
        links.get(parent).left = node.idx;
      else links.get(parent).right = node.idx;
      let z = node.idx;
      while (z !== root && colorOf(parentOf(z)) === 0) {
        const p = parentOf(z);
        const gp = parentOf(p);
        if (p === leftOf(gp)) {
          const uncle = rightOf(gp);
          if (colorOf(uncle) === 0) {
            colors.set(p, 1);
            colors.set(uncle, 1);
            colors.set(gp, 0);
            z = gp;
          } else {
            if (z === rightOf(p)) {
              z = p;
              rotateLeft(z);
            }
            colors.set(parentOf(z), 1);
            colors.set(parentOf(parentOf(z)), 0);
            rotateRight(parentOf(parentOf(z)));
          }
        } else {
          const uncle = leftOf(gp);
          if (colorOf(uncle) === 0) {
            colors.set(p, 1);
            colors.set(uncle, 1);
            colors.set(gp, 0);
            z = gp;
          } else {
            if (z === leftOf(p)) {
              z = p;
              rotateRight(z);
            }
            colors.set(parentOf(z), 1);
            colors.set(parentOf(parentOf(z)), 0);
            rotateLeft(parentOf(parentOf(z)));
          }
        }
      }
      if (root >= 0) colors.set(root, 1);
    }
    return { root, links, colors };
  }
  for (let i = 0; i < dirN * 4; i++) {
    const base = i * 128;
    dv.setInt32(base + 68, -1, true);
    dv.setInt32(base + 72, -1, true);
    dv.setInt32(base + 76, -1, true);
  }
  const streamMap = /* @__PURE__ */ new Map();
  for (const s of streams) {
    streamMap.set(s.dirIdx, s);
  }
  writeDirEntry(
    0,
    "Root Entry",
    5,
    0,
    -1,
    -1,
    -1,
    miniStreamStartSec,
    miniStreamData.length
  );
  const hasBinData = binImages.length > 0;
  const rootTree = buildSiblingTree([
    ...hasBinData ? [{ idx: 5, name: "BinData" }] : [],
    { idx: 3, name: "BodyText" },
    { idx: 2, name: "DocInfo" },
    { idx: 1, name: "FileHeader" },
    { idx: prvTextDirIdx, name: "PrvText" }
  ]);
  dv.setInt32(76, rootTree.root, true);
  const rootLinks = (idx) => rootTree.links.get(idx) ?? { left: -1, right: -1 };
  const fhStream = streamMap.get(1);
  const fhLinks = rootLinks(1);
  writeDirEntry(
    1,
    "FileHeader",
    2,
    rootTree.colors.get(1) ?? 1,
    fhLinks.left,
    fhLinks.right,
    -1,
    fhStream.startSec,
    fhStream.data.length
  );
  const diStream = streamMap.get(2);
  const diLinks = rootLinks(2);
  writeDirEntry(
    2,
    "DocInfo",
    2,
    rootTree.colors.get(2) ?? 1,
    diLinks.left,
    diLinks.right,
    -1,
    diStream.startSec,
    diStream.data.length
  );
  const bodyLinks = rootLinks(3);
  writeDirEntry(
    3,
    "BodyText",
    1,
    rootTree.colors.get(3) ?? 1,
    bodyLinks.left,
    bodyLinks.right,
    4,
    ENDOFCHAIN,
    0
  );
  const s0Stream = streamMap.get(4);
  writeDirEntry(
    4,
    "Section0",
    2,
    1,
    -1,
    -1,
    -1,
    s0Stream.startSec,
    s0Stream.data.length
  );
  const prvTextStream = streamMap.get(prvTextDirIdx);
  const prvTextLinks = rootLinks(prvTextDirIdx);
  writeDirEntry(
    prvTextDirIdx,
    "PrvText",
    2,
    rootTree.colors.get(prvTextDirIdx) ?? 1,
    prvTextLinks.left,
    prvTextLinks.right,
    -1,
    prvTextStream.startSec,
    prvTextStream.data.length
  );
  if (hasBinData) {
    const binLinks = rootLinks(5);
    const binTree = buildSiblingTree(
      binImages.map((img, i) => ({
        idx: 6 + i,
        name: `BIN${img.id.toString(16).toUpperCase().padStart(4, "0")}.${img.ext}`
      }))
    );
    writeDirEntry(
      5,
      "BinData",
      1,
      rootTree.colors.get(5) ?? 1,
      binLinks.left,
      binLinks.right,
      binTree.root,
      ENDOFCHAIN,
      0
    );
    for (let i = 0; i < binImages.length; i++) {
      const imgStream = streamMap.get(6 + i);
      const imgLinks = binTree.links.get(6 + i) ?? { left: -1, right: -1 };
      writeDirEntry(
        6 + i,
        imgStream.name,
        2,
        binTree.colors.get(6 + i) ?? 1,
        imgLinks.left,
        imgLinks.right,
        -1,
        imgStream.startSec,
        imgStream.data.length
      );
    }
  }
  const hdr = new Uint8Array(SS);
  const hdv = new DataView(hdr.buffer);
  const MAGIC = [208, 207, 17, 224, 161, 177, 26, 225];
  MAGIC.forEach((b, i) => {
    hdr[i] = b;
  });
  hdv.setUint16(24, 62, true);
  hdv.setUint16(26, 3, true);
  hdv.setUint16(28, 65534, true);
  hdv.setUint16(30, 9, true);
  hdv.setUint16(32, 6, true);
  hdv.setUint32(40, 0, true);
  hdv.setUint32(44, fatN, true);
  hdv.setUint32(48, dirStartSec, true);
  hdv.setUint32(52, 0, true);
  hdv.setUint32(56, 4096, true);
  hdv.setUint32(60, miniFatN > 0 ? miniFatStartSec : ENDOFCHAIN, true);
  hdv.setUint32(64, miniFatN, true);
  hdv.setUint32(68, difatN > 0 ? difatStartSec : ENDOFCHAIN, true);
  hdv.setUint32(72, difatN, true);
  for (let i = 0; i < 109; i++) {
    hdv.setUint32(76 + i * 4, i < fatN ? i : FREESECT, true);
  }
  const out = new Uint8Array(SS + totalSec * SS);
  let outOff = 0;
  out.set(hdr, outOff);
  outOff += SS;
  out.set(fatBuf, outOff);
  outOff += fatN * SS;
  if (difatN > 0) {
    out.set(difatBuf, outOff);
    outOff += difatN * SS;
  }
  out.set(dirBuf, outOff);
  outOff += dirN * SS;
  if (miniFatN > 0) {
    out.set(miniFatBuf, outOff);
    outOff += miniFatN * SS;
  }
  if (miniStreamN > 0) {
    const miniStreamPad = new Uint8Array(miniStreamN * SS);
    miniStreamPad.set(miniStreamData);
    out.set(miniStreamPad, outOff);
    outOff += miniStreamN * SS;
  }
  for (let i = 0; i < regularStreams.length; i++) {
    out.set(regPads[i], outOff);
    outOff += regNs[i] * SS;
  }
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
        const key = String(img.b64 ?? "").replace(/\s/g, "");
        if (seenB64.has(key)) return;
        seenB64.add(key);
        const raw = TextKit.base64Decode(img.b64);
        const ext = img.mime === "image/png" ? "png" : img.mime === "image/gif" ? "gif" : img.mime === "image/bmp" ? "bmp" : img.mime === "image/x-wmf" ? "wmf" : img.mime === "image/x-emf" ? "emf" : "jpg";
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