import { IFileSystem } from "../abstraction/FileSystem";
import { xmlEscape } from "../hwpx/utils";
// import { imageSize } from 'image-size'; // Removed for browser compatibility
import { getHeaderTemplate, PageMargin } from "../hml_template";

// Constants

// Python 버전과 동일한 헤더 상수
const H1_START = '<P ParaShape="20" Style="0"><TEXT CharShape="10"><CHAR>';
const H2_START = '<P ParaShape="20" Style="0"><TEXT CharShape="11"><CHAR>';
const H3_START = '<P ParaShape="20" Style="0"><TEXT CharShape="12"><CHAR>';
const H4_START = '<P ParaShape="20" Style="0"><TEXT CharShape="13"><CHAR>';
const H5_START = '<P ParaShape="20" Style="0"><TEXT CharShape="14"><CHAR>';
const H6_START = '<P ParaShape="20" Style="0"><TEXT CharShape="7"><CHAR>';

const INTRO = `
`;

export class HmlConverter {
  private fs: IFileSystem;
  private BIN_ITEM_ENTRIES: string[] = [];
  private BIN_STORAGE_ENTRIES: string[] = [];
  private PLACEHOLDERS: { [key: string]: string } = {};
  private ph_counter = 0;
  private borderFills: { id: number; xml: string }[] = [];

  constructor(fileSystem: IFileSystem) {
    this.fs = fileSystem;
  }

  private register_placeholder(xml_content: string): string {
    const key = `__MD2HML_PH_${this.ph_counter}__`;
    this.PLACEHOLDERS[key] = xml_content;
    this.ph_counter++;
    return key;
  }

  private async fetchImage(
    url: string,
  ): Promise<{ data: Uint8Array; ext: string } | null> {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const arrayBuffer = await res.arrayBuffer();
      const buffer = new Uint8Array(arrayBuffer); // Use Uint8Array for generic

      let ext = "jpg";
      const contentType = res.headers.get("content-type");
      if (contentType) {
        if (contentType.includes("png")) ext = "png";
        else if (contentType.includes("gif")) ext = "gif";
        else if (contentType.includes("bmp")) ext = "bmp";
      } else {
        // Try URL
        const parts = url.split(".");
        const extPart =
          parts.length > 1 ? parts[parts.length - 1].toLowerCase() : ""; // Weak check
        if (extPart) ext = extPart;
      }
      return { data: buffer, ext };
    } catch (e) {
      throw e;
    }
  }

  private async getImageSize(
    buffer: Uint8Array,
  ): Promise<{ width: number; height: number }> {
    return new Promise((resolve) => {
      try {
        if (typeof window === "undefined") {
          // Node environment fallback or just return 0
          resolve({ width: 0, height: 0 });
          return;
        }
        const blob = new Blob([buffer as any]);
        const url = URL.createObjectURL(blob);
        const img = new Image();
        img.onload = () => {
          const width = img.naturalWidth;
          const height = img.naturalHeight;
          URL.revokeObjectURL(url);
          resolve({ width, height });
        };
        img.onerror = () => {
          URL.revokeObjectURL(url);
          resolve({ width: 0, height: 0 });
        };
        img.src = url;
      } catch (e) {
        console.warn("Image size calculation failed:", e);
        resolve({ width: 0, height: 0 });
      }
    });
  }

  private async process_image(alt: string, imgPath: string): Promise<string> {
    let buffer: Uint8Array = new Uint8Array(0);
    let ext: string = "jpg";

    try {
      if (imgPath.startsWith("data:")) {
        const commaIdx = imgPath.indexOf(",");
        if (commaIdx > -1) {
          const base64 = imgPath.substring(commaIdx + 1);
          const mime = imgPath.substring(5, commaIdx).split(";")[0];

          if (typeof Buffer !== "undefined") {
            buffer = new Uint8Array(Buffer.from(base64, "base64"));
          } else {
            const binaryStr = atob(base64);
            const len = binaryStr.length;
            buffer = new Uint8Array(len);
            for (let i = 0; i < len; i++) buffer[i] = binaryStr.charCodeAt(i);
          }

          if (mime.includes("png")) ext = "png";
          else if (mime.includes("gif")) ext = "gif";
          else if (mime.includes("bmp")) ext = "bmp";
          else ext = "jpg";
        }
      } else if (
        imgPath.startsWith("http://") ||
        imgPath.startsWith("https://")
      ) {
        const result = await this.fetchImage(imgPath);
        if (!result) throw new Error("Fetch failed");
        buffer = result.data;
        ext = result.ext;
      } else {
        let found = false;
        if (await this.fs.exists(imgPath)) {
          buffer = await this.fs.readFile(imgPath);
          found = true;
        } else {
          // Try basename search (handling flat file system in browser)
          const basename = imgPath.split("/").pop() || imgPath;
          if (await this.fs.exists(basename)) {
            buffer = await this.fs.readFile(basename);
            found = true;
          }
        }

        if (!found) {
          return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>[Image not found: ${imgPath}]</CHAR></TEXT></P>`;
        }

        const parts = imgPath.split(".");
        ext = parts.length > 1 ? parts[parts.length - 1].toLowerCase() : "jpg";
      }

      if (ext === "jpeg") ext = "jpg";
      const validExts = ["jpg", "png", "gif", "bmp"];
      if (!validExts.includes(ext)) ext = "jpg";

      // Convert to Base64
      // Browser compatible base64
      let encodedString = "";
      if (typeof Buffer !== "undefined") {
        encodedString = Buffer.from(buffer).toString("base64");
      } else {
        // Browser fallback
        let binary = "";
        const len = buffer.byteLength;
        for (let i = 0; i < len; i++) {
          binary += String.fromCharCode(buffer[i]);
        }
        encodedString = window.btoa(binary);
      }

      const encodedLength = encodedString.length;

      const bin_id = this.BIN_ITEM_ENTRIES.length + 1;

      this.BIN_ITEM_ENTRIES.push(
        `<BINITEM BinData="${bin_id}" Format="${ext}" Type="Embedding" />`,
      );
      this.BIN_STORAGE_ENTRIES.push(
        `<BINDATA Encoding="Base64" Id="${bin_id}" Size="${encodedLength}">${encodedString}</BINDATA>`,
      );

      let width = 28346;
      let height = 28346;

      try {
        // Use browser-compatible image size calculation
        const dims = await this.getImageSize(buffer);

        if (dims.width && dims.height) {
          width = Math.floor(dims.width * 75);
          height = Math.floor(dims.height * 75);

          const MAX_WIDTH = 42000;
          if (width > MAX_WIDTH) {
            const ratio = MAX_WIDTH / width;
            width = MAX_WIDTH;
            height = Math.floor(height * ratio);
          }
        }
      } catch (err) {
        console.warn("Image size error (using default):", err);
      }

      const instId = 10000000 + bin_id;

      const xml_code = `<P ParaShape="0" Style="0"><TEXT CharShape="0"><PICTURE Reverse="false">
<SHAPEOBJECT Lock="false" NumberingType="Figure" TextFlow="BothSides" TextWrap="Square" ZOrder="6">
<SIZE Height="${height}" HeightRelTo="Absolute" Protect="false" Width="${width}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="false" AllowOverlap="true" FlowWithText="true" HoldAnchorAndSO="false" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="true" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="0" Left="0" Right="0" Top="0"/>
<SHAPECOMMENT>${xmlEscape(alt)}</SHAPECOMMENT>
</SHAPEOBJECT>
<SHAPECOMPONENT GroupLevel="0" HorzFlip="false" InstID="${instId}" OriHeight="${height}" OriWidth="${width}" VertFlip="false" XPos="0" YPos="0">
<ROTATIONINFO Angle="0" CenterX="${Math.floor(width / 2)}" CenterY="${Math.floor(height / 2)}" Rotate="1"/>
<RENDERINGINFO>
<TRANSMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
<SCAMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
<ROTMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
</RENDERINGINFO>
</SHAPECOMPONENT>
<IMAGERECT X0="0" X1="${width}" X2="${width}" X3="0" Y0="0" Y1="0" Y2="${height}" Y3="${height}"/>
<IMAGECLIP Bottom="${height}" Left="0" Right="${width}" Top="0"/>
<INSIDEMARGIN Bottom="0" Left="0" Right="0" Top="0"/>
<IMAGE Alpha="0" BinItem="${bin_id}" Bright="0" Contrast="0" Effect="RealPic"/>
<EFFECTS/>
</PICTURE></TEXT></P>`;
      return xml_code;
    } catch (e: any) {
      return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>[Image error: ${e.message}]</CHAR></TEXT></P>`;
    }
  }

  private process_code_block(text: string): string {
    const lines = text.trim().split("\n");
    let cell_content = "";

    lines.forEach((line) => {
      cell_content += `<P ParaShape="0" Style="0"><TEXT CharShape="16"><CHAR>${xmlEscape(line)}</CHAR></TEXT></P>`;
    });

    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="1" PageBreak="Cell" RepeatHeader="1" RowCount="1">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="1">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="41954" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<ROW>
<CELL BorderFill="3" ColAddr="0" ColSpan="1" Dirty="0" Editable="0" HasMargin="0" Header="0" Height="282" Protect="0" RowAddr="0" RowSpan="1" Width="41954">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${cell_content}
</PARALIST>
</CELL>
</ROW>
</TABLE><CHAR/></TEXT></P>`;
  }

  private process_table_native(table_block: string): string {
    const lines = table_block.trim().split("\n");
    if (lines.length < 2) return table_block;

    const header_cells = lines[0]
      .trim()
      .replace(/^\||\|$/g, "")
      .split("|")
      .map((c) => c.trim());
    const col_count = header_cells.length;

    const rows_data: string[][] = [];
    rows_data.push(header_cells);

    for (let i = 2; i < lines.length; i++) {
      let row_cells = lines[i]
        .trim()
        .replace(/^\||\|$/g, "")
        .split("|")
        .map((c) => c.trim());
      if (row_cells.length < col_count) {
        row_cells = row_cells.concat(
          new Array(col_count - row_cells.length).fill(""),
        );
      }
      rows_data.push(row_cells.slice(0, col_count));
    }

    const row_count = rows_data.length;
    const total_width = 41954;
    const col_width = Math.floor(total_width / col_count);

    let hml_rows = "";

    rows_data.forEach((row, r_idx) => {
      let cells_xml = "";
      row.forEach((cell_text, c_idx) => {
        const borderFillId = r_idx === 0 ? 4 : 3;
        const isHeader = r_idx === 0 ? 1 : 0;
        const paraShape = r_idx === 0 ? 26 : 0;
        const charShape = r_idx === 0 ? 7 : 0;

        const lines = cell_text.split(/<br\s*\/?>|\n/i);
        let p_xml = "";
        lines.forEach((line) => {
          p_xml += `<P ParaShape="${paraShape}" Style="0"><TEXT CharShape="${charShape}"><CHAR>${xmlEscape(line.trim() || " ")}</CHAR></TEXT></P>\n`;
        });

        cells_xml += `<CELL BorderFill="${borderFillId}" ColAddr="${c_idx}" ColSpan="1" Dirty="0" Editable="0" HasMargin="0" Header="${isHeader}" Height="282" Protect="0" RowAddr="${r_idx}" RowSpan="1" Width="${col_width}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${p_xml.trim()}
</PARALIST>
</CELL>`;
      });
      hml_rows += `<ROW>${cells_xml}</ROW>`;
    });

    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="${col_count}" PageBreak="Cell" RepeatHeader="1" RowCount="${row_count}">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="${total_width}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
${hml_rows}
</TABLE><CHAR/></TEXT></P>`;
  }

  private processInline(text: string): string {
    // Bold+Italic: ***text*** or ___text___ (must come before ** and *)
    // CharShape="19" is Bold+Italic
    text = text.replace(
      /\*\*\*(?=\S)([\s\S]+?)(?<=\S)\*\*\*/g,
      (_match, p1) => {
        return `</CHAR></TEXT><TEXT CharShape="19"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
      },
    );
    text = text.replace(/___(?=\S)([\s\S]+?)(?<=\S)___/g, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="19"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    // Bold: **text** -> CharShape="7" (Bold)
    text = text.replace(/\*\*(?=\S)([\s\S]+?)(?<=\S)\*\*/g, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="7"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    // Italic: *text* -> CharShape="15" (Italic)
    text = text.replace(
      /(?<!\*)\*(?=\S)([\s\S]+?)(?<=\S)\*(?!\*)/g,
      (_match, p1) => {
        return `</CHAR></TEXT><TEXT CharShape="15"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
      },
    );

    // Bold: __text__ -> CharShape="7"
    text = text.replace(/__(?=\S)([\s\S]+?)(?<=\S)__/g, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="7"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    // Italic: _text_ -> CharShape="15"
    text = text.replace(
      /(?<!_)_(?=\S)([\s\S]+?)(?<=\S)_(?!_)/g,
      (_match, p1) => {
        return `</CHAR></TEXT><TEXT CharShape="15"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
      },
    );

    // Strikethrough: ~~text~~ -> CharShape="17"
    text = text.replace(/~~(?=\S)([\s\S]+?)(?<=\S)~~/g, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="17"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    // Underline: <u>text</u> - xmlEscape 후 &lt;u&gt; 형태로 변환됨
    text = text.replace(/&lt;u&gt;([\s\S]*?)&lt;\/u&gt;/gi, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="18"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    // Inline Code: `text` -> CharShape="16" (Monospace+Bold)
    text = text.replace(/`([^`]+)`/g, (_match, p1) => {
      return `</CHAR></TEXT><TEXT CharShape="16"><CHAR>${p1}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`;
    });

    return text;
  }

  public async convert(markdown: string): Promise<string> {
    let ret = markdown;
    ret = ret.replace(/\r/g, "");
    ret = "\n" + ret;

    this.BIN_ITEM_ENTRIES = [];
    this.BIN_STORAGE_ENTRIES = [];
    this.PLACEHOLDERS = {};
    this.ph_counter = 0;
    this.borderFills = [];

    const imgRegex = /!\[(.*?)\]\((.*?)\)/g;
    let match;
    const imgMatches = [];
    while ((match = imgRegex.exec(ret)) !== null) {
      imgMatches.push({
        fullMatch: match[0],
        alt: match[1],
        path: match[2],
      });
    }

    const processedImages = await Promise.all(
      imgMatches.map(async (m) => {
        const xml = await this.process_image(m.alt, m.path);
        const placeholder = this.register_placeholder(xml);
        return { fullMatch: m.fullMatch, placeholder };
      }),
    );

    let imgIdx = 0;
    ret = ret.replace(imgRegex, () => {
      const res = processedImages[imgIdx];
      imgIdx++;
      return res ? res.placeholder : "";
    });

    ret = ret.replace(/^```([\s\S]*?)```/gm, (_match, p1) =>
      this.register_placeholder(this.process_code_block(p1)),
    );

    let prev = "";
    while (ret !== prev) {
      prev = ret;
      ret = ret.replace(/<table(?:(?!<table)[\s\S])*?<\/table>/gi, (_match) =>
        this.register_placeholder(this.process_html_table(_match)),
      );
    }

    // 3. XML Escape
    ret = xmlEscape(ret);

    // Metadata
    let t = "No Title";
    let a = "Unknown";
    let d = "Unknown";
    let t_match = ret.match(/(?<=title: ).*/);
    if (t_match) t = t_match[0];
    let a_match = ret.match(/(?<=author: ).*/);
    if (a_match) a = a_match[0];
    let d_match = ret.match(/(?<=date: ).*/);
    if (d_match) d = d_match[0];

    // Parse page margins from frontmatter
    let pm: PageMargin | undefined;
    const fmMatch = ret.match(/^\n?---[\s\S]*?\n---/);
    if (fmMatch) {
      const fm = fmMatch[0];
      const getFmStr = (key: string) => {
        const m = fm.match(new RegExp(key + ":\\s*(\\d+)"));
        return m ? parseInt(m[1], 10) : null;
      };
      const pw = getFmStr("pageWidth");
      const ph = getFmStr("pageHeight");
      const ml = getFmStr("marginLeft");
      const mr = getFmStr("marginRight");
      const mt = getFmStr("marginTop");
      const mb = getFmStr("marginBottom");
      const mh = getFmStr("marginHeader");
      const mf = getFmStr("marginFooter");

      if (
        pw !== null ||
        ph !== null ||
        ml !== null ||
        mr !== null ||
        mt !== null ||
        mb !== null ||
        mh !== null ||
        mf !== null
      ) {
        pm = {
          width: pw !== null ? pw * 5 : 59528,
          height: ph !== null ? ph * 5 : 84188,
          left: ml !== null ? ml * 5 : 8504,
          right: mr !== null ? mr * 5 : 8504,
          top: mt !== null ? mt * 5 : 5668,
          bottom: mb !== null ? mb * 5 : 4252,
          header: mh !== null ? mh * 5 : 4252,
          footer: mf !== null ? mf * 5 : 4252,
        };
      }
    }

    // remove front matter
    ret = ret.replace(/^\n?---[\s\S]*?\n---/, "");

    // Processing replacements
    ret = ret.replace(
      /\n---/g,
      `\n<P ParaShape="10" Style="0"><TEXT CharShape="0"><CHAR></CHAR></TEXT></P>`,
    );
    ret = ret.replace(/\[(.*?)\]\((.*?)\)/g, "$1 ($2)");

    // Python 버전과 동일한 리스트 처리 (bullet 제거, ParaShape 21-26)
    let BASE_INST_ID = 3102933706;
    ret = ret.replace(
      /\n( {0,24})([-+*]) /g,
      (_match, spaces: string, _bullet: string) => {
        const depth = Math.floor(spaces.length / 4);
        const para = 21 + depth;

        let instAttr = "";
        if (para === 22) {
          BASE_INST_ID++;
          instAttr = ` InstId="${BASE_INST_ID}"`;
        }

        return `\n<P${instAttr} ParaShape="${para}" Style="0"><TEXT CharShape="0"><CHAR>`;
      },
    );

    // 번호 목록 - Python 버전: ParaShape="25"
    ret = ret.replace(
      /\n(\d+)\. /g,
      '\n<P ParaShape="25" Style="0"><TEXT CharShape="0"><CHAR>',
    );
    ret = ret.replace(
      /\n    (\d+)\. /g,
      '\n<P InstId="95545006$1" ParaShape="2" Style="2"><TEXT CharShape="0"><CHAR>',
    );

    ret = ret.replace(/\[ \]/g, "☐");
    ret = ret.replace(/\[x\]/gi, "☑");
    ret = ret.replace(
      /\n(&gt;|>) (.*)/g,
      '\n<P ParaShape="11" Style="0"><TEXT CharShape="8"><CHAR>$2</CHAR></TEXT></P>',
    );

    // Tables (Native MD tables only, HTML are already placeholders)
    ret = ret.replace(/((\n?\|.*\|\s*)+)/gm, (_match, p1) =>
      this.process_table_native(p1),
    );

    // Restore Placeholders in reverse order to handle nesting
    for (let i = this.ph_counter - 1; i >= 0; i--) {
      const key = `__MD2HML_PH_${i}__`;
      if (this.PLACEHOLDERS.hasOwnProperty(key)) {
        // Use split/join to replace all occurrences efficiently just in case
        ret = ret.split(key).join(this.PLACEHOLDERS[key]);
      }
    }

    // Headers
    ret = ret.replace(/\n# /g, "\n" + H1_START);
    ret = ret.replace(/\n## /g, "\n" + H2_START);
    ret = ret.replace(/\n### /g, "\n" + H3_START);
    ret = ret.replace(/\n#### /g, "\n" + H4_START);
    ret = ret.replace(/\n##### /g, "\n" + H5_START);
    ret = ret.replace(/\n###### /g, "\n" + H6_START);

    ret = ret.replace(
      /\n(?=[^<])/g,
      '\n<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>',
    );

    ret = ret.replace(
      /(<CHAR>)([^<]*?)(?=\s*(?:<\/TEXT>|<\/P>|<P|$))/gs,
      "$1$2</CHAR>",
    );

    ret = ret.replace(/(<\/CHAR>)(?!\s*<\/TEXT>)/g, "$1</TEXT></P>");

    ret = ret.replace(/(<CHAR>[^<]*)$/gm, "$1</CHAR></TEXT></P>");

    ret = ret.replace(
      /(<TEXT CharShape="0"><CHAR>)(.*?)(<\/CHAR>)/gs,
      (_match, p1, p2, p3) => p1 + this.processInline(p2) + p3,
    );
    ret = ret.replace(
      /(<TEXT CharShape="0"><CHAR>)(.*?)(<\/CHAR>)/gs,
      (_match, p1, p2, p3) => p1 + this.processInline(p2) + p3,
    );

    let bindataListXml = "";
    if (this.BIN_ITEM_ENTRIES.length > 0) {
      bindataListXml =
        `<BINDATALIST Count="${this.BIN_ITEM_ENTRIES.length}">` +
        this.BIN_ITEM_ENTRIES.join("") +
        `</BINDATALIST>`;
    }

    const additionalBorderFillsXml = this.borderFills
      .map((b) => b.xml)
      .join("\n");
    const headerContent = this.HEADER(
      bindataListXml,
      additionalBorderFillsXml,
      this.borderFills.length,
      pm,
    );

    let bindataStorageXml = "";
    if (this.BIN_STORAGE_ENTRIES.length > 0) {
      bindataStorageXml =
        `<BINDATASTORAGE>` +
        this.BIN_STORAGE_ENTRIES.join("") +
        `</BINDATASTORAGE>`;
    }
    const footerWithBin = `</SECTION></BODY><TAIL>${bindataStorageXml}</TAIL></HWPML>`;

    ret = headerContent + INTRO + ret + footerWithBin;
    ret = ret.replace("___TITLE___", t);
    ret = ret.replace("___AUTHOR___", a);
    ret = ret.replace("___DATE___", d);

    return ret;
  }

  private HEADER(
    bindata_str: string,
    additionalBorderFills: string = "",
    additionalBorderCount: number = 0,
    pm?: PageMargin,
  ): string {
    return getHeaderTemplate(
      bindata_str,
      additionalBorderFills,
      additionalBorderCount,
      pm,
    );
  }

  private process_html_table(table_html: string): string {
    const rows_data: {
      text: string;
      style: string;
      colspan: number;
      rowspan: number;
    }[][] = [];
    let col_count = 0;

    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let rowMatch;
    while ((rowMatch = rowRegex.exec(table_html)) !== null) {
      const rowHtml = rowMatch[1];
      const cellRegex = /<(td|th)([^>]*)>([\s\S]*?)<\/\1>/gi;
      let cellMatch;
      const row_cells = [];
      while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
        const tag = cellMatch[1].toLowerCase();
        const attrs = cellMatch[2];
        const text = cellMatch[3].trim();
        let colspan = 1;
        let rowspan = 1;
        let style = tag; // Use 'th' or 'td' as default style/type marker

        const csMatch = attrs.match(/colspan=["'](\d+)["']/i);
        if (csMatch) colspan = parseInt(csMatch[1]) || 1;

        const rsMatch = attrs.match(/rowspan=["'](\d+)["']/i);
        if (rsMatch) rowspan = parseInt(rsMatch[1]) || 1;

        const styleMatch = attrs.match(/style=["']([^"']+)["']/i);
        if (styleMatch) style += " " + styleMatch[1];

        row_cells.push({ text, style, colspan, rowspan });
      }
      if (row_cells.length > col_count) {
        let current_cols = 0;
        row_cells.forEach((c) => (current_cols += c.colspan));
        if (current_cols > col_count) col_count = current_cols;
      }
      if (row_cells.length > 0) rows_data.push(row_cells);
    }

    if (rows_data.length === 0) return table_html;

    const virtualGrid: boolean[][] = [];
    let max_col_count = 0;

    rows_data.forEach((row, r_idx) => {
      if (!virtualGrid[r_idx]) virtualGrid[r_idx] = [];
      let c_idx_real = 0;

      row.forEach((cell) => {
        // Find next available column in this row
        while (virtualGrid[r_idx][c_idx_real]) {
          c_idx_real++;
        }

        // Mark the grid as occupied for the span of this cell
        for (let r = 0; r < cell.rowspan; r++) {
          const targetRow = r_idx + r;
          if (!virtualGrid[targetRow]) virtualGrid[targetRow] = [];
          for (let c = 0; c < cell.colspan; c++) {
            virtualGrid[targetRow][c_idx_real + c] = true;
          }
        }

        max_col_count = Math.max(max_col_count, c_idx_real + cell.colspan);

        // Temporarily store the calculated colAddr on the cell object
        (cell as any).colAddr = c_idx_real;

        c_idx_real += cell.colspan;
      });
    });

    if (max_col_count > col_count) {
      col_count = max_col_count;
    }

    const total_width = 41954;

    // 각 컬럼의 폭을 결정: 첫 번째 행의 CSS width가 있으면 비율로 정규화, 없으면 균등 분배
    const rawColWidths: number[] = new Array(col_count).fill(0);
    let hasExplicitWidths = false;
    if (rows_data.length > 0) {
      let colCursor = 0;
      for (const cell of rows_data[0]) {
        const styleStr = cell.style.replace(/^(th|td)\s*/, "");
        const wMatch = styleStr.match(/(?:^|;)\s*width\s*:\s*(\d+)\s*px/i);
        if (wMatch) {
          hasExplicitWidths = true;
          const perColPx = parseInt(wMatch[1]) / cell.colspan;
          for (let c = 0; c < cell.colspan && colCursor + c < col_count; c++) {
            rawColWidths[colCursor + c] = perColPx;
          }
        }
        colCursor += cell.colspan;
      }
    }

    let colWidthArray: number[];
    if (hasExplicitWidths && rawColWidths.every((w) => w > 0)) {
      // CSS 폭을 비율로 정규화하여 total_width에 맞춤
      const totalRaw = rawColWidths.reduce((s, w) => s + w, 0);
      colWidthArray = rawColWidths.map((w) =>
        Math.floor((w / totalRaw) * total_width),
      );
      const rounded = colWidthArray.reduce((s, w) => s + w, 0);
      colWidthArray[col_count - 1] += total_width - rounded;
    } else {
      const equalWidth = Math.floor(total_width / col_count);
      colWidthArray = new Array(col_count).fill(equalWidth);
      colWidthArray[col_count - 1] += total_width - equalWidth * col_count;
    }

    let hml_rows = "";
    rows_data.forEach((row, r_idx) => {
      let cells_xml = "";
      row.forEach((cell) => {
        const isTh = cell.style.startsWith("th") || r_idx === 0;
        let borderFillId = isTh ? 4 : 3;
        const isHeader = isTh ? 1 : 0;
        const paraShape = isTh ? 26 : 0;
        const charShape = isTh ? 7 : 0;

        const c_idx_real = (cell as any).colAddr;

        // CSS style에서 border/background 파싱하여 동적 BorderFill 생성
        const styleStr = cell.style.replace(/^(th|td)\s*/, "");
        if (styleStr.trim()) {
          const bf = this.parseCssToHmlBorderFill(styleStr);
          if (bf) {
            const bfId = 4 + this.borderFills.length + 1;
            this.borderFills.push({ id: bfId, xml: bf });
            borderFillId = bfId;
          }
        }

        // PARALIST 안에 치환자(placeholder)가 단독으로 있는 경우 (중첩 표 등)
        // <P><TEXT><CHAR>로 감싸지 않고 그대로 출력하도록 분기
        let p_xml = "";
        if (cell.text.match(/^__MD2HML_PH_\d+__$/)) {
          p_xml = `${cell.text}\n`;
        } else {
          const lines = cell.text.split(/<br\s*\/?>|\n/i);
          lines.forEach((line) => {
            p_xml += `<P ParaShape="${paraShape}" Style="0"><TEXT CharShape="${charShape}"><CHAR>${xmlEscape(line.trim() || " ")}</CHAR></TEXT></P>\n`;
          });
        }

        // 컬럼 폭 배열에서 colspan 범위 합산
        const cellWidth = colWidthArray
          .slice(c_idx_real, c_idx_real + cell.colspan)
          .reduce((s, w) => s + w, 0);
        const cellHeight = 282;

        cells_xml += `<CELL BorderFill="${borderFillId}" ColAddr="${c_idx_real}" ColSpan="${cell.colspan}" Dirty="0" Editable="0" HasMargin="0" Header="${isHeader}" Height="${cellHeight}" Protect="0" RowAddr="${r_idx}" RowSpan="${cell.rowspan}" Width="${cellWidth}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${p_xml.trim()}
</PARALIST>
</CELL>`;
      });
      hml_rows += `<ROW>${cells_xml}</ROW>`;
    });

    const row_count = rows_data.length;

    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="${col_count}" PageBreak="Cell" RepeatHeader="1" RowCount="${row_count}">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="${total_width}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
${hml_rows}
</TABLE><CHAR/></TEXT></P>`;
  }

  /**
   * CSS style 문자열에서 border/background 정보를 파싱하여
   * HML BORDERFILL XML을 생성합니다.
   */
  private parseCssToHmlBorderFill(styleStr: string): string | null {
    const parseBorderProp = (
      prop: string,
    ): { width: string; color: string; none?: boolean } | null => {
      // Check for explicitly "none" or "0"
      const noneRegex = new RegExp(
        `${prop}\\s*:\\s*(?:none|0(?:px)?)(?:\\s*;|\\s*$|\\s)`,
        "i",
      );
      if (noneRegex.test(styleStr)) {
        return { width: "0.1mm", color: "#000000", none: true };
      }

      // "border-left:2px solid #ff0000" or "border:1px solid #000"
      const regex = new RegExp(
        `${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)px\\s+\\w+\\s+([#\\w]+)`,
        "i",
      );
      const m = styleStr.match(regex);
      if (!m) return null;
      const px = parseFloat(m[1]);
      const mm = ((px * 25.4) / 96).toFixed(2) + "mm";
      return { width: mm, color: m[2] };
    };

    // 개별 면 또는 shorthand border
    const left = parseBorderProp("border-left") || parseBorderProp("border");
    const right = parseBorderProp("border-right") || parseBorderProp("border");
    const top = parseBorderProp("border-top") || parseBorderProp("border");
    const bottom =
      parseBorderProp("border-bottom") || parseBorderProp("border");

    // 배경색 파싱
    let faceColor: number | null = null;
    const bgMatch = styleStr.match(
      /(?:background|background-color)\s*:\s*([#\w]+)/i,
    );
    if (bgMatch) {
      const c = bgMatch[1];
      if (c.startsWith("#") && c.length >= 7) {
        const r = parseInt(c.slice(1, 3), 16);
        const g = parseInt(c.slice(3, 5), 16);
        const b = parseInt(c.slice(5, 7), 16);
        faceColor = r | (g << 8) | (b << 16);
      }
    }

    // 아무 스타일도 없으면 null 반환
    if (!left && !right && !top && !bottom && faceColor === null) return null;

    const borderTag = (
      name: string,
      b: { width: string; color: string; none?: boolean } | null,
    ) => {
      if (!b || b.none) return `<${name} Type="None" Width="0.1mm" />`;
      return `<${name} Color="${this.cssColorToHmlInt(b.color)}" Type="Solid" Width="${b.width}" />`;
    };

    const bfId = 4 + this.borderFills.length + 1;
    let xml = `<BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0" CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="${bfId}" Shadow="false" Slash="0" ThreeD="false">\n`;
    xml += borderTag("LEFTBORDER", left) + "\n";
    xml += borderTag("RIGHTBORDER", right) + "\n";
    xml += borderTag("TOPBORDER", top) + "\n";
    xml += borderTag("BOTTOMBORDER", bottom) + "\n";
    xml += '<DIAGONAL Type="Solid" Width="0.1mm" />\n';

    if (faceColor !== null) {
      xml += `<FILLBRUSH>\n<WINDOWBRUSH Alpha="0" FaceColor="${faceColor}" HatchColor="4294967295" />\n</FILLBRUSH>\n`;
    }

    xml += "</BORDERFILL>";
    return xml;
  }

  private cssColorToHmlInt(color: string): number {
    if (color.startsWith("#") && color.length >= 7) {
      const r = parseInt(color.slice(1, 3), 16);
      const g = parseInt(color.slice(3, 5), 16);
      const b = parseInt(color.slice(5, 7), 16);
      return r | (g << 8) | (b << 16);
    }
    return 0; // black fallback
  }
}
