import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  buildCell,
  buildGrid,
  buildPara,
  buildRoot,
  buildRow,
  buildSheet,
  buildSpan,
} from "../model/builders";
import { DocxDecoder } from "./docx/DocxDecoder";
import { DocxEncoder } from "../encoders/docx/DocxEncoder";
import { HwpScanner } from "./hwp/HwpScanner";
import { HwpEncoder } from "../encoders/hwp/HwpEncoder";
import { HwpxDecoder } from "./hwpx/HwpxDecoder";
import { HwpxEncoder } from "../encoders/hwpx/HwpxEncoder";

const W =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("paragraph and table layout defaults", () => {
  it("keeps Word's paragraph defaults and explicit zero spacing distinct", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    );
    zip.file(
      "word/styles.xml",
      `<?xml version="1.0"?>
      <w:styles xmlns:w="${W}">
        <w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
      </w:styles>`,
    );
    zip.file(
      "word/document.xml",
      `<?xml version="1.0"?>
      <w:document xmlns:w="${W}"><w:body>
        <w:p><w:r><w:t>default</w:t></w:r></w:p>
        <w:p><w:pPr><w:spacing w:after="0" w:line="300" w:lineRule="exact"/></w:pPr>
          <w:r><w:t>explicit</w:t></w:r>
        </w:p>
      </w:body></w:document>`,
    );

    const encoded = await zip.generateAsync({ type: "uint8array" });
    const decoded = await new DocxDecoder().decode(encoded);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;

    const [defaults, explicit] = decoded.data.kids[0].kids.filter(
      (node) => node.tag === "para",
    );
    expect(defaults.props.spaceBefore).toBe(0);
    expect(defaults.props.spaceAfter).toBe(8);
    expect(defaults.props.lineHeight).toBe(1.15);
    expect(explicit.props.spaceAfter).toBe(0);
    expect(explicit.props.lineHeightFixed).toBe(15);
    expect(explicit.props.lineHeightRule).toBe("exact");
  });

  it("writes table-level cell padding and vertical alignment explicitly", async () => {
    const paragraph = buildPara([buildSpan("셀")], {
      spaceBefore: 0,
      spaceAfter: 0,
      lineHeight: 1.6,
    });
    const table = buildGrid(
      [buildRow([buildCell([paragraph], { props: { va: "mid" } })])],
      {
        colWidths: [100],
        cellPadT: 1.41,
        cellPadB: 1.41,
        cellPadL: 1.41,
        cellPadR: 1.41,
      },
    );
    const document = buildRoot({}, [buildSheet([table])]);
    const encoded = await new DocxEncoder().encode(document);
    expect(encoded.ok).toBe(true);
    if (!encoded.ok) return;

    const zip = await JSZip.loadAsync(encoded.data);
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(xml).toContain(
      '<w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="28" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar>',
    );
    expect(xml).toContain(
      '<w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/>',
    );
  });

  it("uses Hancom's truncation when converting percentage line spacing to Word", async () => {
    const paragraph = buildPara([buildSpan("132%")], {
      spaceBefore: 0,
      spaceAfter: 0,
      lineHeight: 1.32,
    });
    const document = buildRoot({}, [buildSheet([paragraph])]);
    const encoded = await new DocxEncoder().encode(document);
    expect(encoded.ok).toBe(true);
    if (!encoded.ok) return;

    const zip = await JSZip.loadAsync(encoded.data);
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain(
      '<w:spacing w:before="0" w:after="0" w:line="316" w:lineRule="auto"/>',
    );
  });

  it("round-trips HWP cell alignment and inherited table padding", async () => {
    const table = buildGrid(
      [
        buildRow([
          buildCell([buildPara([buildSpan("가운데")])], {
            props: { va: "mid" },
          }),
        ]),
      ],
      {
        colWidths: [100],
        cellPadT: 1.41,
        cellPadB: 1.41,
        cellPadL: 1.41,
        cellPadR: 1.41,
      },
    );
    const document = buildRoot({}, [buildSheet([table])]);
    const encoded = await new HwpEncoder().encode(document);
    expect(encoded.ok).toBe(true);
    if (!encoded.ok) return;

    const decoded = await new HwpScanner().decode(encoded.data);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    const grid = decoded.data.kids[0].kids.find(
      (node) => node.tag === "grid",
    );
    expect(grid?.tag).toBe("grid");
    if (!grid || grid.tag !== "grid") return;
    expect(grid.props.cellPadL).toBeCloseTo(1.41);
    expect(grid.props.cellPadR).toBeCloseTo(1.41);
    expect(grid.kids[0].kids[0].props.va).toBe("mid");
  });

  it("round-trips HWP fixed line spacing without doubling it", async () => {
    const paragraph = buildPara([buildSpan("고정 줄간격")], {
      spaceBefore: 3,
      spaceAfter: 1,
      lineHeightFixed: 15,
      lineHeightRule: "atLeast",
    });
    const document = buildRoot({}, [buildSheet([paragraph])]);
    const encoded = await new HwpEncoder().encode(document);
    expect(encoded.ok).toBe(true);
    if (!encoded.ok) return;

    const decoded = await new HwpScanner().decode(encoded.data);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    const decodedParagraph = decoded.data.kids[0].kids.find(
      (node) => node.tag === "para",
    );
    expect(decodedParagraph?.tag).toBe("para");
    if (!decodedParagraph || decodedParagraph.tag !== "para") return;
    expect(decodedParagraph.props.spaceBefore).toBeCloseTo(3);
    expect(decodedParagraph.props.spaceAfter).toBeCloseTo(1);
    expect(decodedParagraph.props.lineHeightFixed).toBeCloseTo(15);
    expect(decodedParagraph.props.lineHeightRule).toBe("atLeast");
  });

  it("round-trips HWPX paraPr spacing in its doubled storage unit", async () => {
    const paragraph = buildPara([buildSpan("HWPX 문단")], {
      indentPt: 28.1,
      indentRightPt: 2,
      firstLineIndentPt: -28.1,
      spaceBefore: 3,
      spaceAfter: 1,
      lineHeightFixed: 15,
      lineHeightRule: "atLeast",
    });
    const document = buildRoot({}, [buildSheet([paragraph])]);
    const encoded = await new HwpxEncoder().encode(document);
    expect(encoded.ok).toBe(true);
    if (!encoded.ok) return;

    const zip = await JSZip.loadAsync(encoded.data);
    const headerXml = await zip.file("Contents/header.xml")!.async("string");
    expect(headerXml).toContain('<hc:intent value="-5620" unit="HWPUNIT"/>');
    // Body margin 28.1pt with -28.1pt hanging starts its first line at zero.
    expect(headerXml).toContain('<hc:intent value="-5620" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/>');
    expect(headerXml).toContain('<hc:right value="400" unit="HWPUNIT"/>');
    expect(headerXml).toContain('<hc:prev value="600" unit="HWPUNIT"/>');
    expect(headerXml).toContain('<hc:next value="200" unit="HWPUNIT"/>');
    expect(headerXml).toContain(
      '<hh:lineSpacing type="AT_LEAST" value="3000" unit="HWPUNIT"/>',
    );

    const decoded = await new HwpxDecoder().decode(encoded.data);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    const decodedParagraph = decoded.data.kids[0].kids.find(
      (node) => node.tag === "para",
    );
    expect(decodedParagraph?.tag).toBe("para");
    if (!decodedParagraph || decodedParagraph.tag !== "para") return;
    expect(decodedParagraph.props.indentPt).toBeCloseTo(28.1);
    expect(decodedParagraph.props.indentRightPt).toBeCloseTo(2);
    expect(decodedParagraph.props.firstLineIndentPt).toBeCloseTo(-28.1);
    expect(decodedParagraph.props.spaceBefore).toBeCloseTo(3);
    expect(decodedParagraph.props.spaceAfter).toBeCloseTo(1);
    expect(decodedParagraph.props.lineHeightFixed).toBeCloseTo(15);
    expect(decodedParagraph.props.lineHeightRule).toBe("atLeast");
  });
});
