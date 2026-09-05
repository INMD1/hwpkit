import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildPara, buildRoot, buildSheet, buildSpan } from "../../model/builders";
import { DocxEncoder } from "./DocxEncoder";

/**
 * ECMA-376 §17.3.1.12: first-line position = body margin - hanging.
 * DocRoot.indentPt already holds the body margin. Negative first-line positions
 * are representable; forcing left >= hanging would change the source layout.
 */
async function encodeIndent(indentPt: number, firstLineIndentPt: number) {
  const document = buildRoot({}, [
    buildSheet([
      buildPara([buildSpan("들여쓰기")], { indentPt, firstLineIndentPt }),
    ]),
  ]);
  const encoded = await new DocxEncoder().encode(document);
  expect(encoded.ok).toBe(true);
  if (!encoded.ok) throw new Error("encode failed");

  const zip = await JSZip.loadAsync(encoded.data);
  const xml = await zip.file("word/document.xml")!.async("string");
  const ind = /<w:ind [^/]*\/>/.exec(xml)?.[0] ?? "";
  return {
    xml,
    ind,
    warns: encoded.warns ?? [],
    left: Number(/w:left="(\d+)"/.exec(ind)?.[1] ?? 0),
    hanging: Number(/w:hanging="(\d+)"/.exec(ind)?.[1] ?? 0),
  };
}

describe("DocxEncoder hanging indent", () => {
  it("preserves the body margin independently of the hanging width", async () => {
    // 본문 왼쪽 여백 18.3pt(366dxa), 내어쓰기 20pt(400dxa)
    const { left, hanging } = await encodeIndent(18.3, -20);
    expect(hanging).toBe(400);
    expect(left).toBe(366);
    expect(left - hanging).toBe(-34);
  });

  it("preserves a zero body margin with a hanging first line", async () => {
    const { left, hanging } = await encodeIndent(0, -20);
    expect(hanging).toBe(400);
    expect(left).toBe(0);
    expect(left - hanging).toBe(-400);
  });

  it("does not mislabel a negative first-line position as invalid", async () => {
    // 첫 줄이 본문 여백보다 왼쪽인 것은 문서의 표현 가능한 배치다.
    const { warns } = await encodeIndent(18.3, -20);
    expect(warns.filter((w) => w.includes("exceeds w:left"))).toEqual([]);
  });

  it("emits w:firstLine instead of w:hanging for positive first-line indent", async () => {
    const { ind } = await encodeIndent(10, 20);
    expect(ind).toContain('w:firstLine="400"');
    expect(ind).not.toContain("w:hanging");
  });
});
