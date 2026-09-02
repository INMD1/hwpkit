import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildPara, buildRoot, buildSheet, buildSpan } from "../../model/builders";
import { DocxEncoder } from "./DocxEncoder";

/**
 * w:hanging 은 w:left 기준에서 첫 줄을 왼쪽으로 당기는 값이다.
 * 따라서 DocxEncoder 는 본문 왼쪽 여백에 hanging 폭을 더해 w:left 를 방출해야 하며,
 * 방출된 w:left 는 항상 w:hanging 이상이어야 한다.
 * (hwpkit_research/README.md 문단 여백 계약과 동일한 불변식)
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
  it("adds the hanging width to w:left so w:left >= w:hanging", async () => {
    // 본문 왼쪽 여백 18.3pt(366dxa), 내어쓰기 20pt(400dxa)
    const { left, hanging } = await encodeIndent(18.3, -20);
    expect(hanging).toBe(400);
    expect(left).toBe(366 + 400);
    expect(left).toBeGreaterThanOrEqual(hanging);
  });

  it("keeps w:left >= w:hanging even with no body left margin", async () => {
    const { left, hanging } = await encodeIndent(0, -20);
    expect(hanging).toBe(400);
    expect(left).toBe(400);
    expect(left).toBeGreaterThanOrEqual(hanging);
  });

  it("does not warn about a violation it never emits", async () => {
    // 보정 덕분에 실제 출력은 항상 유효하므로 exceeds 경고가 있으면 안 된다.
    const { warns } = await encodeIndent(18.3, -20);
    expect(warns.filter((w) => w.includes("exceeds w:left"))).toEqual([]);
  });

  it("emits w:firstLine instead of w:hanging for positive first-line indent", async () => {
    const { ind } = await encodeIndent(10, 20);
    expect(ind).toContain('w:firstLine="400"');
    expect(ind).not.toContain("w:hanging");
  });
});
