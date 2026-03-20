import * as cfb from "cfb";
import pako from "pako";
import { CharShape, ParaShape, BorderFill, PageDef } from "./types";
import { parseTextData, applyCharShapes } from "./utils";

function colorRefToHex(v: number) {
  const r = v & 0xff;
  const g = (v >> 8) & 0xff;
  const b = (v >> 16) & 0xff;
  return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}

export class HwpParser {
  private ole: cfb.CFB$Container;
  private charShapes: CharShape[] = [];
  private paraShapes: ParaShape[] = [];
  private pageDefs: PageDef[] = [];
  private borderFills: BorderFill[] = [
    {
      left: { type: 0, thickness: 0, color: "#000000" },
      right: { type: 0, thickness: 0, color: "#000000" },
      top: { type: 0, thickness: 0, color: "#000000" },
      bottom: { type: 0, thickness: 0, color: "#000000" },
    },
  ]; // ID 0 is null
  private compressed: boolean = false;

  constructor(data: Uint8Array) {
    this.ole = cfb.read(data, { type: "buffer" });
  }

  public parse(): string {
    this.parseFileHeader();
    this.parseDocInfo();
    const mdBody = this.parseBodyText();

    // Create YAML frontmatter if page margins exist
    let frontmatter = "";
    if (this.pageDefs.length > 0) {
      const pd = this.pageDefs[0];
      frontmatter = "---\n";
      frontmatter += `pageWidth: ${pd.width}\n`;
      frontmatter += `pageHeight: ${pd.height}\n`;
      frontmatter += `marginLeft: ${pd.leftMargin}\n`;
      frontmatter += `marginRight: ${pd.rightMargin}\n`;
      frontmatter += `marginTop: ${pd.topMargin}\n`;
      frontmatter += `marginBottom: ${pd.bottomMargin}\n`;
      frontmatter += `marginHeader: ${pd.headerMargin}\n`;
      frontmatter += `marginFooter: ${pd.footerMargin}\n`;
      frontmatter += "---\n\n";
    }

    return frontmatter + mdBody;
  }

  private parseFileHeader() {
    const fileHeaderEntry = this.ole.FullPaths.find((p) =>
      p.toLowerCase().endsWith("fileheader"),
    );
    if (!fileHeaderEntry) {
      throw new Error("Not a valid HWP 5.0 file (FileHeader not found)");
    }

    const fileHeader = cfb.find(this.ole, fileHeaderEntry);
    if (!fileHeader || !fileHeader.content || fileHeader.content.length < 40) {
      throw new Error("Invalid FileHeader");
    }

    const headerContent = new Uint8Array(fileHeader.content as any);
    const headerView = new DataView(
      headerContent.buffer,
      headerContent.byteOffset,
      headerContent.byteLength,
    );
    const attr = headerView.getUint32(36, true);
    this.compressed = (attr & 1) !== 0;
  }

  private parseDocInfo() {
    const docInfoEntry = this.ole.FullPaths.find((p) =>
      p.toLowerCase().endsWith("docinfo"),
    );
    if (!docInfoEntry) return;

    const docInfoFile = cfb.find(this.ole, docInfoEntry);
    if (!docInfoFile || !docInfoFile.content) return;

    let stream = new Uint8Array(docInfoFile.content as any);
    if (this.compressed) {
      try {
        stream = pako.inflate(stream, { raw: true });
      } catch (e) {}
    }

    const streamView = new DataView(
      stream.buffer,
      stream.byteOffset,
      stream.byteLength,
    );
    let offset = 0;

    while (offset < stream.length) {
      if (offset + 4 > stream.length) break;
      const header = streamView.getUint32(offset, true);
      offset += 4;
      const tagId = header & 0x3ff;
      let size = (header >>> 20) & 0xfff;

      if (size === 0xfff) {
        if (offset + 4 > stream.length) break;
        size = streamView.getUint32(offset, true);
        offset += 4;
      }

      if (tagId === 18) {
        // HWPTAG_CHAR_SHAPE
        if (offset + 60 <= stream.length) {
          const property = streamView.getUint32(offset + 36, true);
          const isItalic = (property & 1) !== 0; // Bit 0
          const isBold = (property & 2) !== 0; // Bit 1
          const isUnderline = (property & 4) !== 0; // Bit 2
          const isStrike = (property & 8) !== 0; // Bit 3

          const baseSizeInHwpUnit = streamView.getUint32(offset + 56, true); // baseSize (HWPUNITs)
          const baseSizePt = baseSizeInHwpUnit / 100;

          this.charShapes.push({
            isBold,
            isItalic,
            isUnderline,
            isStrike,
            baseSize: baseSizePt,
          });
        } else if (offset + 40 <= stream.length) {
          const property = streamView.getUint32(offset + 36, true);
          const isBold = (property & 2) !== 0;
          this.charShapes.push({ isBold });
        }
      } else if (tagId === 22) {
        // HWPTAG_PARA_SHAPE
        if (offset + 4 <= stream.length) {
          const prop1 = streamView.getUint32(offset, true);
          const align = (prop1 >>> 2) & 0x7;
          this.paraShapes.push({ alignment: align });
        }
      } else if (tagId === 19) {
        let faceColor: string | undefined;

        const borderTypes = [];
        const borderWidths = [];
        const borderColors = [];

        for (let i = 0; i < 4; i++) {
          borderTypes.push(streamView.getUint8(offset + 2 + i));
          borderWidths.push(streamView.getUint8(offset + 6 + i));
          borderColors.push(
            colorRefToHex(streamView.getUint32(offset + 10 + i * 4, true)),
          );
        }

        const left = {
          type: borderTypes[0],
          thickness: borderWidths[0],
          color: borderColors[0],
        };
        const right = {
          type: borderTypes[1],
          thickness: borderWidths[1],
          color: borderColors[1],
        };
        const top = {
          type: borderTypes[2],
          thickness: borderWidths[2],
          color: borderColors[2],
        };
        const bottom = {
          type: borderTypes[3],
          thickness: borderWidths[3],
          color: borderColors[3],
        };

        const end = offset + size;
        let pos = offset + 32;

        if (pos + 4 <= end) {
          const fillType = streamView.getUint32(pos, true);
          pos += 4;

          if (fillType & 1 && pos + 4 <= end) {
            const col = streamView.getUint32(pos, true);
            pos += 4;

            if (col !== 0xffffffff) faceColor = colorRefToHex(col);
          }
        }

        this.borderFills.push({ left, right, top, bottom, faceColor });
      }

      offset += size;
    }
  }

  private parseBodyText(): string {
    const sections = this.ole.FullPaths.filter((p) =>
      p.toLowerCase().includes("bodytext/section"),
    ).sort((a, b) => {
      const aMatch = a.toLowerCase().match(/section(\d+)/);
      const bMatch = b.toLowerCase().match(/section(\d+)/);
      const aNum = aMatch ? parseInt(aMatch[1], 10) : 0;
      const bNum = bMatch ? parseInt(bMatch[1], 10) : 0;
      return aNum - bNum;
    });

    if (sections.length === 0) {
      throw new Error("BodyText sections not found");
    }

    let docText = "";

    for (const secPath of sections) {
      const secEntry = cfb.find(this.ole, secPath);
      if (!secEntry || !secEntry.content) continue;

      let stream = new Uint8Array(secEntry.content as any);

      if (this.compressed) {
        try {
          stream = pako.inflate(stream, { raw: true });
        } catch (e) {
          console.error(`Decompression failed for ${secPath}`, e);
          continue;
        }
      }

      const streamView = new DataView(
        stream.buffer,
        stream.byteOffset,
        stream.byteLength,
      );
      let offset = 0;

      let currentAlign = 0;
      let currentParaTextData: Uint8Array | null = null;
      let currentParaCharShapes: { pos: number; shapeId: number }[] = [];
      let currentCharShapeCount = 0;

      // Table tracking state
      let inTable = false;
      let tableData: any[][] = [];
      let currentCellText = "";
      let currentCellRowIdx = -1;
      let currentCellColIdx = -1;
      let currentCellColSpan = 1;
      let currentCellRowSpan = 1;
      let currentCellBorderFillId = 0;

      let tableLevel = -1;

      while (offset < stream.length) {
        if (offset + 4 > stream.length) break;

        const header = streamView.getUint32(offset, true);
        offset += 4;

        const tagId = header & 0x3ff;
        const level = (header >>> 10) & 0x3ff;
        let size = (header >>> 20) & 0xfff;

        if (size === 0xfff) {
          if (offset + 4 > stream.length) break;
          size = streamView.getUint32(offset, true);
          offset += 4;
        }

        if (size === 0) {
          break;
        }

        // If we hit a tag with level < tableLevel, the table is finished
        if (inTable && level < tableLevel) {
          if (currentParaTextData) {
            const paraText =
              currentParaCharShapes.length > 0
                ? applyCharShapes(
                    currentParaTextData,
                    currentParaCharShapes,
                    this.charShapes,
                  )
                : parseTextData(currentParaTextData);
            if (paraText) {
              let alignedText = paraText;
              if (currentAlign === 2) {
                alignedText = `<div align="right">${alignedText}</div>`;
              } else if (currentAlign === 3) {
                alignedText = `<div align="center">${alignedText}</div>`;
              }
              currentCellText += (currentCellText ? "<br>" : "") + alignedText;
            }
            currentParaTextData = null;
            currentParaCharShapes = [];
          }

          if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
            while (tableData.length <= currentCellRowIdx) tableData.push([]);
            while (tableData[currentCellRowIdx].length <= currentCellColIdx)
              tableData[currentCellRowIdx].push(null);
            tableData[currentCellRowIdx][currentCellColIdx] = {
              text: currentCellText.trim(),
              colSpan: currentCellColSpan,
              rowSpan: currentCellRowSpan,
              borderFillId: currentCellBorderFillId,
            };
          }

          if (tableData.length > 0) {
            let maxCols = 0;
            for (const row of tableData) {
              if (row.length > maxCols) maxCols = row.length;
            }

            docText +=
              '\n<table style="border-collapse:collapse;border-spacing:0">\n';
            for (let r = 0; r < tableData.length; r++) {
              docText += "  <tr>\n";
              for (let c = 0; c < maxCols; c++) {
                const cellObj = tableData[r][c];
                if (cellObj) {
                  const cellStyle = this.buildCellStyle(cellObj.borderFillId);
                  let attrs = ` style="${cellStyle}"`;
                  if (cellObj.colSpan > 1)
                    attrs += ` colspan="${cellObj.colSpan}"`;
                  if (cellObj.rowSpan > 1)
                    attrs += ` rowspan="${cellObj.rowSpan}"`;
                  docText += `    <td${attrs}>${cellObj.text}</td>\n`;
                }
              }
              docText += "  </tr>\n";
            }
            docText += "</table>\n\n";
          }

          inTable = false;
          tableData = [];
          currentCellText = "";
          currentCellRowIdx = -1;
          currentCellColIdx = -1;
          currentCellBorderFillId = 0;
          tableLevel = -1;
        }

        if (tagId === 66) {
          // HWPTAG_PARA_HEADER
          if (currentParaTextData && currentParaCharShapes.length > 0) {
            const paraText = applyCharShapes(
              currentParaTextData,
              currentParaCharShapes,
              this.charShapes,
            );
            if (paraText) {
              let alignedText = paraText;
              if (currentAlign === 2) {
                alignedText = `<div align="right">${alignedText}</div>`;
              } else if (currentAlign === 3) {
                alignedText = `<div align="center">${alignedText}</div>`;
              }
              if (inTable) {
                currentCellText +=
                  (currentCellText ? "<br>" : "") + alignedText;
              } else {
                docText += alignedText + "\n\n";
              }
            }
            currentParaTextData = null;
            currentParaCharShapes = [];
          } else if (currentParaTextData) {
            const paraText = parseTextData(currentParaTextData);
            if (paraText) {
              let alignedText = paraText;
              if (currentAlign === 2) {
                alignedText = `<div align="right">${alignedText}</div>`;
              } else if (currentAlign === 3) {
                alignedText = `<div align="center">${alignedText}</div>`;
              }
              if (inTable) {
                currentCellText +=
                  (currentCellText ? "<br>" : "") + alignedText;
              } else {
                docText += alignedText + "\n\n";
              }
            }
            currentParaTextData = null;
          }

          if (offset + size <= stream.length) {
            try {
              const paraShapeId = streamView.getUint32(offset + 8, true);
              if (paraShapeId < this.paraShapes.length) {
                currentAlign = this.paraShapes[paraShapeId].alignment;
              } else {
                currentAlign = 0;
              }

              if (offset + 24 <= stream.length) {
                currentCharShapeCount = streamView.getUint16(offset + 22, true);
              } else {
                currentCharShapeCount = 0;
              }
              currentParaCharShapes = [];
            } catch (e) {}
          }
        } else if (tagId === 67) {
          // HWPTAG_PARA_TEXT
          if (offset + size <= stream.length) {
            currentParaTextData = stream.slice(offset, offset + size);
          }
        } else if (tagId === 68) {
          // HWPTAG_PARA_CHAR_SHAPE
          if (offset + size <= stream.length) {
            for (let i = 0; i < currentCharShapeCount; i++) {
              const shapeOffset = offset + i * 8;
              if (
                shapeOffset + 8 <= offset + size &&
                shapeOffset + 8 <= stream.length
              ) {
                const pos = streamView.getUint32(shapeOffset, true);
                const shapeId = streamView.getUint32(shapeOffset + 4, true);
                currentParaCharShapes.push({ pos, shapeId });
              }
            }
          }
        } else if (tagId === 73) {
          // HWPTAG_PAGE_DEF
          if (size >= 40) {
            this.pageDefs.push({
              width: Math.round(streamView.getUint32(offset, true) / 5),
              height: Math.round(streamView.getUint32(offset + 4, true) / 5),
              leftMargin: Math.round(
                streamView.getUint32(offset + 8, true) / 5,
              ),
              rightMargin: Math.round(
                streamView.getUint32(offset + 12, true) / 5,
              ),
              topMargin: Math.round(
                streamView.getUint32(offset + 16, true) / 5,
              ),
              bottomMargin: Math.round(
                streamView.getUint32(offset + 20, true) / 5,
              ),
              headerMargin: Math.round(
                streamView.getUint32(offset + 24, true) / 5,
              ),
              footerMargin: Math.round(
                streamView.getUint32(offset + 28, true) / 5,
              ),
            });
          }
        } else if (tagId === 77) {
          // HWPTAG_TABLE
          if (currentParaTextData) {
            const paraText =
              currentParaCharShapes.length > 0
                ? applyCharShapes(
                    currentParaTextData,
                    currentParaCharShapes,
                    this.charShapes,
                  )
                : parseTextData(currentParaTextData);
            if (paraText) {
              let alignedText = paraText;
              if (currentAlign === 2) {
                alignedText = `<div align="right">${alignedText}</div>`;
              } else if (currentAlign === 3) {
                alignedText = `<div align="center">${alignedText}</div>`;
              }
              docText += alignedText + "\n\n";
            }
            currentParaTextData = null;
            currentParaCharShapes = [];
          }

          if (offset + size <= stream.length) {
            inTable = true;
            tableData = [];
            currentCellText = "";
            currentCellRowIdx = -1;
            currentCellColIdx = -1;
            currentCellBorderFillId = 0;
            tableLevel = level;
          }
        } else if (tagId === 72) {
          // HWPTAG_LIST_HEADER
          if (offset + size <= stream.length) {
            if (inTable) {
              if (currentParaTextData) {
                const paraText =
                  currentParaCharShapes.length > 0
                    ? applyCharShapes(
                        currentParaTextData,
                        currentParaCharShapes,
                        this.charShapes,
                      )
                    : parseTextData(currentParaTextData);
                if (paraText) {
                  let alignedText = paraText;
                  if (currentAlign === 2) {
                    alignedText = `<div align="right">${alignedText}</div>`;
                  } else if (currentAlign === 3) {
                    alignedText = `<div align="center">${alignedText}</div>`;
                  }
                  currentCellText +=
                    (currentCellText ? "<br>" : "") + alignedText;
                }
                currentParaTextData = null;
                currentParaCharShapes = [];
              }

              if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
                while (tableData.length <= currentCellRowIdx)
                  tableData.push([]);
                while (tableData[currentCellRowIdx].length <= currentCellColIdx)
                  tableData[currentCellRowIdx].push(null);
                tableData[currentCellRowIdx][currentCellColIdx] = {
                  text: currentCellText.trim(),
                  colSpan: currentCellColSpan,
                  rowSpan: currentCellRowSpan,
                  borderFillId: currentCellBorderFillId,
                };
              }
              currentCellText = "";

              if (size >= 34) {
                const cellAttrOffset = offset + 8;
                currentCellColIdx = streamView.getUint16(cellAttrOffset, true);
                currentCellRowIdx = streamView.getUint16(
                  cellAttrOffset + 2,
                  true,
                );
                currentCellColSpan = streamView.getUint16(
                  cellAttrOffset + 4,
                  true,
                );
                currentCellRowSpan = streamView.getUint16(
                  cellAttrOffset + 6,
                  true,
                );
                currentCellBorderFillId = streamView.getUint16(
                  offset + 32,
                  true,
                );
              } else {
                currentCellColIdx = 0;
                currentCellRowIdx = 0;
                currentCellColSpan = 1;
                currentCellRowSpan = 1;
                currentCellBorderFillId = 0;
              }
            }
          }
        }

        offset += size;
      }

      if (currentParaTextData) {
        const paraText =
          currentParaCharShapes.length > 0
            ? applyCharShapes(
                currentParaTextData,
                currentParaCharShapes,
                this.charShapes,
              )
            : parseTextData(currentParaTextData);

        if (paraText) {
          let alignedText = paraText;
          if (currentAlign === 2) {
            alignedText = `<div align="right">${alignedText}</div>`;
          } else if (currentAlign === 3) {
            alignedText = `<div align="center">${alignedText}</div>`;
          }
          if (inTable) {
            currentCellText += (currentCellText ? "<br>" : "") + alignedText;
          } else {
            docText += alignedText + "\n\n";
          }
        }
        currentParaTextData = null;
        currentParaCharShapes = [];
      }

      if (inTable) {
        // 마지막 셀 저장 (내용이 없어도 저장)
        if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
          while (tableData.length <= currentCellRowIdx) tableData.push([]);
          while (tableData[currentCellRowIdx].length <= currentCellColIdx)
            tableData[currentCellRowIdx].push(null);
          if (!tableData[currentCellRowIdx][currentCellColIdx]) {
            tableData[currentCellRowIdx][currentCellColIdx] = {
              text: currentCellText.trim(),
              colSpan: currentCellColSpan,
              rowSpan: currentCellRowSpan,
              borderFillId: currentCellBorderFillId,
            };
          }
        }

        if (tableData.length > 0) {
          let maxCols = 0;
          for (const row of tableData) {
            if (row.length > maxCols) maxCols = row.length;
          }

          docText +=
            '\n<table style="border-collapse: collapse; width: auto;">\n';
          for (let r = 0; r < tableData.length; r++) {
            docText += "  <tr>\n";
            for (let c = 0; c < maxCols; c++) {
              const cellObj = tableData[r][c];
              if (cellObj) {
                const cellStyle = this.buildCellStyle(cellObj.borderFillId);
                let attrs = ` style="${cellStyle}"`;
                if (cellObj.colSpan > 1)
                  attrs += ` colspan="${cellObj.colSpan}"`;
                if (cellObj.rowSpan > 1)
                  attrs += ` rowspan="${cellObj.rowSpan}"`;
                docText += `    <td${attrs}>${cellObj.text}</td>\n`;
              }
            }
            docText += "  </tr>\n";
          }
          docText += "</table>\n\n";
        }
        inTable = false;
        tableData = [];
        currentCellText = "";
        currentCellRowIdx = -1;
        currentCellColIdx = -1;
        currentCellColSpan = 1;
        currentCellRowSpan = 1;
        currentCellBorderFillId = 0;
        tableLevel = -1;
      }
    }

    return docText.trim();
  }

  // TODO: borderFillId를 사용하여 스타일을 생성해야함
  private buildCellStyle(borderFillId?: number): string {
    const styles: string[] = ["padding:8px"];

    if (
      borderFillId === undefined ||
      borderFillId < 0 ||
      borderFillId >= this.borderFills.length
    ) {
      return styles.join(";");
    }

    const bf = this.borderFills[borderFillId];
    if (!bf) return styles.join(";");

    /* ---------- HWP → CSS 스타일 변환 ---------- */

    function hwpBorderStyle(t: number) {
      switch (t) {
        case 0:
          return "solid";
        case 1:
          return "dashed";
        case 2:
          return "dotted";
        case 7:
          return "double";
        case 11:
          return "wavy";
        default:
          return "solid";
      }
    }

    /* ---------- mm → px ---------- */

    const WIDTH_MM = [
      0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1, 1.5, 2, 3, 4, 5,
    ];

    /* ---------- side renderer ---------- */

    const borderSide = (side: string, b: any) => {
      if (!b || b.thickness === 0) {
        styles.push(`border-${side}:none`);
        return;
      }

      const mm = WIDTH_MM[b.thickness] ?? 0.1;

      let px = (mm * 96) / 25.4;

      const style = hwpBorderStyle(b.type);

      /* double 보정 */
      if (style === "double") px /= 3;

      /* 너무 얇은 선 보정 */
      px = Math.max(1, Math.round(px * 10) / 10);
      //TODO 추후 색 보정 필요
      styles.push(`border-${side}:${px}px ${style}`);
    };

    borderSide("left", bf.left);
    borderSide("right", bf.right);
    borderSide("top", bf.top);
    borderSide("bottom", bf.bottom);

    /* ---------- background ---------- */

    if (bf.faceColor) styles.push(`background:${bf.faceColor}`);

    return styles.join(";");
  }
}
