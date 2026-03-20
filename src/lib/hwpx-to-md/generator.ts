import { MdConversionContext } from "./context";
import { CharProperty } from "./types";

export class HwpxMdGenerator {
  private ctx: MdConversionContext;

  constructor(ctx: MdConversionContext) {
    this.ctx = ctx;
  }

  public convertSection(xml: string, isFirstSection = false): string {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, "text/xml");
    const root = doc.documentElement;
    let md = "";

    // Extract page margins if it's the first section
    if (isFirstSection) {
      const secPrList = root.getElementsByTagNameNS(this.ctx.ns.hp, "secPr");
      if (secPrList && secPrList.length > 0) {
        const secPr = secPrList[0];
        const pagePrList = secPr.getElementsByTagNameNS(
          this.ctx.ns.hp,
          "pagePr",
        );
        if (pagePrList && pagePrList.length > 0) {
          const pagePr = pagePrList[0];
          const width = parseInt(pagePr.getAttribute("width") || "0", 10);
          const height = parseInt(pagePr.getAttribute("height") || "0", 10);

          let leftMargin = 0,
            rightMargin = 0,
            topMargin = 0,
            bottomMargin = 0,
            headerMargin = 0,
            footerMargin = 0;
          const marginElems = pagePr.getElementsByTagNameNS(
            this.ctx.ns.hp,
            "margin",
          );
          if (marginElems && marginElems.length > 0) {
            const marginList = marginElems[0].children;
            for (let j = 0; j < marginList.length; j++) {
              const child = marginList[j];
              const val = parseInt(child.getAttribute("val") || "0", 10);
              switch (child.localName) {
                case "left":
                  leftMargin = val;
                  break;
                case "right":
                  rightMargin = val;
                  break;
                case "top":
                  topMargin = val;
                  break;
                case "bottom":
                  bottomMargin = val;
                  break;
                case "header":
                  headerMargin = val;
                  break;
                case "footer":
                  footerMargin = val;
                  break;
              }
            }
          }

          let frontmatter = "---\n";
          frontmatter += `pageWidth: ${Math.round(width / 5)}\n`;
          frontmatter += `pageHeight: ${Math.round(height / 5)}\n`;
          frontmatter += `marginLeft: ${Math.round(leftMargin / 5)}\n`;
          frontmatter += `marginRight: ${Math.round(rightMargin / 5)}\n`;
          frontmatter += `marginTop: ${Math.round(topMargin / 5)}\n`;
          frontmatter += `marginBottom: ${Math.round(bottomMargin / 5)}\n`;
          frontmatter += `marginHeader: ${Math.round(headerMargin / 5)}\n`;
          frontmatter += `marginFooter: ${Math.round(footerMargin / 5)}\n`;
          frontmatter += "---\n\n";

          md += frontmatter;
        }
      }
    }

    const children = root.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child.nodeType !== 1) continue;
      if ((child as Element).localName === "p") {
        md += this.convertParagraph(child as Element);
      }
    }

    return md.trim();
  }

  private convertParagraph(pElem: Element): string {
    const styleIDRef = pElem.getAttribute("styleIDRef") || "";

    // 제목 레벨 판별
    let headingLevel = 0;
    if (styleIDRef) {
      const style = this.ctx.styles.get(styleIDRef);
      if (style && style.type === "HEADING" && style.level > 0) {
        headingLevel = Math.min(style.level, 6);
      }
    }

    let paraText = "";
    let tableBlocks = "";

    // 단락 내 최상위 표를 직접 탐색 (run 내부이든 ctrl 내부이든 상관없이)
    // 단, 다른 표의 셀(tc) 안에 있는 표는 제외
    const allTables = Array.from(
      pElem.getElementsByTagNameNS(this.ctx.ns.hp, "tbl"),
    );
    const topLevelTables = allTables.filter((tbl) => {
      let p = tbl.parentNode;
      while (p && p !== pElem) {
        if ((p as Element).localName === "tc") return false; // 다른 표 셀 내부
        p = p.parentNode;
      }
      return true;
    });

    for (const tbl of topLevelTables) {
      tableBlocks += this.convertTable(tbl);
    }

    // 일반 텍스트/이미지/도형 처리 (표를 포함하는 run은 건너뜀)
    const runs = pElem.getElementsByTagNameNS(this.ctx.ns.hp, "run");
    for (const run of Array.from(runs)) {
      if (run.parentNode !== pElem) continue;

      // 이 run이 표를 포함하면 텍스트 처리 건너뜀
      if (run.getElementsByTagNameNS(this.ctx.ns.hp, "tbl").length > 0) {
        continue;
      }

      // 이미지 처리
      const pics = run.getElementsByTagNameNS(this.ctx.ns.hp, "pic");
      if (pics.length > 0) {
        for (const pic of Array.from(pics)) {
          paraText += this.convertImage(pic);
        }
        continue;
      }

      // 도형(rect) 내 텍스트 처리
      const rects = run.getElementsByTagNameNS(this.ctx.ns.hp, "rect");
      if (rects.length > 0) {
        for (const rect of Array.from(rects)) {
          paraText += this.convertRect(rect);
        }
        continue;
      }

      // secPr 건너뜀
      const secPr = run.getElementsByTagNameNS(this.ctx.ns.hp, "secPr");
      if (secPr.length > 0) continue;

      // 일반 텍스트
      const charPrIDRef = run.getAttribute("charPrIDRef") || "0";
      const cp = this.ctx.charProperties.get(charPrIDRef);

      let runText = "";
      const tElems = run.getElementsByTagNameNS(this.ctx.ns.hp, "t");
      for (const tElem of Array.from(tElems)) {
        runText += this.getTextContent(tElem);
      }

      if (runText && cp) {
        runText = this.applyInlineFormatting(runText, cp);
      }

      paraText += runText;
    }

    if (tableBlocks) {
      return tableBlocks + (paraText.trim() ? paraText.trim() + "\n\n" : "");
    }

    if (!paraText.trim()) {
      return "\n";
    }

    if (headingLevel > 0) {
      return "#".repeat(headingLevel) + " " + paraText.trim() + "\n\n";
    }

    return paraText.trim() + "\n\n";
  }

  private applyInlineFormatting(text: string, cp: CharProperty): string {
    const leading = text.match(/^(\s*)/)?.[1] ?? "";
    const trailing = text.match(/(\s*)$/)?.[1] ?? "";
    const inner = text.trim();

    if (!inner) return text;

    let result = inner;

    // superscript/subscript는 HTML 인라인으로 처리
    if (cp.supscript) return `${leading}<sup>${result}</sup>${trailing}`;
    if (cp.subscript) return `${leading}<sub>${result}</sub>${trailing}`;

    if (cp.strikeout) result = `~~${result}~~`;
    if (cp.underline) result = `<u>${result}</u>`;

    if (cp.bold && cp.italic) result = `***${result}***`;
    else if (cp.bold) result = `**${result}**`;
    else if (cp.italic) result = `*${result}*`;

    return leading + result + trailing;
  }

  private getTextContent(tElem: Element): string {
    let text = "";
    for (const node of Array.from(tElem.childNodes)) {
      if (node.nodeType === 3) {
        text += node.nodeValue;
      }
    }
    return text;
  }

  private expandCells(row: Element): Element[] {
    let cells = Array.from(row.children).filter(
      (el) => el.localName === "tc",
    ) as Element[];
    if (!cells.length) {
      cells = Array.from(
        row.getElementsByTagNameNS(this.ctx.ns.hp, "tc"),
      ).filter((el) => el.parentNode === row) as Element[];
    }
    return cells;
  }

  private convertTable(tblElem: Element): string {
    let allRows = Array.from(tblElem.children).filter(
      (el) => el.localName === "tr",
    ) as Element[];
    if (!allRows.length) {
      allRows = Array.from(
        tblElem.getElementsByTagNameNS(this.ctx.ns.hp, "tr"),
      ).filter((el) => el.parentNode === tblElem) as Element[];
    }

    if (!allRows.length) return "";

    const firstRowCells = this.expandCells(allRows[0]);

    let colCount = 0;
    for (const tc of firstRowCells) {
      const cs = tc.getElementsByTagNameNS(this.ctx.ns.hp, "cellSpan")[0];
      colCount += cs ? parseInt(cs.getAttribute("colSpan") || "1", 10) : 1;
    }

    const rowspanOccupied = new Set<string>();

    // 1. 표 전체 스타일 (테두리 겹침 방지 및 표 전체 배경색 추출)
    const tableStyles: string[] = ["border-collapse: collapse"];
    const tblPr = tblElem.getElementsByTagNameNS(this.ctx.ns.hp, "tblPr")[0];
    if (tblPr) {
      const tblBorderFillId = tblPr.getAttribute("borderFillIDRef");
      if (tblBorderFillId) {
        const bf = this.ctx.borderFills.get(tblBorderFillId);
        if (bf && bf.backgroundColor) {
          tableStyles.push(`background-color: ${bf.backgroundColor}`);
        }
      }
    }

    let html = `<table style="${tableStyles.join("; ")}"><tbody>`;

    for (let r = 0; r < allRows.length; r++) {
      const row = allRows[r];
      const cells = this.expandCells(row);

      html += "<tr>";

      let colCursor = 0;
      let cellIdx = 0;

      while (colCursor < colCount) {
        if (rowspanOccupied.has(`${r},${colCursor}`)) {
          colCursor++;
          continue;
        }

        if (cellIdx < cells.length) {
          const tc = cells[cellIdx++];

          const cs = tc.getElementsByTagNameNS(this.ctx.ns.hp, "cellSpan")[0];

          const colSpan = cs
            ? parseInt(cs.getAttribute("colSpan") || "1", 10)
            : 1;
          const rowSpan = cs
            ? parseInt(cs.getAttribute("rowSpan") || "1", 10)
            : 1;

          const attrs =
            (colSpan > 1 ? ` colspan="${colSpan}"` : "") +
            (rowSpan > 1 ? ` rowspan="${rowSpan}"` : "");

          const content = this.extractCellText(tc);
          const sizeStyle = this.getCellSizeStyle(tc);

          // 2. 셀 개별 배경색 및 테두리 스타일 추출
          const borderStyle = this.getCellBorderStyle(tc);
          const style = [sizeStyle, borderStyle].filter(Boolean).join("; ");

          html += `<td${attrs}${style ? ` style="${style}"` : ""}>${content}</td>`;

          if (rowSpan > 1) {
            for (let rs = 1; rs < rowSpan; rs++) {
              for (let cs2 = 0; cs2 < colSpan; cs2++) {
                rowspanOccupied.add(`${r + rs},${colCursor + cs2}`);
              }
            }
          }

          colCursor += colSpan;
        } else {
          html += "<td></td>";
          colCursor++;
        }
      }

      html += "</tr>";
    }

    html += "</tbody></table>";
    return html;
  }
  private getCellSizeStyle(tc: Element): string {
    const size = tc.getElementsByTagNameNS(this.ctx.ns.hp, "cellSz")[0];
    if (!size) return "";

    const w = size.getAttribute("width");
    const h = size.getAttribute("height");

    const styles: string[] = [];

    if (w) {
      const px = Math.round((parseInt(w, 10) * 96) / 7200);
      styles.push(`width:${px}px`);
    }

    if (h) {
      const px = Math.round((parseInt(h, 10) * 96) / 7200);
      styles.push(`height:${px}px`);
    }

    return styles.join(";");
  }

  private getCellBorderStyle(tc: Element): string {
    let id = tc.getAttribute("borderFillIDRef");
    if (!id) {
      const cellPr = tc.getElementsByTagNameNS(this.ctx.ns.hp, "cellPr")[0];
      if (cellPr) {
        id = cellPr.getAttribute("borderFillIDRef");
      }
    }

    if (!id) return "";

    const bf = this.ctx.borderFills.get(id);
    if (!bf) return "";

    const styles: string[] = [];

    const add = (side: string, line?: { width: number; color: string }) => {
      if (!line || line.width === 0) {
        styles.push(`border-${side}: none`);
        return;
      }
      const width = Math.max(1, line.width);
      styles.push(`border-${side}: ${width}px solid ${line.color}`);
    };

    add("left", bf.left);
    add("right", bf.right);
    add("top", bf.top);
    add("bottom", bf.bottom);

    if (bf.backgroundColor) {
      styles.push(`background-color: ${bf.backgroundColor}`);
    }

    return styles.join("; ");
  }

  private extractCellText(cellElem: Element): string {
    const innerTbl = cellElem.getElementsByTagNameNS(this.ctx.ns.hp, "tbl")[0];
    if (innerTbl) {
      return this.convertTable(innerTbl);
    }

    let container: Element | null =
      cellElem.getElementsByTagNameNS(this.ctx.ns.hp, "subList")[0] ||
      cellElem.getElementsByTagNameNS(this.ctx.ns.hp, "list")[0] ||
      null;

    if (!container) container = cellElem;

    const lines: string[] = [];

    const paragraphs = Array.from(
      container.getElementsByTagNameNS(this.ctx.ns.hp, "p"),
    );

    for (const p of paragraphs) {
      let line = "";

      const runs = Array.from(p.getElementsByTagNameNS(this.ctx.ns.hp, "run"));

      for (const run of runs) {
        const cp = this.ctx.charProperties.get(
          run.getAttribute("charPrIDRef") || "0",
        );

        let text = "";
        const tElems = run.getElementsByTagNameNS(this.ctx.ns.hp, "t");

        for (const t of Array.from(tElems)) {
          text += this.getTextContent(t);
        }

        if (text && cp) {
          text = this.applyInlineFormatting(text, cp);
        }

        line += text;
      }

      if (line.trim()) lines.push(line.trim());
    }

    return lines.join("<br>");
  }

  //md에서는 미지원 예정
  private convertImage(picElem: Element): string {
    const imgElem = picElem.getElementsByTagNameNS(this.ctx.ns.hc, "img")[0];
    if (!imgElem) return "";

    const binaryItemIDRef = imgElem.getAttribute("binaryItemIDRef");
    if (!binaryItemIDRef) return "![image]()";

    const imgInfo = this.ctx.images.get(binaryItemIDRef);
    if (!imgInfo) return "![image]()";

    const mimeType =
      imgInfo.ext === "png"
        ? "image/png"
        : imgInfo.ext === "jpg" || imgInfo.ext === "jpeg"
          ? "image/jpeg"
          : imgInfo.ext === "gif"
            ? "image/gif"
            : imgInfo.ext === "bmp"
              ? "image/bmp"
              : `image/${imgInfo.ext}`;

    const base64 = this.uint8ArrayToBase64(imgInfo.data);
    return `![image](data:${mimeType};base64,${base64})`;
  }

  private uint8ArrayToBase64(data: Uint8Array): string {
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < data.length; i += chunkSize) {
      binary += String.fromCharCode(...data.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  private convertRect(rectElem: Element): string {
    const drawText = rectElem.getElementsByTagNameNS(
      this.ctx.ns.hp,
      "drawText",
    )[0];
    if (!drawText) return "";
    const subList = drawText.getElementsByTagNameNS(
      this.ctx.ns.hp,
      "subList",
    )[0];
    if (!subList) return "";

    let text = "";
    const paragraphs = Array.from(
      subList.getElementsByTagNameNS(this.ctx.ns.hp, "p"),
    ).filter((p) => p.parentNode === subList);

    for (const p of paragraphs) {
      const runs = Array.from(
        p.getElementsByTagNameNS(this.ctx.ns.hp, "run"),
      ).filter((run) => run.parentNode === p);
      for (const run of runs) {
        const tElems = run.getElementsByTagNameNS(this.ctx.ns.hp, "t");
        for (const tElem of Array.from(tElems)) {
          text += this.getTextContent(tElem);
        }
      }
    }

    return text;
  }
}
