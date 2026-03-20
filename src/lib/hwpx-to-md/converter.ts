import JSZip from "jszip";
import { MdConversionContext } from "./context";
import { HwpxMdParser } from "./parser";
import { HwpxMdGenerator } from "./generator";

export class HwpxToMdConverter {
  async convert(hwpxInput: File | ArrayBuffer | Blob): Promise<string> {
    const zip = await JSZip.loadAsync(hwpxInput);
    return await this.convertFromZip(zip);
  }

  private async convertFromZip(zip: JSZip): Promise<string> {
    const ctx = new MdConversionContext();
    const parser = new HwpxMdParser(ctx);
    const generator = new HwpxMdGenerator(ctx);

    // header.xml 파싱 (스타일, 문자 속성, 단락 속성)
    const header = zip.file("Contents/header.xml");
    if (header) {
      parser.parseHeader(await header.async("string"));
    }

    // BinData 이미지 수집
    const binDataFolder = zip.folder("BinData");
    if (binDataFolder) {
      const imageFiles: { name: string; file: JSZip.JSZipObject }[] = [];
      binDataFolder.forEach((relativePath, file) => {
        if (!file.dir) imageFiles.push({ name: relativePath, file });
      });
      for (const { name, file } of imageFiles) {
        const data = await file.async("uint8array");
        const parts = name.split(".");
        const ext = parts.length > 1 ? parts.pop()!.toLowerCase() : "";
        const id = parts.length > 0 ? parts[0].split("/").pop()! : name;
        ctx.images.set(id, { data, ext });
      }
    }

    // section 파일 순서대로 변환 (section0.xml, section1.xml, ...)
    const sectionFiles = Object.keys(zip.files)
      .filter((path) => /Contents\/section\d+\.xml$/i.test(path))
      .sort((a, b) => {
        const aNum = parseInt(a.match(/section(\d+)/i)?.[1] ?? "0");
        const bNum = parseInt(b.match(/section(\d+)/i)?.[1] ?? "0");
        return aNum - bNum;
      });

    const parts: string[] = [];
    let isFirstSection = true;
    for (const sectionPath of sectionFiles) {
      const sectionFile = zip.file(sectionPath);
      if (sectionFile) {
        const xml = await sectionFile.async("string");
        const md = generator.convertSection(xml, isFirstSection);
        if (md) parts.push(md);
        isFirstSection = false;
      }
    }

    return parts.join("\n\n");
  }
}
