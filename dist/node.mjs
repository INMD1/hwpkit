// src/node/index.ts
import { mkdtemp, readFile, writeFile, mkdir, rm } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { pathToFileURL } from "url";
import { spawn } from "child_process";
function createLibreOfficeDocConverter(options = {}) {
  const executable = options.executable ?? "libreoffice";
  const timeoutMs = options.timeoutMs ?? 6e4;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error("timeoutMs must be positive");
  return async (data) => {
    const directory = await mkdtemp(join(tmpdir(), "hwpkit-doc-"));
    try {
      const input = join(directory, "input.doc");
      const output = join(directory, "output");
      await writeFile(input, data);
      await mkdir(output);
      await new Promise((resolve, reject) => {
        const child = spawn(executable, [
          `-env:UserInstallation=${pathToFileURL(join(directory, "profile")).href}`,
          "--headless",
          "--convert-to",
          "docx:Office Open XML Text",
          "--outdir",
          output,
          input
        ], { stdio: ["ignore", "pipe", "pipe"], detached: process.platform !== "win32" });
        let detail = "";
        let timedOut = false;
        const capture = (chunk) => {
          detail = (detail + chunk.toString()).slice(-4e3);
        };
        child.stdout.on("data", capture);
        child.stderr.on("data", capture);
        const timer = setTimeout(() => {
          timedOut = true;
          try {
            if (process.platform !== "win32" && child.pid) process.kill(-child.pid, "SIGKILL");
            else child.kill("SIGKILL");
          } catch {
            child.kill("SIGKILL");
          }
        }, timeoutMs);
        child.on("error", (error) => {
          clearTimeout(timer);
          reject(new Error(`LibreOffice \uC2E4\uD589 \uC2E4\uD328 (${executable}): ${error.message}`));
        });
        child.on("close", (code) => {
          clearTimeout(timer);
          if (timedOut) reject(new Error(`DOC \uBCC0\uD658 \uC2DC\uAC04 \uCD08\uACFC (${timeoutMs}ms)`));
          else if (code !== 0) reject(new Error(`LibreOffice DOC \uBCC0\uD658 \uC2E4\uD328 (${code}): ${detail.trim()}`));
          else resolve();
        });
      });
      let result;
      try {
        result = new Uint8Array(await readFile(join(output, "input.docx")));
      } catch {
        throw new Error("LibreOffice\uAC00 DOCX\uB97C \uC0DD\uC131\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4. DOC \uD30C\uC77C\uC744 \uD655\uC778\uD558\uC138\uC694.");
      }
      if (result[0] !== 80 || result[1] !== 75) throw new Error("LibreOffice \uCD9C\uB825\uC774 DOCX\uAC00 \uC544\uB2D9\uB2C8\uB2E4.");
      return result;
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  };
}
export {
  createLibreOfficeDocConverter
};
//# sourceMappingURL=node.mjs.map