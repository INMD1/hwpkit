"use strict";
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/node/index.ts
var node_exports = {};
__export(node_exports, {
  createLibreOfficeDocConverter: () => createLibreOfficeDocConverter
});
module.exports = __toCommonJS(node_exports);
var import_promises = require("fs/promises");
var import_node_os = require("os");
var import_node_path = require("path");
var import_node_url = require("url");
var import_node_child_process = require("child_process");
function createLibreOfficeDocConverter(options = {}) {
  const executable = options.executable ?? "libreoffice";
  const timeoutMs = options.timeoutMs ?? 6e4;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error("timeoutMs must be positive");
  return async (data) => {
    const directory = await (0, import_promises.mkdtemp)((0, import_node_path.join)((0, import_node_os.tmpdir)(), "hwpkit-doc-"));
    try {
      const input = (0, import_node_path.join)(directory, "input.doc");
      const output = (0, import_node_path.join)(directory, "output");
      await (0, import_promises.writeFile)(input, data);
      await (0, import_promises.mkdir)(output);
      await new Promise((resolve, reject) => {
        const child = (0, import_node_child_process.spawn)(executable, [
          `-env:UserInstallation=${(0, import_node_url.pathToFileURL)((0, import_node_path.join)(directory, "profile")).href}`,
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
        result = new Uint8Array(await (0, import_promises.readFile)((0, import_node_path.join)(output, "input.docx")));
      } catch {
        throw new Error("LibreOffice\uAC00 DOCX\uB97C \uC0DD\uC131\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4. DOC \uD30C\uC77C\uC744 \uD655\uC778\uD558\uC138\uC694.");
      }
      if (result[0] !== 80 || result[1] !== 75) throw new Error("LibreOffice \uCD9C\uB825\uC774 DOCX\uAC00 \uC544\uB2D9\uB2C8\uB2E4.");
      return result;
    } finally {
      await (0, import_promises.rm)(directory, { recursive: true, force: true });
    }
  };
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  createLibreOfficeDocConverter
});
//# sourceMappingURL=node.js.map