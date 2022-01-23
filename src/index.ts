if (!globalThis.Buffer) {
  (globalThis as any).Buffer = require("buffer/").Buffer;
}

export * from "./Fetcher";
export * from "./ExcelTemplator";
