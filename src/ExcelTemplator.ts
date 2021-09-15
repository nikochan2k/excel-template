import { Column, Row, Workbook } from "exceljs";
import { template } from "lodash";
import { BinarySource, Converter, dataUrlToBase64 } from "univ-conv";

const EXPR_REGEXP = /<%[^%]+%>/;
const URL_REGEXP = /^(https?|blob|data|file):/;
const IMAGE_EXTENSIONS = /^(jpg|jpeg|png|gif)$/i;

interface CellIndex {
  row: number;
  col: number;
}
interface Target {
  tl: CellIndex;
  br: CellIndex;
  text?: string;
}

interface ExcelTemplateOptions {
  forceEmbed?: boolean;
  debug?: boolean;
}

const converter = new Converter();

export class ExcelTemplator {
  public static readFile: (path: string) => Promise<Buffer>;

  private options: ExcelTemplateOptions;

  constructor(
    public xlsx: string | BinarySource,
    options?: ExcelTemplateOptions
  ) {
    if (!options) options = {};
    if (options.debug == null) options.debug = false;
    if (options.forceEmbed == null) options.forceEmbed = false;
    this.options = options;
  }

  public async generate(data: any): Promise<ArrayBuffer> {
    let buffer: ArrayBuffer | undefined;
    if (typeof this.xlsx === "string") {
      const urlLike = this.xlsx;
      const url = new URL(urlLike);
      buffer = await this.fetchURL(url);
    } else {
      buffer = await converter.toArrayBuffer(this.xlsx);
    }

    const workbook = new Workbook();
    await workbook.xlsx.load(buffer);
    outer: for (const ws of workbook.worksheets) {
      const targetMap: { [key: string]: Target } = {};
      const lastRow = ws.lastRow as Row;
      for (
        let row = 1, endRow = lastRow?.number as number;
        row <= endRow;
        row++
      ) {
        const lastColumn = ws.lastColumn as Column;
        for (
          let col = 1, endColumn = lastColumn?.number as number;
          col <= endColumn;
          col++
        ) {
          const cell = ws.getCell(row, col);
          const key = cell.master.address;
          const text = cell.text;
          let target = targetMap[key];
          if (target) {
            target.br = { row, col };
          } else {
            if (!cell.isMerged && !EXPR_REGEXP.test(text)) {
              continue;
            }
            target = { tl: { row, col }, br: { row, col } };
            targetMap[key] = target;
          }
          if (!target.text) {
            try {
              const executor = template(text);
              target.text = executor(data);
            } catch {
              target.text = this.options.debug ? text : "";
            }
          }
        }
      }

      for (const [address, target] of Object.entries(targetMap)) {
        const text = target.text || "";
        try {
          if (URL_REGEXP.test(text)) {
            const url = new URL(text);
            if (this.options.forceEmbed || url.hash === "#embed") {
              const lastIndex = url.pathname.lastIndexOf(".");
              let extension = url.pathname.substr(lastIndex + 1).toLowerCase();
              if (extension === "jpg") extension = "jpeg";
              if (IMAGE_EXTENSIONS.test(extension)) {
                const buffer = await this.fetchURL(url);
                const imageId = workbook.addImage({
                  buffer,
                  extension: extension as any,
                });

                ws.addImage(imageId, {
                  tl: { row: target.tl.row - 1, col: target.tl.col - 1 } as any,
                  br: target.br as any,
                });
              }
              continue outer;
            }
          }
        } catch (e) {
          console.warn(e);
        }

        const cell = ws.getCell(address);
        const value: any = cell.value;
        if (value.font) {
          value.text = text;
        } else {
          cell.value = text;
        }
      }
    }

    return workbook.xlsx.writeBuffer();
  }

  private async fetchURL(url: URL): Promise<ArrayBuffer | Buffer> {
    const proto = url.protocol;
    if (proto === "http:" || proto === "https:" || proto === "blob:") {
      const fetched = await fetch(url.href);
      return fetched.arrayBuffer();
    } else if (proto === "file:") {
      return ExcelTemplator.readFile(url.pathname);
    } else if (proto === "data:") {
      const base64 = dataUrlToBase64(url.href);
      return converter.toArrayBuffer({
        encoding: "Base64",
        value: base64,
      });
    }
    throw new Error("Unknown protocol: " + url.protocol);
  }
}