import { Column, Row, Workbook } from "exceljs";
import { template } from "lodash";
import { BinaryData, Converter, dataUrlToBase64 } from "univ-conv";

const EXPR_REGEXP = /<%=([^%]+)%>/;
const URL_REGEXP = /^(https?|blob|data|file):/;
const IMAGE_EXTENSIONS = /^(jpg|jpeg|png|gif)$/i;

interface CellIndex {
  col: number;
  row: number;
}
class Target {
  public br: CellIndex;
  ext?: { width: number; height: number };
  public heightMap: Map<number, number>;
  public tl: CellIndex;
  public widthMap: Map<number, number>;

  // default
  constructor(row: number, col: number, public expr: string) {
    this.tl = { row, col };
    this.br = { row, col };
    this.widthMap = new Map<number, number>();
    this.heightMap = new Map<number, number>();
  }

  public get height() {
    return Array.from(this.heightMap.values()).reduce(
      (prev, curr) => prev + curr
    );
  }

  public get val(): string | undefined {
    if (!this.expr) {
      return undefined;
    }

    const arr = EXPR_REGEXP.exec(this.expr);
    if (!arr) {
      return undefined;
    }

    return arr[1]?.trim();
  }

  public get width() {
    return Array.from(this.widthMap.values()).reduce(
      (prev, curr) => prev + curr
    );
  }
}

type TargetMap = { [address: string]: Target };
type SheetMap = { [name: string]: TargetMap };

interface ExcelTemplateOptions {
  debug?: boolean;
  forceEmbed?: boolean;
}

const converter = new Converter();

export function fit(target: Target, width: number, height: number) {
  if (!width || !height) {
    return undefined;
  }

  let ratio = 1;
  if (target.width < width) {
    ratio = target.width / width;
  }
  if (target.height < height) {
    const tmp = target.height / height;
    if (tmp < ratio) {
      ratio = tmp;
    }
  }
  if (ratio < 1) {
    width = width * ratio;
    height = height * ratio;
  }

  return { width, height };
}

export class ExcelTemplator {
  private options: ExcelTemplateOptions;
  private workbook?: Workbook;

  public static PX_0_10PT = 7.2;
  public static readFile: (path: string) => Promise<Buffer>;

  constructor(
    public xlsx: string | BinaryData,
    options?: ExcelTemplateOptions
  ) {
    if (!options) options = {};
    if (options.debug == null) options.debug = false;
    if (options.forceEmbed == null) options.forceEmbed = false;
    this.options = options;
  }

  public async generate(data: any, sheetMap?: SheetMap): Promise<ArrayBuffer> {
    if (!sheetMap) {
      sheetMap = await this.parse();
    }

    const workbook = await this.load();
    for (const ws of workbook.worksheets) {
      const targetMap = sheetMap[ws.name];
      if (!targetMap) {
        continue;
      }

      for (const [address, target] of Object.entries(targetMap)) {
        /*
        const obj: any = {};
        obj.tl = target.tl;
        obj.br = target.br;
        obj.width = target.width;
        obj.height = target.height;
        obj.text = target.expr;
        console.log(obj);
        */

        let text: string;
        try {
          const executor = template(target.expr);
          text = executor(data);
        } catch {
          text = this.options.debug ? target.expr : "";
        }
        const cell = ws.getCell(address);
        try {
          if (URL_REGEXP.test(text)) {
            const url = new URL(text);
            if (this.options.forceEmbed || url.hash === "#embed") {
              const res = /[.\/](jpg|jpeg|png|gif)/i.exec(text);
              let extension: "jpeg" | "png" | "gif";
              if (!res) {
                extension = "png";
              } else {
                let ext = res[1]?.toLowerCase();
                if (ext === "jpg" || ext === "jpeg") {
                  extension = "jpeg";
                } else if (ext === "gif") {
                  extension = "gif";
                } else {
                  extension = "png";
                }
              }
              if (IMAGE_EXTENSIONS.test(extension)) {
                const buffer = await this.fetchURL(url);
                const imageId = workbook.addImage({ buffer, extension });
                if (target.ext) {
                  ws.addImage(imageId, {
                    tl: {
                      row: target.tl.row - 1,
                      col: target.tl.col - 1,
                    } as any,
                    ext: target.ext,
                  });
                } else {
                  ws.addImage(imageId, {
                    tl: {
                      row: target.tl.row - 1,
                      col: target.tl.col - 1,
                    } as any,
                    br: target.br as any,
                  });
                }
              }

              cell.value = "";
              continue;
            }
          }
        } catch (e) {
          console.warn(e);
        }

        const value: any = cell.value;
        if (value.font) {
          value.text = text;
        } else {
          cell.value = text;
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    delete this.workbook;
    return buffer;
  }

  public async parse() {
    const workbook = await this.load();
    const sheetMap: SheetMap = {};
    for (const ws of workbook.worksheets) {
      const targetMap: TargetMap = {};
      sheetMap[ws.name] = targetMap;
      const lastColumn = ws.lastColumn as Column;
      const widthMap = new Map<number, number>();
      for (
        let c = 1, endColumn = lastColumn?.number ?? 1;
        c <= endColumn;
        c++
      ) {
        const col = ws.getColumn(c);
        widthMap.set(c, col.width ?? ws.properties.defaultColWidth ?? 8.38);
      }

      const lastRow = ws.lastRow as Row;
      for (let r = 1, endRow = lastRow?.number ?? 1; r <= endRow; r++) {
        const row = ws.getRow(r);
        for (
          let c = 1, endColumn = lastColumn?.number ?? 1;
          c <= endColumn;
          c++
        ) {
          const cell = ws.getCell(r, c);
          const address = cell.master.address;
          let target = targetMap[address];
          if (target) {
            target.br = { row: r, col: c };
          } else {
            const text = cell.text;
            if (!cell.isMerged && !EXPR_REGEXP.test(text)) {
              continue;
            }
            target = new Target(r, c, text);
            targetMap[address] = target;
          }
          target.widthMap.set(c, this.width2px(widthMap.get(c) ?? 8.38));
          target.heightMap.set(r, (row.height * 96) / 72);
        }
      }
    }

    return sheetMap;
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

  private async load() {
    if (this.workbook) {
      return this.workbook;
    }

    let buffer: ArrayBuffer;
    if (typeof this.xlsx === "string") {
      const urlLike = this.xlsx;
      const url = new URL(urlLike);
      buffer = await this.fetchURL(url);
    } else {
      buffer = await converter.toArrayBuffer(this.xlsx);
    }
    this.workbook = new Workbook();
    await this.workbook.xlsx.load(buffer);
    return this.workbook;
  }

  private width2px(width: number) {
    const zeroWidth = ExcelTemplator.PX_0_10PT;
    const pad = Math.round((zeroWidth + 1) / 4) * 2 + 1;
    const zPad = zeroWidth + pad;
    return width < 1 ? width * zPad : width * zeroWidth + pad;
  }
}
