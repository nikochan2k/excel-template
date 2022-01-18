import { Column, Row, Workbook } from "exceljs";
import { template } from "lodash";
import { BinaryData, Converter, dataUrlToBase64 } from "univ-conv";

const EXPR_REGEXP = /<%[^%]+%>/;
const URL_REGEXP = /^(https?|blob|data|file):/;
const IMAGE_EXTENSIONS = /^(jpg|jpeg|png|gif)$/i;

interface CellIndex {
  row: number;
  col: number;
}
class Target {
  tl: CellIndex;
  br: CellIndex;
  widthMap: Map<number, number>;
  heightMap: Map<number, number>;

  get width() {
    return Array.from(this.widthMap.values()).reduce(
      (prev, curr) => prev + curr
    );
  }

  get height() {
    return Array.from(this.heightMap.values()).reduce(
      (prev, curr) => prev + curr
    );
  }

  constructor(row: number, col: number, public text: string) {
    this.tl = { row, col };
    this.br = { row, col };
    this.widthMap = new Map<number, number>();
    this.heightMap = new Map<number, number>();
  }
}

type TargetMap = { [address: string]: Target };
type SheetMap = { [name: string]: TargetMap };

interface ExcelTemplateOptions {
  forceEmbed?: boolean;
  debug?: boolean;
}

const converter = new Converter();

export class ExcelTemplator {
  public static readFile: (path: string) => Promise<Buffer>;

  private options: ExcelTemplateOptions;
  private workbook?: Workbook;

  constructor(
    public xlsx: string | BinaryData,
    options?: ExcelTemplateOptions
  ) {
    if (!options) options = {};
    if (options.debug == null) options.debug = false;
    if (options.forceEmbed == null) options.forceEmbed = false;
    this.options = options;
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
          const text = cell.text;
          let target = targetMap[address];
          if (target) {
            target.br = { row: r, col: c };
          } else {
            if (!cell.isMerged && !EXPR_REGEXP.test(text)) {
              continue;
            }
            target = new Target(r, c, text);
            targetMap[address] = target;
          }
          target.widthMap.set(c, widthMap.get(c) ?? 8.38);
          target.heightMap.set(r, row.height);
        }
      }
    }

    return sheetMap;
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
        obj.text = target.text;
        console.log(obj);
        */

        let text: string;
        try {
          const executor = template(target.text);
          text = executor(data);
        } catch {
          text = this.options.debug ? target.text : "";
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
                ws.addImage(imageId, {
                  tl: { row: target.tl.row - 1, col: target.tl.col - 1 } as any,
                  br: target.br as any,
                });
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
