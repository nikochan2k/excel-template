import type { Alignment, Column, Row, Workbook, WorkbookModel } from "exceljs";
import template from "lodash/template";
import { Fetcher } from "./Fetcher";
const exceljs = require("./exceljs");

const EXPR_REGEXP = /<%=([^%]+)%>/;
const URL_REGEXP = /^(https?|blob|data|file):/;
const IMAGE_EXTENSIONS = /^(jpg|jpeg|png|gif)$/i;

interface Size {
  width: number;
  height: number;
  fill: boolean;
}

interface CellIndex {
  col: number;
  row: number;
}

type HorizontalAlign =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

type VerticalAlign = "justify" | "distributed" | "top" | "middle" | "bottom";

export class Target {
  public br: CellIndex;
  public heightMap: Map<number, number>;
  public tl: CellIndex;
  public widthMap: Map<number, number>;
  public horizontalAlign?: HorizontalAlign;
  public verticalAlign?: VerticalAlign;
  public ext?: { width: number; height: number };

  // default
  constructor(
    row: number,
    col: number,
    public expr: string,
    align?: Partial<Alignment>
  ) {
    this.tl = { row, col };
    this.br = { row, col };
    this.widthMap = new Map<number, number>();
    this.heightMap = new Map<number, number>();
    this.horizontalAlign = align?.horizontal;
    this.verticalAlign = align?.vertical;
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

export type TargetMap = { [address: string]: Target };
export type SheetMap = { [name: string]: TargetMap };

interface ResizeOption {
  fetch: (url: string) => Promise<ArrayBuffer>;
  target: Target;
  url: string;
}
interface ExcelTemplateOptions {
  debug?: boolean;
  forceEmbed?: boolean;
  getSize?: (url: string, target: Target) => Promise<Size>;
  resize?: (options: ResizeOption) => Promise<ArrayBuffer>;
}

export function serialize(workbook: Workbook) {
  const model = workbook.model;
  delete (model as any)._sheet;
  for (const sheet of model.worksheets) {
    const s = sheet as any;
    s.mergeCells = s.merges;
    delete s.merges;
  }
  const json = JSON.stringify(model, (key, value) => {
    if (key === "media" || key === "_media") {
      return [];
    }
    return value;
  });
  return json;
}

export function deserialize(json: string) {
  const model: WorkbookModel = JSON.parse(json, (key, value) => {
    if (key === "media" || key === "_media") {
      return [];
    }
    return /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/.test(value)
      ? new Date(value)
      : value;
  });
  const workbook: Workbook = new exceljs.Workbook();
  workbook.model = model;
  return workbook;
}

export function fit(
  target: Target,
  width: number,
  height: number,
  fill: boolean
) {
  if (!width || !height) {
    return;
  }

  const wRatio = target.width / width;
  const hRatio = target.height / height;
  let ratio = Math.min(wRatio, hRatio);
  if (1 < ratio) {
    ratio = 1;
  }
  width = width * ratio;
  height = height * ratio;
  if (fill) {
    target.ext = { width, height };
    return;
  }

  if (target.horizontalAlign) {
    let xOffset = 0;
    switch (target.horizontalAlign) {
      case "center":
      case "centerContinuous":
      case "distributed":
      case "fill":
      case "justify":
        xOffset = (target.width - width) / 2;
        break;
      case "right":
        xOffset = target.width - width;
        break;
    }
    if (0 < xOffset) {
      let currentWidth = 0;
      for (let c = target.tl.col; c <= target.br.col; c++) {
        const w = target.widthMap.get(c) ?? 0;
        if (currentWidth <= xOffset && xOffset < currentWidth + w) {
          target.tl.col = c + (xOffset - currentWidth) / w;
          break;
        }
        currentWidth += w;
      }
    }
  }

  if (target.verticalAlign) {
    let yOffset = 0;
    switch (target.verticalAlign) {
      case "middle":
      case "distributed":
      case "justify":
        yOffset = (target.height - height) / 2;
        break;
      case "bottom":
        yOffset = target.height - height;
        break;
    }
    if (0 < yOffset) {
      let currentHeight = 0;
      for (let r = target.tl.row; r <= target.br.row; r++) {
        const h = target.widthMap.get(r) ?? 0;
        if (currentHeight <= yOffset && yOffset < currentHeight + h) {
          target.tl.row = r + (yOffset - currentHeight) / h;
          break;
        }
        currentHeight += h;
      }
    }
  }

  target.ext = { width, height };
}

export class ExcelTemplator {
  private options: ExcelTemplateOptions;
  private workbook?: Workbook;

  public static BASE_WIDTH = 7.9;

  constructor(
    public xlsx: string | ArrayBuffer | Blob | Workbook,
    private fetcher: Fetcher,
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
            const url = text;
            if (this.options.forceEmbed || url.endsWith("#embed")) {
              const res = /[.\/](jpg|jpeg|png|gif)/i.exec(url);
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
                let fill = false;
                if (this.options.getSize) {
                  const size = await this.options.getSize(url, target);
                  fill = size.fill;
                  fit(target, size.width, size.height, fill);
                }
                if (target.ext) {
                  let buffer: ArrayBuffer;
                  if (this.options.resize) {
                    buffer = await this.options.resize({
                      fetch: this.fetch,
                      url,
                      target,
                    });
                  } else {
                    buffer = await this.fetch(url);
                  }
                  const imageId = workbook.addImage({ buffer, extension });
                  if (fill) {
                    ws.addImage(imageId, {
                      tl: {
                        row: target.tl.row - 1,
                        col: target.tl.col - 1,
                      } as any,
                      br: target.br as any,
                    });
                  } else {
                    ws.addImage(imageId, {
                      tl: {
                        row: target.tl.row - 1,
                        col: target.tl.col - 1,
                      } as any,
                      ext: target.ext,
                    });
                  }
                } else {
                  const buffer = await this.fetch(url);
                  const imageId = workbook.addImage({ buffer, extension });
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
        if (!value) {
          continue;
        }
        if (value.font) {
          value.text = text;
        } else {
          cell.value = text;
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    delete this.workbook;
    return buffer as any;
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
            target = new Target(r, c, text, cell.style.alignment);
            targetMap[address] = target;
          }
          target.widthMap.set(c, this.width2px(widthMap.get(c) ?? 8.38));
          target.heightMap.set(r, (row.height * 96) / 72);
        }
      }
    }

    return sheetMap;
  }

  private async fetch(url: string): Promise<ArrayBuffer> {
    if (url.startsWith("https:")) {
      return this.fetcher.readHttps(url);
    } else if (url.startsWith("http:")) {
      return this.fetcher.readHttp(url);
    } else if (url.startsWith("blob:")) {
      return this.fetcher.readBlob(url);
    } else if (url.startsWith("file:")) {
      return this.fetcher.readFile(url);
    } else if (url.startsWith("data:")) {
      return this.fetcher.readData(url);
    }
    throw new Error("Unknown protocol: " + url.substring(0, 10));
  }

  private async load() {
    if (this.workbook) {
      return this.workbook;
    }

    let buffer: ArrayBuffer;
    if (typeof this.xlsx === "string") {
      const url = this.xlsx;
      buffer = await this.fetch(url);
    } else if (typeof (this.xlsx as ArrayBuffer).byteLength === "number") {
      buffer = this.xlsx as ArrayBuffer;
    } else if (typeof (this.xlsx as Blob).size === "number") {
      const blob = this.xlsx as Blob;
      buffer = await blob.arrayBuffer();
    } else {
      this.workbook = this.xlsx as Workbook;
      return this.workbook;
    }
    this.workbook = new exceljs.Workbook() as Workbook;
    await this.workbook.xlsx.load(buffer);
    return this.workbook;
  }

  private width2px(width: number) {
    const baseWidth = ExcelTemplator.BASE_WIDTH;
    const pad = Math.round((baseWidth + 1) / 4) * 2 + 1;
    const zPad = baseWidth + pad;
    return width < 1 ? width * zPad : width * baseWidth + pad;
  }
}
