import { ExcelTemplator } from "./ExcelTemplator";

ExcelTemplator.readFile = (_path: string): Promise<Buffer> => {
  throw new Error("file protocol is not supported");
};

export * from "./ExcelTemplator";
