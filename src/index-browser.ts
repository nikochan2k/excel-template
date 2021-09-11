import { ExcelTemplate } from "./ExcelTemplate";

ExcelTemplate.readFile = (_path: string): Promise<Buffer> => {
  throw new Error("file protocol is not supported");
};

export * from "univ-conv";
export * from "./ExcelTemplate";
