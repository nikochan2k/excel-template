import { ExcelTemplate } from "./ExcelTemplate";
import { readFile } from "fs";
import "isomorphic-fetch";

ExcelTemplate.readFile = (path: string): Promise<Buffer> => {
  return new Promise<Buffer>((resolve, reject) => {
    readFile(path, (err, buffer) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(buffer);
    });
  });
};

export * from "./ExcelTemplate";
