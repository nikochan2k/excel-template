import { ExcelTemplator } from "./ExcelTemplator";
import { readFile } from "fs";

ExcelTemplator.readFile = (path: string): Promise<Buffer> => {
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

export * from "./ExcelTemplator";
