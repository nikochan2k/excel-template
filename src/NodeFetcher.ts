import { decode } from "base64-arraybuffer";
import { readFile } from "fs";
import * as http from "http";
import * as https from "https";
import { fileURLToPath } from "url";
import { Fetcher } from "./Fetcher";

export class NodeFetcher implements Fetcher {
  public readBlob = async (_url: string): Promise<Buffer> => {
    throw new Error("Not Implemented");
  };

  public readData = async (url: string): Promise<Buffer> => {
    const base64 = url.substring(url.indexOf(",") + 1);
    const arrayBuffer = decode(base64);
    return Buffer.from(arrayBuffer);
  };

  public readFile = async (url: string): Promise<Buffer> => {
    const path = fileURLToPath(url);
    return new Promise<Buffer>((resolve, reject) => {
      readFile(path, (err, data) => {
        if (err) {
          reject(err);
          return;
        }
        resolve(data);
      });
    });
  };

  public readHttp = async (url: string): Promise<Buffer> => {
    return new Promise<Buffer>((resolve, reject) => {
      http.get(url, (res) => {
        const chunks: Buffer[] = [];
        res.on("error", (err) => {
          reject(err);
        });
        res.on("data", (chunk) => {
          chunks.push(chunk);
        });
        res.on("end", () => {
          const data = Buffer.concat(chunks);
          resolve(data);
        });
      });
    });
  };

  public readHttps = async (url: string): Promise<Buffer> => {
    return new Promise<Buffer>((resolve, reject) => {
      https.get(url, (res) => {
        const chunks: Buffer[] = [];
        res.on("error", (err) => {
          reject(err);
        });
        res.on("data", (chunk) => {
          chunks.push(chunk);
        });
        res.on("end", () => {
          const data = Buffer.concat(chunks);
          resolve(data);
        });
      });
    });
  };
}
