import { decode } from "base64-arraybuffer";
import { readFile } from "fs";
import * as http from "http";
import * as https from "http";
import { fileURLToPath } from "url";
import { Fetcher } from "./Fetcher";

export class NodeFetcher implements Fetcher {
  public readBlob = async (_url: string): Promise<ArrayBuffer> => {
    throw new Error("Not Implemented");
  };

  public readData = async (url: string): Promise<ArrayBuffer> => {
    const base64 = url.substring(url.indexOf(",") + 1);
    return decode(base64);
  };

  public readFile = async (url: string): Promise<ArrayBuffer> => {
    const path = fileURLToPath(url);
    return new Promise<ArrayBuffer>((resolve, reject) => {
      readFile(path, (err, data) => {
        if (err) {
          reject(err);
          return;
        }
        const buffer = data.buffer.slice(
          data.byteOffset,
          data.byteOffset + data.byteLength
        );
        resolve(buffer);
      });
    });
  };

  public readHttp = async (url: string): Promise<ArrayBuffer> => {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      http.get(url, (res) => {
        const chunks: Buffer[] = [];
        res.on("error", (err) => {
          reject(err);
        });
        res.on("data", (chunk) => {
          chunks.push(chunk);
        });
        res.on("end", function () {
          const data = Buffer.concat(chunks);
          const buffer = data.buffer.slice(
            data.byteOffset,
            data.byteOffset + data.byteLength
          );
          resolve(buffer);
        });
      });
    });
  };

  public readHttps = async (url: string): Promise<ArrayBuffer> => {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      https.get(url, (res) => {
        const chunks: Buffer[] = [];
        res.on("error", (err) => {
          reject(err);
        });
        res.on("data", (chunk) => {
          chunks.push(chunk);
        });
        res.on("end", function () {
          const data = Buffer.concat(chunks);
          const buffer = data.buffer.slice(
            data.byteOffset,
            data.byteOffset + data.byteLength
          );
          resolve(buffer);
        });
      });
    });
  };
}
