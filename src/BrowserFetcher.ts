import { decode } from "base64-arraybuffer";
import { Fetcher } from "./Fetcher";

export class BrowserFetcher implements Fetcher {
  public readBlob = async (url: string): Promise<Buffer> => {
    const fetched = await fetch(url);
    const arrayBuffer = await fetched.arrayBuffer();
    return Buffer.from(arrayBuffer);
  };

  public readData = async (url: string): Promise<Buffer> => {
    const base64 = url.substring(url.indexOf(",") + 1);
    const arrayBuffer = decode(base64);
    return Buffer.from(arrayBuffer);
  };

  public readFile = async (_url: string): Promise<Buffer> => {
    throw new Error("Not Implemented");
  };

  public readHttp = async (url: string): Promise<Buffer> => {
    const fetched = await fetch(url);
    const arrayBuffer = await fetched.arrayBuffer();
    return Buffer.from(arrayBuffer);
  };

  public readHttps = async (url: string): Promise<Buffer> => {
    const fetched = await fetch(url);
    const arrayBuffer = await fetched.arrayBuffer();
    return Buffer.from(arrayBuffer);
  };
}
