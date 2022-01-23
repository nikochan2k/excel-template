import { decode } from "base64-arraybuffer";
import { Fetcher } from "./Fetcher";

export class BrowserFetcher implements Fetcher {
  public readBlob = async (url: string): Promise<ArrayBuffer> => {
    const fetched = await fetch(url);
    return fetched.arrayBuffer();
  };

  public readData = async (url: string): Promise<ArrayBuffer> => {
    const base64 = url.substring(url.indexOf(",") + 1);
    return decode(base64);
  };

  public readFile = async (_url: string): Promise<ArrayBuffer> => {
    throw new Error("Not Implemented");
  };

  public readHttp = async (url: string): Promise<ArrayBuffer> => {
    const fetched = await fetch(url);
    return fetched.arrayBuffer();
  };

  public readHttps = async (url: string): Promise<ArrayBuffer> => {
    const fetched = await fetch(url);
    return fetched.arrayBuffer();
  };
}
