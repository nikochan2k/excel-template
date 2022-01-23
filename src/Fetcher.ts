export interface Fetcher {
  readBlob: (url: string) => Promise<Buffer>;
  readData: (url: string) => Promise<Buffer>;
  readFile: (url: string) => Promise<Buffer>;
  readHttp: (url: string) => Promise<Buffer>;
  readHttps: (url: string) => Promise<Buffer>;
}
