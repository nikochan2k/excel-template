export interface Fetcher {
  readBlob: (url: string) => Promise<ArrayBuffer>;
  readData: (url: string) => Promise<ArrayBuffer>;
  readFile: (url: string) => Promise<ArrayBuffer>;
  readHttp: (url: string) => Promise<ArrayBuffer>;
  readHttps: (url: string) => Promise<ArrayBuffer>;
}
