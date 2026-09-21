declare module 'cfb' {
  interface CFBEntry {
    name: string;
    type: number;
    content: Uint8Array | null;
    size: number;
    L: number;
    R: number;
    C: number;
    startPos: number;
    clsid: string;
    state: number;
    storage: boolean;
  }

  interface CFBResult {
    FileIndex: CFBEntry[];
    FullPaths: string[];
  }

  interface ReadOptions {
    type: 'buffer' | 'file' | 'base64' | 'binary';
  }

  export function read(data: Buffer | Uint8Array | string, options?: ReadOptions): CFBResult;
}
