export interface ElectronAPI {
  isElectron?: boolean;
  minimize: () => void;
  maximize: () => void;
  close: () => void;
  openNativeFileDialog: () => Promise<{
    name: string;
    path: string;
    buffer: number[];
  } | null>;
  saveNativeFileDialog: (content: string, defaultName?: string) => Promise<boolean>;
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}
