// Platform shell: file open/save.
// - Desktop (Tauri, like HOP): native dialogs + file I/O through Rust commands.
// - Browser: File System Access API when available, else <input type=file> and downloads.
import { invoke } from '@tauri-apps/api/core';

export interface OpenedFile { name: string; bytes: Uint8Array; handle?: FileSystemFileHandle; path?: string }

export const isDesktop = () => '__TAURI_INTERNALS__' in window;

interface NativeFile { name: string; path: string; bytes: number[] }

export async function launchFiles(): Promise<OpenedFile[]> {
  if (!isDesktop()) return [];
  const paths = await invoke<string[]>('launch_paths');
  const out: OpenedFile[] = [];
  for (const p of paths) {
    const f = await invoke<NativeFile>('read_document', { path: p });
    out.push({ name: f.name, path: f.path, bytes: new Uint8Array(f.bytes) });
  }
  return out;
}

const DOCX_TYPES = [{
  description: 'Word 문서',
  accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] },
}];

type FsWindow = Window & {
  showOpenFilePicker?: (o: unknown) => Promise<FileSystemFileHandle[]>;
  showSaveFilePicker?: (o: unknown) => Promise<FileSystemFileHandle>;
};

export async function pickOpenFile(input: HTMLInputElement): Promise<OpenedFile | null> {
  if (isDesktop()) {
    const f = await invoke<NativeFile | null>('open_document');
    return f ? { name: f.name, path: f.path, bytes: new Uint8Array(f.bytes) } : null;
  }
  const w = window as FsWindow;
  if (w.showOpenFilePicker) {
    try {
      const [handle] = await w.showOpenFilePicker({ types: DOCX_TYPES, multiple: false });
      const f = await handle.getFile();
      return { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()), handle };
    } catch (e) {
      if ((e as DOMException).name === 'AbortError') return null;
    }
  }
  return new Promise((resolve) => {
    input.value = '';
    input.onchange = async () => {
      const f = input.files?.[0];
      resolve(f ? { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) } : null);
    };
    input.click();
  });
}

export async function saveFile(blob: Blob, name: string, handle: FileSystemFileHandle | undefined, saveAs: boolean, path?: string): Promise<{ name: string; handle?: FileSystemFileHandle; path?: string } | null> {
  if (isDesktop()) {
    const bytes = Array.from(new Uint8Array(await blob.arrayBuffer()));
    const saved = await invoke<string | null>('save_document', { path: saveAs ? null : path ?? null, suggested: name, bytes });
    return saved ? { name: saved.split(/[\\/]/).pop()!, path: saved } : null;
  }
  const w = window as FsWindow;
  if (handle && !saveAs) {
    const ws = await (handle as FileSystemFileHandle & { createWritable(): Promise<FileSystemWritableFileStream> }).createWritable();
    await ws.write(blob);
    await ws.close();
    return { name, handle };
  }
  if (w.showSaveFilePicker) {
    try {
      const h = await w.showSaveFilePicker({ suggestedName: name, types: DOCX_TYPES });
      const ws = await (h as FileSystemFileHandle & { createWritable(): Promise<FileSystemWritableFileStream> }).createWritable();
      await ws.write(blob);
      await ws.close();
      return { name: h.name, handle: h };
    } catch (e) {
      if ((e as DOMException).name === 'AbortError') return null;
    }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  return { name };
}

export function pickImage(input: HTMLInputElement): Promise<File | null> {
  return new Promise((resolve) => {
    input.value = '';
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.click();
  });
}
