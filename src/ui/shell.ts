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

// ---------------------------------------------------------------- windows (View ▸ Switch Windows)

export interface WindowInfo { id: string; title: string; current: boolean }

// Browser: tabs of the same browser find each other over a BroadcastChannel.
const tabId = Math.random().toString(36).slice(2, 10);
let channel: BroadcastChannel | null = null;

function browserChannel(): BroadcastChannel | null {
  if (isDesktop() || typeof BroadcastChannel === 'undefined') return null;
  if (!channel) {
    channel = new BroadcastChannel('oword-windows');
    channel.onmessage = (e) => {
      const m = e.data as { type: string; id?: string; from?: string };
      if (m.type === 'who') channel!.postMessage({ type: 'here', id: tabId, title: document.title, to: m.from });
      if (m.type === 'focus' && m.id === tabId) window.focus();
    };
  }
  return channel;
}
browserChannel();

export async function listWindows(): Promise<WindowInfo[]> {
  if (isDesktop()) {
    const { getCurrentWindow } = await import('@tauri-apps/api/window');
    const me = getCurrentWindow().label;
    const list = await invoke<{ label: string; title: string }[]>('list_windows');
    return list.map((w) => ({ id: w.label, title: w.title || 'Oword', current: w.label === me }));
  }
  const ch = browserChannel();
  const out: WindowInfo[] = [{ id: tabId, title: document.title, current: true }];
  if (!ch) return out;
  const listen = (e: MessageEvent) => {
    const m = e.data as { type: string; id: string; title: string; to: string };
    if (m.type === 'here' && m.to === tabId && !out.some((w) => w.id === m.id)) out.push({ id: m.id, title: m.title, current: false });
  };
  ch.addEventListener('message', listen);
  ch.postMessage({ type: 'who', from: tabId });
  await new Promise((r) => setTimeout(r, 150));
  ch.removeEventListener('message', listen);
  return out;
}

export async function focusWindow(id: string) {
  if (isDesktop()) { await invoke('focus_window', { label: id }); return; }
  browserChannel()?.postMessage({ type: 'focus', id });
}

/** Open another editor window (desktop) or tab (browser), optionally with a document path. */
export async function openNewWindow(path?: string) {
  if (isDesktop()) { await invoke('new_window', { path: path ?? null }); return; }
  window.open(location.pathname, '_blank');
}

/** Desktop windows opened with New Window receive the document to open (or null for a blank one). */
export function windowOpenRequest(): string | null | undefined {
  return (window as unknown as { __OWORD_OPEN__?: string | null }).__OWORD_OPEN__;
}

export async function readPath(path: string): Promise<OpenedFile> {
  const f = await invoke<NativeFile>('read_document', { path });
  return { name: f.name, path: f.path, bytes: new Uint8Array(f.bytes) };
}
