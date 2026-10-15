// OPC (Open Packaging Conventions) container: zip parts, relationships, content types.
import JSZip from 'jszip';
import { NS, children, parseXml, serializeXml } from './xml';

export interface Relationship {
  id: string;
  type: string;
  target: string; // absolute part path (no leading slash) for internal targets
  external: boolean;
}

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
};

export const CT = {
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
};

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export function dirname(p: string) {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i);
}

export function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base ? base.split('/') : []).concat(target.split('/'));
  const out: string[] = [];
  for (const s of parts) {
    if (s === '..') out.pop();
    else if (s !== '.' && s !== '') out.push(s);
  }
  return out.join('/');
}

export function relativePath(fromDir: string, to: string): string {
  const a = fromDir ? fromDir.split('/') : [];
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}

export function relsPathFor(part: string) {
  const d = dirname(part);
  const f = part.slice(d.length ? d.length + 1 : 0);
  return (d ? d + '/' : '') + '_rels/' + f + '.rels';
}

export class OpcPackage {
  private xml = new Map<string, XMLDocument>();
  private binaries = new Map<string, Uint8Array>();
  private deleted = new Set<string>();

  private constructor(private zip: JSZip) {}

  static async load(data: ArrayBuffer | Uint8Array): Promise<OpcPackage> {
    const zip = await JSZip.loadAsync(data);
    const pkg = new OpcPackage(zip);
    // Eagerly parse every XML part so that synchronous access is possible afterwards.
    const jobs: Promise<void>[] = [];
    zip.forEach((p, f) => {
      if (f.dir) return;
      if (/\.(xml|rels)$/i.test(p)) {
        jobs.push(f.async('string').then((t) => { pkg.xml.set(p, parseXml(t)); }));
      }
    });
    await Promise.all(jobs);
    return pkg;
  }

  has(p: string) {
    if (this.deleted.has(p)) return false;
    return this.xml.has(p) || this.binaries.has(p) || !!this.zip.file(p);
  }

  getXml(p: string): XMLDocument | null {
    if (this.deleted.has(p)) return null;
    return this.xml.get(p) ?? null;
  }

  setXml(p: string, doc: XMLDocument) {
    this.deleted.delete(p);
    this.xml.set(p, doc);
  }

  async getBinary(p: string): Promise<Uint8Array | null> {
    if (this.deleted.has(p)) return null;
    const b = this.binaries.get(p);
    if (b) return b;
    const f = this.zip.file(p);
    return f ? f.async('uint8array') : null;
  }

  setBinary(p: string, data: Uint8Array) {
    this.deleted.delete(p);
    this.binaries.set(p, data);
  }

  remove(p: string) {
    this.xml.delete(p);
    this.binaries.delete(p);
    this.deleted.add(p);
  }

  partNames(): string[] {
    const names = new Set<string>();
    this.zip.forEach((p, f) => { if (!f.dir) names.add(p); });
    for (const k of this.xml.keys()) names.add(k);
    for (const k of this.binaries.keys()) names.add(k);
    return [...names].filter((n) => !this.deleted.has(n));
  }

  // ---- relationships ----
  rels(part: string): Relationship[] {
    const doc = this.getXml(relsPathFor(part));
    if (!doc) return [];
    const base = dirname(part);
    return children(doc.documentElement).map((r) => {
      const external = r.getAttribute('TargetMode') === 'External';
      const t = r.getAttribute('Target') || '';
      return {
        id: r.getAttribute('Id') || '',
        type: r.getAttribute('Type') || '',
        target: external ? t : resolvePath(base, t),
        external,
      };
    });
  }

  rel(part: string, id: string | null): Relationship | undefined {
    if (!id) return undefined;
    return this.rels(part).find((r) => r.id === id);
  }

  relOfType(part: string, type: string): Relationship | undefined {
    return this.rels(part).find((r) => r.type === type);
  }

  addRel(part: string, type: string, target: string): string {
    const rp = relsPathFor(part);
    let doc = this.getXml(rp);
    if (!doc) {
      doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"/>`);
      this.setXml(rp, doc);
    }
    const used = new Set(children(doc.documentElement).map((r) => r.getAttribute('Id')));
    let n = used.size + 1;
    while (used.has('rId' + n)) n++;
    const id = 'rId' + n;
    const el = doc.createElementNS(NS.rel, 'Relationship');
    el.setAttribute('Id', id);
    el.setAttribute('Type', type);
    el.setAttribute('Target', relativePath(dirname(part), target));
    doc.documentElement.appendChild(el);
    return id;
  }

  removeRel(part: string, id: string) {
    const doc = this.getXml(relsPathFor(part));
    if (!doc) return;
    for (const r of children(doc.documentElement)) if (r.getAttribute('Id') === id) r.remove();
  }

  // ---- content types ----
  private ctDoc(): XMLDocument {
    const d = this.getXml('[Content_Types].xml');
    if (!d) throw new Error('Missing [Content_Types].xml');
    return d;
  }

  addOverride(part: string, contentType: string) {
    const d = this.ctDoc();
    const name = '/' + part;
    if (children(d.documentElement).some((e) => e.getAttribute('PartName') === name)) return;
    const el = d.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', name);
    el.setAttribute('ContentType', contentType);
    d.documentElement.appendChild(el);
  }

  removeOverride(part: string) {
    const d = this.ctDoc();
    for (const e of children(d.documentElement)) if (e.getAttribute('PartName') === '/' + part) e.remove();
  }

  ensureDefault(ext: string, contentType: string) {
    const d = this.ctDoc();
    const has = children(d.documentElement).some(
      (e) => e.localName === 'Default' && e.getAttribute('Extension')?.toLowerCase() === ext.toLowerCase(),
    );
    if (has) return;
    const el = d.createElementNS(NS.ct, 'Default');
    el.setAttribute('Extension', ext);
    el.setAttribute('ContentType', contentType);
    d.documentElement.insertBefore(el, d.documentElement.firstElementChild);
  }

  uniqueName(pattern: (n: number) => string): string {
    let n = 1;
    while (this.has(pattern(n))) n++;
    return pattern(n);
  }

  async save(): Promise<Blob> {
    return new Blob([await this.saveBytes() as BlobPart], {
      type: DOCX_MIME,
    });
  }

  async saveBytes(): Promise<Uint8Array> {
    const out = new JSZip();
    // [Content_Types].xml must be the first entry for some consumers.
    const names = this.partNames().sort((a, b) =>
      a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0,
    );
    for (const n of names) {
      const x = this.xml.get(n);
      if (x) out.file(n, serializeXml(x));
      else {
        const b = await this.getBinary(n);
        if (b) out.file(n, b);
      }
    }
    return out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /** Serialize a set of parts, used for undo snapshots. null marks a missing part. */
  snapshot(parts: string[]): Map<string, string | null> {
    const m = new Map<string, string | null>();
    for (const p of parts) {
      const d = this.getXml(p);
      m.set(p, d ? serializeXml(d) : null);
    }
    return m;
  }

  restore(snap: Map<string, string | null>) {
    for (const [p, s] of snap) {
      if (s == null) this.remove(p);
      else this.setXml(p, parseXml(s));
    }
  }
}

