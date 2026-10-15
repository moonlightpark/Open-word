// Paragraph content model. A w:p is read as a flat list of items (text, tab, break, inline object,
// hidden marker). Every item has a length (text length, 1, or 0) and those lengths define the
// "model offsets" that the editor's DOM positions map to. Commands split runs at offsets,
// change the runs in between, and merge identical neighbours afterwards.
import { NS, PPR_ORDER, RPR_ORDER, child, children, create, ensureOrdered, is, removeChild, setAttr, wa } from './xml';

export type ItemKind = 'text' | 'tab' | 'br' | 'page' | 'atom' | 'field' | 'hidden';

export interface Item {
  kind: ItemKind;
  len: number;
  text: string; // text for 'text', display text for fields/atoms
  /** Paragraph-level XML nodes (runs, bookmarks, fields) that make up the item. */
  nodes: Element[];
  /** True when `nodes` are the live elements (not synthetic copies of part of a run). */
  live: boolean;
  /** The run's direct properties (for text/tab/br/atom). */
  rPr: Element | null;
  /** Inline containers (hyperlink, ins, smartTag…), outermost first. */
  containers: Element[];
  /** For atoms: the run child that defines the object (w:drawing, w:sym, …). */
  obj?: Element;
}

const CONTAINERS = new Set(['hyperlink', 'ins', 'smartTag', 'customXml', 'dir', 'bdo', 'moveTo']);
const ATOM_CHILDREN = new Set(['drawing', 'pict', 'object', 'sym', 'footnoteReference', 'endnoteReference', 'noBreakHyphen', 'AlternateContent', 'ruby']);
const SKIP_CHILDREN = new Set(['rPr', 'lastRenderedPageBreak']);

function runContent(r: Element) {
  return children(r).filter((c) => !(c.namespaceURI === NS.w && SKIP_CHILDREN.has(c.localName)));
}

function runWith(r: Element, c: Element): Element {
  const out = r.cloneNode(false) as Element;
  const rPr = child(r, 'w:rPr');
  if (rPr) out.appendChild(rPr.cloneNode(true));
  out.appendChild(c.cloneNode(true));
  return out;
}

function fldCharType(r: Element): string | null {
  const f = child(r, 'w:fldChar');
  return f ? wa(f, 'fldCharType') : null;
}

/** Flatten a paragraph into items. */
export function paraItems(p: Element): Item[] {
  const items: Item[] = [];
  let field: { nodes: Element[]; depth: number; instr: string; result: string; sep: boolean; rPr: Element | null; containers: Element[] } | null = null;

  // Field characters may share one run (begin, instrText, separate, result, end).
  const feed = (r: Element) => {
    const f = field!;
    for (const c of children(r)) {
      if (is(c, 'w:fldChar')) {
        const t = wa(c, 'fldCharType');
        if (t === 'begin') f.depth++;
        else if (t === 'separate' && f.depth === 1) f.sep = true;
        else if (t === 'end') f.depth--;
      } else if (is(c, 'w:instrText') && !f.sep && f.depth === 1) f.instr += c.textContent ?? '';
      else if (is(c, 'w:t') && f.sep && f.depth === 1) { f.result += c.textContent ?? ''; if (!f.rPr) f.rPr = child(r, 'w:rPr'); }
      else if (is(c, 'w:tab') && f.sep && f.depth === 1) f.result += '\t';
    }
    if (f.depth <= 0) {
      items.push({ kind: 'field', len: 1, text: f.result, nodes: f.nodes, live: true, rPr: f.rPr, containers: f.containers });
      (items.at(-1) as Item & { instr?: string }).instr = f.instr;
      field = null;
    }
  };

  const scan = (parent: Element, containers: Element[]) => {
    for (const n of children(parent)) {
      const local = n.localName;
      const isW = n.namespaceURI === NS.w;
      if (field) {
        // Inside a complex field every node belongs to the field until the matching end.
        field.nodes.push(n);
        if (isW && local === 'r') feed(n);
        continue;
      }
      if (!isW) {
        if (n.namespaceURI === NS.m) items.push({ kind: 'atom', len: 1, text: n.textContent ?? '', nodes: [n], live: true, rPr: null, containers, obj: n });
        else items.push({ kind: 'hidden', len: 0, text: '', nodes: [n], live: true, rPr: null, containers });
        continue;
      }
      if (local === 'pPr') continue;
      if (CONTAINERS.has(local)) { scan(n, [...containers, n]); continue; }
      if (local === 'fldSimple') {
        items.push({ kind: 'field', len: 1, text: n.textContent ?? '', nodes: [n], live: true, rPr: child(child(n, 'w:r'), 'w:rPr'), containers });
        (items.at(-1) as Item & { instr?: string }).instr = wa(n, 'instr') ?? '';
        continue;
      }
      if (local === 'sdt') {
        items.push({ kind: 'atom', len: 1, text: child(n, 'w:sdtContent')?.textContent ?? '', nodes: [n], live: true, rPr: null, containers, obj: n });
        continue;
      }
      if (local !== 'r') { items.push({ kind: 'hidden', len: 0, text: '', nodes: [n], live: true, rPr: null, containers }); continue; }
      // A run.
      const t = fldCharType(n);
      if (t === 'begin') {
        field = { nodes: [n], depth: 0, instr: '', result: '', sep: false, rPr: null, containers };
        feed(n);
        continue;
      }
      if (t) { items.push({ kind: 'hidden', len: 0, text: '', nodes: [n], live: true, rPr: null, containers }); continue; }
      const rPr = child(n, 'w:rPr');
      const content = runContent(n);
      const single = content.length === 1;
      for (const c of content) {
        const nodes = single ? [n] : [runWith(n, c)];
        const base = { nodes, live: single, rPr, containers };
        if (c.namespaceURI === NS.w) {
          switch (c.localName) {
            case 't': { const text = c.textContent ?? ''; if (text) items.push({ kind: 'text', len: text.length, text, ...base }); continue; }
            case 'tab': case 'ptab': items.push({ kind: 'tab', len: 1, text: '\t', ...base }); continue;
            case 'br': case 'cr': {
              const type = wa(c, 'type');
              items.push({ kind: type === 'page' ? 'page' : 'br', len: 1, text: '\n', ...base });
              continue;
            }
            case 'delText': case 'instrText': case 'commentReference': case 'softHyphen': case 'annotationRef': case 'separator': case 'continuationSeparator':
              items.push({ kind: 'hidden', len: 0, text: '', ...base });
              continue;
          }
          if (ATOM_CHILDREN.has(c.localName)) { items.push({ kind: 'atom', len: 1, text: '', ...base, obj: c }); continue; }
        } else if (c.namespaceURI === NS.mc) {
          items.push({ kind: 'atom', len: 1, text: '', ...base, obj: c });
          continue;
        }
        items.push({ kind: 'hidden', len: 0, text: '', ...base });
      }
    }
  };
  scan(p, []);
  // An unterminated field (it continues into following paragraphs, e.g. a TOC) is one atom here.
  const open = field as { nodes: Element[]; result: string; rPr: Element | null; containers: Element[] } | null;
  if (open) items.push({ kind: 'field', len: 1, text: open.result, nodes: open.nodes, live: true, rPr: open.rPr, containers: open.containers });
  return items;
}

/** Plain text of a paragraph using the model's length conventions (objects are U+FFFC). */
export function paraText(p: Element): string {
  return paraItems(p).map((it) => it.kind === 'text' ? it.text : it.kind === 'tab' ? '\t' : it.kind === 'br' || it.kind === 'page' ? '\n' : it.len ? '￼' : '').join('');
}

export function paraLength(p: Element) {
  return paraItems(p).reduce((n, it) => n + it.len, 0);
}

/** Paragraphs in reading order: body, then into tables (rows → cells), recursively. */
export function walkParagraphs(container: Element | null, out: Element[] = []): Element[] {
  for (const c of children(container)) {
    if (c.namespaceURI !== NS.w) continue;
    if (c.localName === 'p') out.push(c);
    else if (c.localName === 'tbl') {
      for (const tr of children(c, 'w:tr')) for (const tc of children(tr, 'w:tc')) walkParagraphs(tc, out);
    }
  }
  return out;
}

// ---------------------------------------------------------------- run surgery

/** Split every multi-content run so each item is backed by its own live node(s). */
export function explode(p: Element) {
  const visit = (parent: Element) => {
    for (const n of children(parent)) {
      if (n.namespaceURI !== NS.w) continue;
      if (CONTAINERS.has(n.localName)) { visit(n); continue; }
      if (n.localName !== 'r') continue;
      const lrpb = children(n, 'w:lastRenderedPageBreak');
      lrpb.forEach((x) => x.remove());
      const content = runContent(n);
      if (content.length <= 1) continue;
      for (const c of content) parent.insertBefore(runWith(n, c), n);
      n.remove();
    }
  };
  visit(p);
}

/** Make sure a run boundary exists at `off` (paragraph must be exploded). */
export function splitAt(p: Element, off: number) {
  let pos = 0;
  for (const it of paraItems(p)) {
    if (it.kind === 'text' && pos < off && off < pos + it.len) {
      const r = it.nodes[0];
      const t = child(r, 'w:t')!;
      const k = off - pos;
      const after = r.cloneNode(true) as Element;
      setText(t, it.text.slice(0, k));
      setText(child(after, 'w:t')!, it.text.slice(k));
      r.after(after);
      return;
    }
    pos += it.len;
    if (pos > off) return;
  }
}

export function setText(t: Element, s: string) {
  t.textContent = s;
  if (/^\s|\s$/.test(s)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  else t.removeAttributeNS('http://www.w3.org/XML/1998/namespace', 'space');
}

/** Items fully inside [start, end) after splitting the boundaries. */
export function itemsInRange(p: Element, start: number, end: number): Item[] {
  explode(p);
  splitAt(p, end);
  splitAt(p, start);
  const out: Item[] = [];
  let pos = 0;
  for (const it of paraItems(p)) {
    if (it.len && pos >= start && pos + it.len <= end) out.push(it);
    pos += it.len;
  }
  return out;
}

/** Runs (w:r) that make up an item, including those inside simple fields. */
export function itemRuns(it: Item): Element[] {
  const out: Element[] = [];
  for (const n of it.nodes) {
    if (is(n, 'w:r')) out.push(n);
    else if (is(n, 'w:fldSimple')) out.push(...children(n, 'w:r'));
  }
  return out;
}

/** Merge adjacent text runs whose properties are identical. */
export function mergeRuns(p: Element) {
  const ser = new XMLSerializer();
  const key = (r: Element) => {
    const rPr = child(r, 'w:rPr');
    return rPr ? ser.serializeToString(rPr).replace(/ xmlns(:\w+)?="[^"]*"/g, '') : '';
  };
  const isPlain = (r: Element | null): r is Element => !!r && is(r, 'w:r') && runContent(r).length === 1 && is(runContent(r)[0], 'w:t') && r.attributes.length === 0;
  const visit = (parent: Element) => {
    for (const n of children(parent)) if (n.namespaceURI === NS.w && CONTAINERS.has(n.localName)) visit(n);
    let prev: Element | null = null;
    for (const n of children(parent)) {
      if (isPlain(n) && isPlain(prev) && key(prev) === key(n)) {
        const t = child(prev, 'w:t')!;
        setText(t, (t.textContent ?? '') + (child(n, 'w:t')!.textContent ?? ''));
        n.remove();
        continue;
      }
      prev = n;
    }
  };
  visit(p);
}

/** Insert nodes at model offset `off` (they land at the container level of the preceding item). */
export function insertAt(p: Element, off: number, nodes: Element[]) {
  explode(p);
  splitAt(p, off);
  let pos = 0;
  let after: Element | null = null;
  for (const it of paraItems(p)) {
    if (pos >= off) break;
    after = it.nodes[it.nodes.length - 1];
    pos += it.len;
  }
  if (after) {
    let ref: Element = after;
    for (const n of nodes) { ref.after(n); ref = n; }
  } else {
    const pPr = child(p, 'w:pPr');
    const ref: Node | null = pPr ? pPr.nextSibling : p.firstChild;
    for (const n of nodes) p.insertBefore(n, ref);
  }
}

/** Split paragraph `p` at `off`; returns the new second paragraph (inserted after `p`). */
export function splitParagraph(p: Element, off: number): Element {
  explode(p);
  splitAt(p, off);
  const second = p.cloneNode(false) as Element;
  for (const a of ['paraId', 'textId']) second.removeAttributeNS(NS.w14, a);
  const pPr = child(p, 'w:pPr');
  if (pPr) {
    const c = pPr.cloneNode(true) as Element;
    second.appendChild(c);
    // A section break stays with the last part of the paragraph.
    removeChild(pPr, 'w:sectPr');
  }
  let pos = 0;
  const containerCopies = new Map<Element, Element>();
  for (const it of paraItems(p)) {
    if (pos >= off && !(it.len === 0 && pos === off && it.kind === 'hidden' && isRangeEnd(it.nodes[0]))) {
      // Containers are duplicated so a link split in two remains a link on both sides.
      let target: Element = second;
      for (const c of it.containers) {
        let copy = containerCopies.get(c);
        if (!copy) { copy = c.cloneNode(false) as Element; target.appendChild(copy); containerCopies.set(c, copy); }
        target = copy;
      }
      for (const n of it.nodes) target.appendChild(n);
    }
    pos += it.len;
  }
  // Drop containers left empty in the first paragraph.
  const prune = (el: Element) => { for (const c of children(el)) { if (CONTAINERS.has(c.localName) && c.namespaceURI === NS.w) { prune(c); if (!c.firstElementChild) c.remove(); } } };
  prune(p);
  p.after(second);
  return second;
}

function isRangeEnd(n: Element) {
  return n.namespaceURI === NS.w && (n.localName === 'bookmarkEnd' || n.localName === 'commentRangeEnd');
}

/** Paragraph-mark run properties converted into properties for a new run. */
export function markRPr(p: Element | null): Element | null {
  const m = child(child(p, 'w:pPr'), 'w:rPr');
  if (!m) return null;
  const r = m.cloneNode(true) as Element;
  for (const c of children(r)) if (['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(c.localName)) c.remove();
  return r.firstElementChild ? r : null;
}

export function makeRun(doc: Document, rPr: Element | null, content: Element | string): Element {
  const r = create(doc, 'w:r');
  if (rPr && rPr.firstElementChild) r.appendChild(rPr.cloneNode(true));
  if (typeof content === 'string') {
    const t = create(doc, 'w:t');
    setText(t, content);
    r.appendChild(t);
  } else r.appendChild(content);
  return r;
}

// ---------------------------------------------------------------- property writers

export interface RunProps {
  b?: boolean; i?: boolean; u?: boolean | string; strike?: boolean;
  sz?: number; // points
  color?: string | null; // RRGGBB, null = automatic
  highlight?: string | null; // highlight name, null = none
  font?: string;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
}

function toggle(rPr: Element, qname: string, on: boolean | undefined) {
  if (on === undefined) return;
  const el = ensureOrdered(rPr, qname, RPR_ORDER);
  if (on) el.removeAttributeNS(NS.w, 'val'); else setAttr(el, 'w:val', '0');
}

export function setRunProps(rPr: Element, p: RunProps) {
  toggle(rPr, 'w:b', p.b);
  if (p.b !== undefined) toggle(rPr, 'w:bCs', p.b);
  toggle(rPr, 'w:i', p.i);
  if (p.i !== undefined) toggle(rPr, 'w:iCs', p.i);
  toggle(rPr, 'w:strike', p.strike);
  if (p.strike) removeChild(rPr, 'w:dstrike');
  if (p.u !== undefined) setAttr(ensureOrdered(rPr, 'w:u', RPR_ORDER), 'w:val', p.u === true ? 'single' : p.u === false ? 'none' : p.u);
  if (p.sz !== undefined) {
    const hp = String(Math.round(p.sz * 2));
    setAttr(ensureOrdered(rPr, 'w:sz', RPR_ORDER), 'w:val', hp);
    setAttr(ensureOrdered(rPr, 'w:szCs', RPR_ORDER), 'w:val', hp);
  }
  if (p.color !== undefined) {
    const el = ensureOrdered(rPr, 'w:color', RPR_ORDER);
    for (const a of ['themeColor', 'themeShade', 'themeTint']) el.removeAttributeNS(NS.w, a);
    setAttr(el, 'w:val', p.color ?? 'auto');
  }
  if (p.highlight !== undefined) {
    if (p.highlight) setAttr(ensureOrdered(rPr, 'w:highlight', RPR_ORDER), 'w:val', p.highlight);
    else { removeChild(rPr, 'w:highlight'); }
  }
  if (p.font !== undefined) {
    const el = ensureOrdered(rPr, 'w:rFonts', RPR_ORDER);
    for (const a of ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']) el.removeAttributeNS(NS.w, a);
    for (const a of ['w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs']) setAttr(el, a, p.font);
  }
  if (p.vertAlign !== undefined) {
    if (p.vertAlign === 'baseline') removeChild(rPr, 'w:vertAlign');
    else setAttr(ensureOrdered(rPr, 'w:vertAlign', RPR_ORDER), 'w:val', p.vertAlign);
  }
}

export function ensureRPr(r: Element): Element {
  let rPr = child(r, 'w:rPr');
  if (!rPr) { rPr = create(r.ownerDocument, 'w:rPr'); r.insertBefore(rPr, r.firstChild); }
  return rPr;
}

export function ensurePPr(p: Element): Element {
  let pPr = child(p, 'w:pPr');
  if (!pPr) { pPr = create(p.ownerDocument, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
  return pPr;
}

export interface ParaProps {
  jc?: string;
  left?: number | null; right?: number | null; firstLine?: number | null; hanging?: number | null; // twips
  before?: number | null; after?: number | null; // twips
  line?: number | null; lineRule?: string;
  style?: string | null;
  numId?: string | null; ilvl?: number;
}

export function setParaProps(p: Element, v: ParaProps) {
  const pPr = ensurePPr(p);
  if (v.style !== undefined) {
    if (v.style) setAttr(ensureOrdered(pPr, 'w:pStyle', PPR_ORDER), 'w:val', v.style);
    else removeChild(pPr, 'w:pStyle');
  }
  if (v.jc !== undefined) setAttr(ensureOrdered(pPr, 'w:jc', PPR_ORDER), 'w:val', v.jc);
  if (v.left !== undefined || v.right !== undefined || v.firstLine !== undefined || v.hanging !== undefined) {
    const ind = ensureOrdered(pPr, 'w:ind', PPR_ORDER);
    const put = (name: string, alt: string, val: number | null | undefined) => {
      if (val === undefined) return;
      ind.removeAttributeNS(NS.w, alt);
      ind.removeAttributeNS(NS.w, alt + 'Chars');
      ind.removeAttributeNS(NS.w, name + 'Chars');
      if (val === null) ind.removeAttributeNS(NS.w, name); else setAttr(ind, 'w:' + name, Math.round(val));
    };
    put('left', 'start', v.left);
    put('right', 'end', v.right);
    if (v.hanging != null && v.hanging > 0) { put('hanging', 'firstLine', v.hanging); }
    else if (v.firstLine != null) { put('firstLine', 'hanging', v.firstLine); }
    else if (v.firstLine === null || v.hanging === null) { ind.removeAttributeNS(NS.w, 'firstLine'); ind.removeAttributeNS(NS.w, 'hanging'); }
    if (!ind.attributes.length) ind.remove();
  }
  if (v.before !== undefined || v.after !== undefined || v.line !== undefined) {
    const sp = ensureOrdered(pPr, 'w:spacing', PPR_ORDER);
    const put = (name: string, val: number | null | undefined) => {
      if (val === undefined) return;
      sp.removeAttributeNS(NS.w, name + 'Lines');
      sp.removeAttributeNS(NS.w, name + 'Autospacing');
      if (val === null) sp.removeAttributeNS(NS.w, name); else setAttr(sp, 'w:' + name, Math.round(val));
    };
    put('before', v.before);
    put('after', v.after);
    if (v.line !== undefined) {
      if (v.line === null) { sp.removeAttributeNS(NS.w, 'line'); sp.removeAttributeNS(NS.w, 'lineRule'); }
      else { setAttr(sp, 'w:line', Math.round(v.line)); setAttr(sp, 'w:lineRule', v.lineRule ?? 'auto'); }
    }
    if (!sp.attributes.length) sp.remove();
  }
  if (v.numId !== undefined) {
    if (v.numId === null) removeChild(pPr, 'w:numPr');
    else {
      const np = ensureOrdered(pPr, 'w:numPr', PPR_ORDER);
      while (np.firstChild) np.firstChild.remove();
      np.appendChild(create(p.ownerDocument, 'w:ilvl', { 'w:val': v.ilvl ?? 0 }));
      np.appendChild(create(p.ownerDocument, 'w:numId', { 'w:val': v.numId }));
    }
  } else if (v.ilvl !== undefined) {
    const np = child(pPr, 'w:numPr');
    if (np) {
      let il = child(np, 'w:ilvl');
      if (!il) { il = create(p.ownerDocument, 'w:ilvl'); np.insertBefore(il, np.firstChild); }
      setAttr(il, 'w:val', v.ilvl);
    }
  }
  if (!pPr.firstElementChild && !pPr.attributes.length) pPr.remove();
}
