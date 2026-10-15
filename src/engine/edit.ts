// Bridge between the contenteditable DOM and the model:
// - DOM selection ↔ model positions {pi, off} using the same item lengths as model.ts
// - committing edited DOM back to w:p / w:tbl elements. Unchanged paragraphs keep their original
//   XML; changed ones are rebuilt from the registered run properties and inline objects, so
//   formatting, fields, images and bookmarks survive typing.
import { DocxDocument, Pos } from './document';
import { makeRun, markRPr, setText } from './model';
import { Registry } from './render';
import { NS, child, children, create, is } from './xml';

const ZWSP = /​/g;

export function isPara(n: Node | null): n is HTMLElement {
  return n instanceof HTMLElement && n.tagName === 'P' && n.classList.contains('w-p');
}

function atomLen(el: HTMLElement) { return el.dataset.len === '0' ? 0 : 1; }

/** Last node in walk order when it is a <br> (the placeholder that keeps an empty line). */
function trailingBr(p: HTMLElement): HTMLElement | null {
  // The last <br> that has no visible content after it (empty spans and hidden markers don't count).
  let cand: HTMLElement | null = null;
  const visit = (node: Node) => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === 3) { if ((n.textContent ?? '').replace(ZWSP, '')) cand = null; continue; }
      if (!(n instanceof HTMLElement)) continue;
      if (n.tagName === 'BR') { cand = n; continue; }
      if (n.dataset.a != null) { if (atomLen(n)) cand = null; continue; }
      visit(n);
    }
  };
  visit(p);
  return cand;
}

/** Model offset of a DOM boundary point inside paragraph element `p`. */
export function measure(p: HTMLElement, stopNode: Node, stopOffset: number): number {
  let n = 0;
  let done = false;
  const tail = trailingBr(p);
  const visit = (node: Node) => {
    if (done) return;
    if (node.nodeType === 3) {
      const t = node.textContent ?? '';
      if (node === stopNode) { n += t.slice(0, stopOffset).replace(ZWSP, '').length; done = true; return; }
      n += t.replace(ZWSP, '').length;
      return;
    }
    if (!(node instanceof HTMLElement)) return;
    if (node.dataset.a != null) {
      if (node === stopNode || node.contains(stopNode)) {
        // A boundary inside an object: before it when at its very start, otherwise after.
        n += node === stopNode && stopOffset === 0 ? 0 : atomLen(node);
        done = true;
        return;
      }
      n += atomLen(node);
      return;
    }
    if (node.tagName === 'BR') { if (node !== tail) n += 1; return; }
    const kids = Array.from(node.childNodes);
    if (node === stopNode) {
      for (const k of kids.slice(0, stopOffset)) visit(k);
      done = true;
      return;
    }
    for (const k of kids) { visit(k); if (done) return; }
  };
  visit(p);
  return n;
}

export function paraElOf(node: Node | null, root: HTMLElement): HTMLElement | null {
  for (let n: Node | null = node; n && n !== root; n = n.parentNode) if (isPara(n)) return n;
  return null;
}

export function domToPos(root: HTMLElement, node: Node, offset: number): Pos | null {
  let p = paraElOf(node, root);
  if (!p) {
    // A boundary between blocks: use the nearest paragraph.
    const ref = node === root ? root.childNodes[Math.min(offset, root.childNodes.length - 1)] : node;
    const el = ref instanceof HTMLElement ? ref : ref?.parentElement;
    p = (el?.matches?.('p.w-p') ? el : el?.querySelector?.('p.w-p')) as HTMLElement | null;
    if (!p) return null;
    return { pi: Number(p.dataset.pi), off: 0 };
  }
  if (p.dataset.pi == null) return null;
  return { pi: Number(p.dataset.pi), off: measure(p, node, offset) };
}

/** DOM boundary point for a model offset in paragraph `p`. Prefers the end of the previous text. */
export function posToDom(p: HTMLElement, off: number): [Node, number] {
  let n = 0;
  let result: [Node, number] | null = null;
  const tail = trailingBr(p);
  const visit = (node: Node): boolean => {
    if (node.nodeType === 3) {
      const t = node.textContent ?? '';
      const len = t.replace(ZWSP, '').length;
      if (off <= n + len && (len > 0 || off === n)) {
        let k = 0, seen = 0;
        while (k < t.length && seen < off - n) { if (t[k] !== '​') seen++; k++; }
        while (k < t.length && t[k] === '​') k++;
        result = [node, k];
        return true;
      }
      n += len;
      return false;
    }
    if (!(node instanceof HTMLElement)) return false;
    const parent = node.parentNode!;
    const idx = Array.prototype.indexOf.call(parent.childNodes, node);
    if (node.dataset.a != null || node.tagName === 'BR') {
      const len = node.tagName === 'BR' && node.dataset.a == null ? (node === tail ? 0 : 1) : atomLen(node);
      if (off === n && len > 0) { result = [parent, idx]; return true; }
      n += len;
      if (off === n && len > 0) { result = [parent, idx + 1]; return true; }
      return false;
    }
    for (const k of Array.from(node.childNodes)) if (visit(k)) return true;
    return false;
  };
  for (const k of Array.from(p.childNodes)) if (visit(k)) break;
  if (result) return result;
  if (tail && tail.parentNode) return [tail.parentNode, Array.prototype.indexOf.call(tail.parentNode.childNodes, tail)];
  return [p, p.childNodes.length];
}

// ---------------------------------------------------------------- commit

interface CommitState { doc: DocxDocument; reg: Registry; used: Set<Element>; pairs: [HTMLElement, Element][]; changed: number }

/**
 * Rebuild the body from the editor DOM. Returns the (DOM paragraph, XML paragraph) pairs so the
 * editor can refresh its bookkeeping, plus the number of rebuilt paragraphs.
 */
export function commitDom(root: HTMLElement, doc: DocxDocument, reg: Registry): { pairs: [HTMLElement, Element][]; changed: number } {
  const st: CommitState = { doc, reg, used: new Set(), pairs: [], changed: 0 };
  const body = doc.body;
  const sect = child(body, 'w:sectPr');
  const blocks = buildBlocks(root, st, true);
  if (!blocks.some((b) => is(b, 'w:p')) || !is(blocks[blocks.length - 1], 'w:p')) blocks.push(create(doc.xml, 'w:p'));
  for (const c of children(body)) if (c !== sect) c.remove();
  for (const b of blocks) body.insertBefore(b, sect);
  return { pairs: st.pairs, changed: st.changed };
}

function buildBlocks(host: HTMLElement, st: CommitState, isBody: boolean): Element[] {
  const out: Element[] = [];
  let loose: HTMLElement | null = null;
  const flushLoose = () => {
    if (loose) { out.push(rebuildParagraph(loose, null, st)); loose = null; }
  };
  for (const n of Array.from(host.childNodes)) {
    if (n instanceof HTMLElement) {
      if (n.tagName === 'P' || (n.tagName === 'DIV' && !st.reg.dom.has(n))) {
        flushLoose();
        out.push(paragraphFor(n, st));
        continue;
      }
      if (n.tagName === 'TABLE') {
        flushLoose();
        const t = tableFor(n as HTMLTableElement, st);
        if (t) out.push(...t);
        continue;
      }
      const mapped = st.reg.dom.get(n);
      if (mapped && n.tagName === 'DIV') { flushLoose(); out.push(mapped); continue; }
    }
    if (n.nodeType === 3 && !(n.textContent ?? '').replace(ZWSP, '').trim()) continue;
    if (n.nodeType !== 1 && n.nodeType !== 3) continue;
    if (!loose) loose = document.createElement('p');
    loose.appendChild(n.cloneNode(true));
  }
  flushLoose();
  void isBody;
  return out;
}

function paragraphFor(el: HTMLElement, st: CommitState): Element {
  const orig = st.reg.dom.get(el);
  if (orig && is(orig, 'w:p') && !st.used.has(orig) && st.reg.sig.get(el) === el.innerHTML) {
    st.used.add(orig);
    st.pairs.push([el, orig]);
    return orig;
  }
  // Rebuilt: from its own source paragraph, or (for browser-made copies) the one it was cloned from.
  const src = orig && is(orig, 'w:p') ? orig : st.doc.para(Number(el.dataset.pi ?? -1));
  const p = rebuildParagraph(el, src, st);
  if (src && st.used.has(src)) for (const a of ['paraId', 'textId']) p.removeAttributeNS(NS.w14, a);
  if (src) st.used.add(src);
  st.pairs.push([el, p]);
  st.changed++;
  return p;
}

function tableFor(el: HTMLTableElement, st: CommitState): Element[] | null {
  const tbl = st.reg.dom.get(el);
  if (!tbl || !is(tbl, 'w:tbl') || st.used.has(tbl)) {
    // A table that did not come from the document (e.g. dragged copy): keep its text as paragraphs.
    const out: Element[] = [];
    for (const td of Array.from(el.querySelectorAll('td'))) out.push(...buildBlocks(td, st, false));
    return out;
  }
  st.used.add(tbl);
  for (const tr of Array.from(el.rows)) {
    for (const td of Array.from(tr.cells)) {
      const tc = st.reg.dom.get(td);
      if (!tc || !is(tc, 'w:tc')) continue;
      const kids = buildBlocks(td, st, false);
      if (!kids.length || !is(kids[kids.length - 1], 'w:p')) kids.push(create(st.doc.xml, 'w:p'));
      const tcPr = child(tc, 'w:tcPr');
      for (const c of children(tc)) if (c !== tcPr) c.remove();
      for (const k of kids) tc.appendChild(k);
    }
  }
  return [tbl];
}

/** Build a w:p from an edited paragraph element. */
export function rebuildParagraph(el: HTMLElement, src: Element | null, st: CommitState): Element {
  const { doc, reg } = st;
  const xml = doc.xml;
  const p = src ? src.cloneNode(false) as Element : create(xml, 'w:p');
  const pPr = child(src, 'w:pPr');
  if (pPr) p.appendChild(pPr.cloneNode(true));
  const mark = markRPr(src);
  let lastRPr: Element | null = null;
  const tail = trailingBr(el);

  // Container nesting currently open in the XML (hyperlinks etc.).
  const open: { id: string; xml: Element }[] = [];
  const target = (chain: string[]): Element => {
    let k = 0;
    while (k < open.length && k < chain.length && open[k].id === chain[k]) k++;
    open.length = k;
    for (; k < chain.length; k++) {
      const tpl = reg.containers[Number(chain[k])];
      if (!tpl) break;
      const c = tpl.cloneNode(false) as Element;
      (open[k - 1]?.xml ?? p).appendChild(c);
      open.push({ id: chain[k], xml: c });
    }
    return open.at(-1)?.xml ?? p;
  };
  const chainOf = (node: Node): string[] => {
    const ids: string[] = [];
    for (let n = node.parentElement; n && n !== el; n = n.parentElement) if (n.dataset.c != null) ids.unshift(n.dataset.c);
    return ids;
  };
  const rPrOf = (node: Node): Element | null => {
    for (let n = node.parentElement; n && n !== el; n = n.parentElement) {
      if (n.dataset.r != null) {
        const v = reg.runs[Number(n.dataset.r)];
        return v === undefined ? lastRPr : v;
      }
    }
    return lastRPr ?? mark;
  };
  const ser = new XMLSerializer();
  const sameRPr = (a: Element | null, b: Element | null) => (a ? ser.serializeToString(a) : '') === (b ? ser.serializeToString(b) : '');

  const emitText = (text: string, node: Node) => {
    const parent = target(chainOf(node));
    const rPr = rPrOf(node);
    lastRPr = rPr;
    const parts = text.split(/(\t|\n)/);
    for (const part of parts) {
      if (!part) continue;
      if (part === '\t' || part === '\n') { parent.appendChild(makeRun(xml, rPr, create(xml, part === '\t' ? 'w:tab' : 'w:br'))); continue; }
      const prev = parent.lastElementChild;
      if (prev && is(prev, 'w:r') && children(prev).length === (child(prev, 'w:rPr') ? 2 : 1) && child(prev, 'w:t') && sameRPr(child(prev, 'w:rPr'), rPr && rPr.firstElementChild ? rPr : null)) {
        const t = child(prev, 'w:t')!;
        setText(t, (t.textContent ?? '') + part);
      } else parent.appendChild(makeRun(xml, rPr, part));
    }
  };

  const walk = (node: Node) => {
    for (const n of Array.from(node.childNodes)) {
      if (n.nodeType === 3) {
        const text = (n.textContent ?? '').replace(ZWSP, '');
        if (text) emitText(text, n);
        continue;
      }
      if (!(n instanceof HTMLElement)) continue;
      if (n.dataset.a != null) {
        const nodes = reg.atoms[Number(n.dataset.a)];
        if (!nodes) continue;
        const parent = target(chainOf(n));
        for (const x of nodes) parent.appendChild(x.cloneNode(true));
        const r = nodes.find((x) => is(x, 'w:r'));
        if (r && !child(r, 'w:fldChar')) lastRPr = child(r, 'w:rPr');
        continue;
      }
      if (n.tagName === 'BR') {
        if (n !== tail) target(chainOf(n)).appendChild(makeRun(xml, rPrOf(n), create(xml, 'w:br')));
        continue;
      }
      if (n.tagName === 'IMG') continue;
      walk(n);
    }
  };
  walk(el);
  return p;
}
