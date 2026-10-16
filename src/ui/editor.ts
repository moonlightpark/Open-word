// Document editor: one contenteditable flow laid out on page frames.
// Typing is handled natively by the browser and committed to the model on idle (edit.ts);
// structural keys (Enter, Tab) and every formatting command go through the engine, after which
// the affected paragraphs are re-rendered and the selection is restored in model coordinates.
import { DocxDocument, Pos, Range, comparePos, normRange } from '../engine/document';
import { commitContainer, commitDom, domToPos, isPara, measure, paraElOf, posToDom } from '../engine/edit';
import { RunProps, setRunProps } from '../engine/model';
import { ImageCache, Registry, newRegistry, renderBlocks, renderHeaderFooter, renderParagraph, runCss } from '../engine/render';
import { create, is } from '../engine/xml';

export interface EditorHooks {
  doc: () => DocxDocument | null;
  images: () => ImageCache;
  selectionChanged: () => void;
  layoutChanged: () => void;
  status: (msg: string) => void;
  hfChanged?: () => void;
}

export interface HFEdit { kind: 'header' | 'footer'; part: string; page: number; el: HTMLElement; reg: Registry; pending: boolean }

export interface BlockBox { el: HTMLElement; y: number; h: number; mt: number; page: number; endPage: number }

const PX = 4 / 3; // CSS px per pt

export class DocEditor {
  readonly root: HTMLElement;
  readonly frames: HTMLElement;
  readonly host: HTMLElement;
  zoom = 1;
  gapPt = 14;
  pages = 1;
  layout: BlockBox[] = [];
  showMarks = false;
  reg: Registry = newRegistry();
  private pending = false;
  private composing = false;
  private commitTimer = 0;
  private layoutQueued = false;
  private pendingFmt: { pos: Pos; span: HTMLElement } | null = null;
  private selImg: HTMLElement | null = null;
  private lastRange: Range | null = null;

  constructor(private hooks: EditorHooks, private wrap: HTMLElement) {
    this.host = document.createElement('div');
    this.host.id = 'page-host';
    this.frames = document.createElement('div');
    this.frames.id = 'page-frames';
    this.root = document.createElement('div');
    this.root.id = 'doc-root';
    this.root.contentEditable = 'true';
    this.root.spellcheck = false;
    this.host.append(this.frames, this.root);
    wrap.appendChild(this.host);
    this.bind();
  }

  private get doc() { return this.hooks.doc()!; }

  // ---------------------------------------------------------------- rendering

  private ctx() {
    return { doc: this.doc, images: this.hooks.images(), part: this.doc.docPath, reg: this.reg, editable: true };
  }

  render() {
    const doc = this.hooks.doc();
    this.hideTableHandle();
    this.root.innerHTML = '';
    this.reg = newRegistry();
    this.pending = false;
    this.pendingFmt = null;
    this.selImg = null;
    if (!doc) return;
    renderBlocks(doc.body, this.ctx(), this.root);
    this.root.querySelectorAll('img').forEach((img) => img.addEventListener('load', () => this.layoutSoon(), { once: true }));
    this.paginate();
  }

  /** Re-render paragraphs [start, start+oldCount) as the model's [start, start+newCount). */
  refresh(start: number, oldCount: number, newCount: number) {
    const doc = this.doc;
    const olds = Array.from(this.root.querySelectorAll<HTMLElement>('p.w-p[data-pi]')).filter((p) => {
      const i = Number(p.dataset.pi);
      return i >= start && i < start + oldCount;
    });
    const paras = doc.paragraphs().slice(start, start + newCount);
    const contiguous = olds.length === oldCount && olds.every((p, k) => k === 0 || olds[k - 1].nextElementSibling === p);
    const parent = olds[0]?.parentElement;
    const xmlParent = paras[0]?.parentElement;
    const sameParent = paras.every((p) => p.parentElement === xmlParent) && parent && this.reg.dom.get(parent) === xmlParent;
    const bodyParent = parent === this.root && xmlParent === doc.body;
    if (!oldCount || !newCount || !contiguous || !(sameParent || bodyParent)) { this.render(); return; }
    const ctx = this.ctx();
    const styleOf = (el: Element | null) => el && is(el, 'w:p') ? doc.pStyleOf(el) ?? doc.styles.defaultPara : undefined;
    const els = paras.map((p) => {
      const f = doc.paraFmt(p);
      const st = styleOf(p);
      return renderParagraph(p, ctx, { noBefore: !!f.contextual && styleOf(p.previousElementSibling) === st, noAfter: !!f.contextual && styleOf(p.nextElementSibling) === st });
    });
    olds[0].before(...els);
    olds.forEach((o) => o.remove());
    els.forEach((e) => e.querySelectorAll('img').forEach((img) => img.addEventListener('load', () => this.layoutSoon(), { once: true })));
    this.reindex();
    this.paginate();
  }

  /** Refresh data-pi and list labels after the paragraph list changed. */
  reindex() {
    const doc = this.doc;
    const labels = doc.labels();
    for (const p of Array.from(this.root.querySelectorAll<HTMLElement>('p.w-p'))) {
      const x = this.reg.dom.get(p);
      if (!x) continue;
      const i = doc.indexOf(x);
      if (i >= 0) p.dataset.pi = String(i);
      const lab = labels.get(x);
      if (lab) p.dataset.label = lab.text + (lab.level.suff === 'space' ? ' ' : '');
      else if (p.dataset.label) delete p.dataset.label;
    }
  }

  // ---------------------------------------------------------------- pagination

  geometry() {
    const ps = this.doc.pageSetup();
    const W = ps.w / 20, H = ps.h / 20;
    return { ps, W, H, top: ps.top / 20, bottom: ps.bottom / 20, left: ps.left / 20, right: ps.right / 20, CW: (ps.w - ps.left - ps.right) / 20 };
  }

  layoutSoon() {
    if (this.layoutQueued) return;
    this.layoutQueued = true;
    requestAnimationFrame(() => { this.layoutQueued = false; this.paginate(); });
  }

  /**
   * Push blocks that would cross a page bottom to the next page (via --push margins) and build the
   * page frames. Whole paragraphs move, like Word with "keep lines together".
   */
  /** 인쇄 모양 (pages) or 웹 모양 (text wraps to the window, no pages). */
  mode: 'print' | 'web' = 'print';

  paginate() {
    const doc = this.hooks.doc();
    if (!doc) return;
    if (this.mode === 'web') { this.layoutWeb(); return; }
    this.host.classList.remove('web-mode');
    const g = this.geometry();
    this.root.style.left = g.left + 'pt';
    this.root.style.top = g.top + 'pt';
    this.root.style.width = g.CW + 'pt';
    this.host.style.width = g.W + 'pt';
    this.root.classList.toggle('show-marks', this.showMarks);
    const blocks = Array.from(this.root.children).filter((c): c is HTMLElement => c instanceof HTMLElement && !c.hidden);
    for (const b of blocks) b.style.removeProperty('--push');
    this.root.querySelectorAll(':scope > table > tbody > tr.pg-spacer').forEach((r) => r.remove());
    const stride = (g.H + this.gapPt) * PX; // px from one page's content top to the next
    const content = (g.H - g.top - g.bottom) * PX;
    const measured = blocks.map((b) => {
      const h = b.offsetHeight;
      // A table only needs its first row to fit; later rows move to the next page one by one.
      const first = b instanceof HTMLTableElement ? b.rows[0] : null;
      return { el: b, top: b.offsetTop, h, fit: first ? first.offsetTop + first.offsetHeight : h, mt: parseFloat(getComputedStyle(b).marginTop) || 0 };
    });
    const out: BlockBox[] = [];
    let shift = 0, page = 0, breakNext = false;
    for (const m of measured) {
      let y = m.top + shift;
      const pageTop = page * stride;
      const atTop = y - m.mt <= pageTop + 0.5;
      const pushTo = (to: number) => {
        const delta = to * stride - (y - m.mt);
        if (delta > 0) { shift += delta; y += delta; m.el.style.setProperty('--push', delta + 'px'); }
        page = to;
      };
      const forced = breakNext || m.el.dataset.breakBefore === '1';
      breakNext = false;
      if (forced && !(atTop && y > 0)) pushTo(page + 1);
      else if (y + m.fit > pageTop + content && m.fit <= content && !atTop) {
        const prev = out[out.length - 1];
        if (prev && prev.el.dataset.keepNext === '1' && prev.page === page && prev.y - prev.mt > pageTop + 0.5 && prev.h + m.fit <= content) {
          // Keep a heading with the paragraph that follows it: move the heading instead.
          const delta = (page + 1) * stride - (prev.y - prev.mt);
          const cur = parseFloat(prev.el.style.getPropertyValue('--push')) || 0;
          prev.el.style.setProperty('--push', cur + delta + 'px');
          shift += delta;
          prev.y += delta;
          prev.page = prev.endPage = ++page;
          y = m.top + shift;
        } else pushTo(page + 1);
      }
      let h = m.h;
      if (m.el instanceof HTMLTableElement) {
        // Split the table between rows: a spacer row carries the rest of the page and the gap.
        let extra = 0, pg = page;
        const cols = m.el.querySelector('colgroup')?.children.length || 1;
        for (const tr of Array.from(m.el.rows).slice(1)) {
          // offsetTop already includes spacer rows inserted above this one.
          const rt = y + tr.offsetTop, rh = tr.offsetHeight;
          const bottom = pg * stride + content;
          if (rt + rh <= bottom) continue;
          if (rh > content || rt <= pg * stride + 0.5) { pg = Math.floor((rt + rh) / stride); continue; }
          const gap = (pg + 1) * stride - rt;
          const sp = document.createElement('tr');
          sp.className = 'pg-spacer';
          sp.contentEditable = 'false';
          const td = document.createElement('td');
          td.colSpan = cols;
          td.style.height = gap + 'px';
          sp.appendChild(td);
          tr.before(sp);
          pg++;
        }
        extra = m.el.offsetHeight - m.h;
        shift += extra;
        h += extra;
      }
      const endPage = Math.max(page, Math.floor((y + h - 1) / stride));
      out.push({ el: m.el, y, h, mt: m.mt, page, endPage });
      page = endPage;
      if (m.el.dataset.breakAfter === '1' || m.el.querySelector('.w-pgbr')) breakNext = true;
    }
    this.layout = out;
    this.pages = Math.max(1, (out.at(-1)?.endPage ?? 0) + 1 + (breakNext ? 1 : 0));
    this.updateFields(out, stride);
    this.buildFrames(g);
    this.applyZoom();
    this.hooks.layoutChanged();
  }

  /** Web layout: one white surface as wide as the window; tables keep their widths. */
  private layoutWeb() {
    if (this.hf) this.exitHF();
    const sc = this.wrap.closest('#scroll-container') as HTMLElement | null;
    const availPx = Math.max(320, (sc?.clientWidth ?? 900) - 4);
    const pad = 18; // pt
    const wPt = (availPx / this.zoom) * 0.75;
    this.host.classList.add('web-mode');
    this.frames.innerHTML = '';
    for (const b of Array.from(this.root.children) as HTMLElement[]) b.style.removeProperty('--push');
    this.root.querySelectorAll(':scope > table > tbody > tr.pg-spacer').forEach((r) => r.remove());
    this.root.classList.toggle('show-marks', this.showMarks);
    Object.assign(this.root.style, { left: pad + 'pt', top: pad + 'pt', width: wPt - 2 * pad + 'pt' });
    this.host.style.width = wPt + 'pt';
    const hPt = this.root.offsetHeight * 0.75 + 2 * pad;
    this.host.style.height = hPt + 'pt';
    this.layout = Array.from(this.root.children).filter((c): c is HTMLElement => c instanceof HTMLElement && !c.hidden)
      .map((el) => ({ el, y: el.offsetTop, h: el.offsetHeight, mt: 0, page: 0, endPage: 0 }));
    this.pages = 1;
    this.updateFields(this.layout, Number.MAX_SAFE_INTEGER);
    this.host.style.transform = `scale(${this.zoom})`;
    this.wrap.style.width = availPx + 'px';
    this.wrap.style.height = hPt * (4 / 3) * this.zoom + 'px';
    this.hooks.layoutChanged();
  }

  /** Switch between print and web layout, keeping the selection. */
  setMode(mode: 'print' | 'web') {
    if (mode === this.mode) return;
    const r = this.getRange();
    if (this.hf) this.exitHF();
    this.mode = mode;
    this.paginate();
    if (r) this.setRange(r);
  }

  /** PAGE / NUMPAGES fields in the body show the laid-out values. */
  private updateFields(boxes: BlockBox[], stride: number) {
    for (const b of boxes) {
      for (const f of Array.from(b.el.querySelectorAll<HTMLElement>('.w-fld[data-fld]'))) {
        const p = f.closest('p.w-p') as HTMLElement | null;
        const clean = p && this.reg.sig.get(p) === p.innerHTML;
        const top = b.y + (f.getBoundingClientRect().top - b.el.getBoundingClientRect().top) / this.zoom;
        const v = String(f.dataset.fld === 'PAGE' ? Math.floor(top / stride) + 1 : this.pages);
        if (f.textContent !== v) f.textContent = v;
        if (p && clean) this.reg.sig.set(p, p.innerHTML);
      }
    }
  }

  private buildFrames(g: ReturnType<DocEditor['geometry']>) {
    const doc = this.doc;
    // The header/footer being edited stays in place while the rest of the frames are rebuilt.
    const editing = this.hf?.el;
    editing?.remove();
    this.frames.innerHTML = '';
    for (let i = 0; i < this.pages; i++) {
      const f = document.createElement('div');
      f.className = 'page';
      f.style.top = i * (g.H + this.gapPt) + 'pt';
      f.style.width = g.W + 'pt';
      f.style.height = g.H + 'pt';
      const hf = doc.headerFooter(i + 1);
      if (hf.header) {
        const h = renderHeaderFooter(doc, this.hooks.images(), hf.header, i + 1, this.pages);
        h.className = 'page-header';
        Object.assign(h.style, { left: g.left + 'pt', width: g.CW + 'pt', top: g.ps.header / 20 + 'pt' });
        f.appendChild(h);
      }
      if (hf.footer) {
        const ft = renderHeaderFooter(doc, this.hooks.images(), hf.footer, i + 1, this.pages);
        ft.className = 'page-footer';
        Object.assign(ft.style, { left: g.left + 'pt', width: g.CW + 'pt', bottom: g.ps.footer / 20 + 'pt' });
        f.appendChild(ft);
      }
      // Margin corner marks, like Word's text boundaries.
      const marks = document.createElement('div');
      marks.className = 'margin-marks';
      Object.assign(marks.style, { left: g.left + 'pt', top: g.top + 'pt', width: g.CW + 'pt', height: g.H - g.top - g.bottom + 'pt' });
      f.appendChild(marks);
      this.frames.appendChild(f);
    }
    this.host.style.height = this.pages * (g.H + this.gapPt) - this.gapPt + 'pt';
    if (editing && this.hf) {
      const f = this.frames.children[Math.min(this.hf.page, this.pages) - 1] as HTMLElement | undefined;
      f?.querySelector(this.hf.kind === 'header' ? '.page-header' : '.page-footer')?.remove();
      f?.appendChild(editing);
    }
  }

  // ---------------------------------------------------------------- header / footer editing

  hf: HFEdit | null = null;

  /** Start editing the header or footer of page `page` (1-based), creating it if needed. */
  enterHF(kind: 'header' | 'footer', page: number) {
    const doc = this.doc;
    // Headers and footers only exist on pages: editing one switches to print layout, as in Word.
    if (this.mode === 'web') { this.mode = 'print'; this.hooks.layoutChanged(); }
    if (this.hf) this.exitHF();
    this.commit();
    const part = doc.openHeaderFooter(kind, page);
    this.paginate();
    const g = this.geometry();
    const el = document.createElement('div');
    el.className = (kind === 'header' ? 'page-header' : 'page-footer') + ' hf-editing';
    el.contentEditable = 'true';
    el.spellcheck = false;
    el.dataset.label = (kind === 'header' ? '머리글' : '바닥글') + ({ first: ' - 첫 페이지', even: ' - 짝수 페이지', default: doc.hfOptions().evenOdd ? ' - 홀수 페이지' : '' } as Record<string, string>)[doc.hfType(page)];
    Object.assign(el.style, { left: g.left + 'pt', width: g.CW + 'pt' });
    if (kind === 'header') el.style.top = g.ps.header / 20 + 'pt'; else el.style.bottom = g.ps.footer / 20 + 'pt';
    this.hf = { kind, part, page, el, reg: newRegistry(), pending: false };
    this.renderHF();
    el.addEventListener('input', () => { if (this.hf) this.hf.pending = true; });
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.exitHF(); }
      if (e.key === 'Tab') { e.preventDefault(); document.execCommand('insertText', false, '\t'); }
    });
    el.addEventListener('beforeinput', (e) => { if (e.inputType.startsWith('format')) e.preventDefault(); });
    this.host.classList.add('hf-mode');
    this.buildFrames(g);
    const last = el.querySelectorAll('p.w-p');
    const p = last[last.length - 1] as HTMLElement | undefined;
    el.focus({ preventScroll: true });
    if (p) { const [n, o] = posToDom(p, Number.MAX_SAFE_INTEGER); window.getSelection()!.setBaseAndExtent(n, o, n, o); }
    (this.frames.children[page - 1] as HTMLElement | undefined)?.scrollIntoView({ block: kind === 'header' ? 'start' : 'end' });
    this.hooks.hfChanged?.();
  }

  private renderHF() {
    const hf = this.hf!;
    const root = this.doc.pkg.getXml(hf.part)!.documentElement;
    hf.el.innerHTML = '';
    hf.reg = newRegistry();
    renderBlocks(root, { doc: this.doc, images: this.hooks.images(), part: hf.part, reg: hf.reg, editable: true, page: hf.page, pages: this.pages }, hf.el);
    hf.pending = false;
  }

  /** Write the edited header/footer to its part. */
  commitHF() {
    const hf = this.hf;
    if (!hf || !hf.pending) return;
    hf.pending = false;
    const root = this.doc.pkg.getXml(hf.part)!.documentElement;
    let pairs: [HTMLElement, Element][] = [];
    this.doc.transact(hf.kind === 'header' ? '머리글 편집' : '바닥글 편집', [hf.part], () => { ({ pairs } = commitContainer(hf.el, root, this.doc, hf.reg)); });
    for (const [el, x] of pairs) { hf.reg.dom.set(el, x); hf.reg.sig.set(el, el.innerHTML); }
  }

  exitHF() {
    if (!this.hf) return;
    this.commitHF();
    this.hf = null;
    this.host.classList.remove('hf-mode');
    this.paginate();
    this.focus();
    this.hooks.hfChanged?.();
  }

  /** Selected paragraph segments inside the header/footer being edited. */
  private hfSegments(): { p: Element; s: number; e: number; el: HTMLElement }[] {
    const hf = this.hf!;
    const sel = window.getSelection();
    if (!sel?.rangeCount || !hf.el.contains(sel.anchorNode)) return [];
    const r = sel.getRangeAt(0);
    const out: { p: Element; s: number; e: number; el: HTMLElement }[] = [];
    for (const pel of Array.from(hf.el.querySelectorAll<HTMLElement>('p.w-p'))) {
      if (!r.intersectsNode(pel)) continue;
      const x = hf.reg.dom.get(pel);
      if (!x) continue;
      const s = pel.contains(r.startContainer) ? measure(pel, r.startContainer, r.startOffset) : 0;
      const e = pel.contains(r.endContainer) ? measure(pel, r.endContainer, r.endOffset) : this.doc.lengthOf(x);
      out.push({ p: x, s, e, el: pel });
    }
    return out;
  }

  /** Character or paragraph formatting inside the header/footer (whole paragraph when nothing is selected). */
  hfFormat(op: { run?: RunProps; jc?: string }) {
    const hf = this.hf;
    if (!hf) return;
    this.commitHF();
    const segs = this.hfSegments();
    if (!segs.length) return;
    const idx = Array.from(hf.el.querySelectorAll('p.w-p')).indexOf(segs[0].el);
    const keep = { i: idx, s: segs[0].s, e: segs[segs.length - 1].e, last: Array.from(hf.el.querySelectorAll('p.w-p')).indexOf(segs[segs.length - 1].el) };
    this.doc.formatSegments(hf.part, segs.map((x) => ({ p: x.p, s: x.s, e: x.e === x.s && op.run ? -1 : x.e })), op);
    this.renderHF();
    const ps = hf.el.querySelectorAll<HTMLElement>('p.w-p');
    const a = ps[keep.i], b = ps[keep.last];
    if (a && b) { const [an, ao] = posToDom(a, keep.s); const [bn, bo] = posToDom(b, keep.e); window.getSelection()!.setBaseAndExtent(an, ao, bn, bo); }
  }

  /** Character formatting of the run at the caret inside the header/footer. */
  hfCaret(): { p: Element; rPr: Element | null } | null {
    const hf = this.hf;
    const sel = window.getSelection();
    if (!hf || !sel?.anchorNode || !hf.el.contains(sel.anchorNode)) return null;
    const node = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode as HTMLElement;
    const pel = node?.closest('p.w-p') as HTMLElement | null;
    const p = pel ? hf.reg.dom.get(pel) : null;
    if (!p) return null;
    const span = node?.closest('[data-r]') as HTMLElement | null;
    return { p, rPr: span ? hf.reg.runs[Number(span.dataset.r)] ?? null : null };
  }

  /** Insert a PAGE / NUMPAGES field at the caret in the header/footer. */
  hfInsertField(kw: 'PAGE' | 'NUMPAGES') {
    const hf = this.hf;
    const sel = window.getSelection();
    if (!hf || !sel?.rangeCount || !hf.el.contains(sel.anchorNode)) return;
    const xml = this.doc.pkg.getXml(hf.part)!;
    const fld = xml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:fldSimple');
    fld.setAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:instr', ` ${kw} `);
    const r = xml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
    const t = xml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
    t.textContent = '1';
    r.appendChild(t);
    fld.appendChild(r);
    const span = document.createElement('span');
    span.className = 'w-fld';
    span.contentEditable = 'false';
    span.dataset.a = String(hf.reg.atoms.push([fld]) - 1);
    span.textContent = String(kw === 'PAGE' ? hf.page : this.pages);
    const range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(span);
    range.setStartAfter(span);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    hf.pending = true;
  }

  setZoom(z: number) {
    this.zoom = Math.max(0.2, Math.min(5, z));
    if (this.mode === 'web') this.layoutWeb(); else this.applyZoom();
  }

  private applyZoom() {
    const doc = this.hooks.doc();
    if (!doc || this.mode === 'web') return;
    const g = this.geometry();
    this.host.style.transform = `scale(${this.zoom})`;
    this.wrap.style.width = g.W * PX * this.zoom + 'px';
    this.wrap.style.height = (this.pages * (g.H + this.gapPt) - this.gapPt) * PX * this.zoom + 'px';
  }

  /** 1-based page of the caret. */
  caretPage(): number {
    if (this.mode === 'web') return 1;
    const sel = window.getSelection();
    const node = sel?.focusNode;
    if (!node || !this.root.contains(node)) return 1;
    let el: Node | null = node;
    while (el && el.parentNode !== this.root) el = el.parentNode;
    const box = this.layout.find((b) => b.el === el);
    if (!box) return 1;
    const g = this.geometry();
    const stride = (g.H + this.gapPt) * PX;
    // Inside long blocks use the caret rectangle.
    const r = sel!.rangeCount ? sel!.getRangeAt(0).getBoundingClientRect() : null;
    if (r && r.height && box.endPage > box.page) {
      const rootTop = this.root.getBoundingClientRect().top;
      return Math.floor((r.top - rootTop) / this.zoom / stride) + 1;
    }
    return box.page + 1;
  }

  scrollToPara(pi: number) {
    const p = this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${pi}"]`);
    if (!p) return;
    p.scrollIntoView({ block: 'center' });
    this.setRange({ start: { pi, off: 0 }, end: { pi, off: 0 } });
  }

  // ---------------------------------------------------------------- selection

  getRange(): Range | null {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !this.root.contains(sel.anchorNode)) return this.lastRange;
    const a = domToPos(this.root, sel.anchorNode!, sel.anchorOffset);
    const f = domToPos(this.root, sel.focusNode!, sel.focusOffset);
    if (!a || !f) return this.lastRange;
    return { start: a, end: f };
  }

  setRange(r: Range | null, scroll = true) {
    if (!r) return;
    const pa = this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${r.start.pi}"]`);
    const pf = this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${r.end.pi}"]`);
    if (!pa || !pf) return;
    const [an, ao] = posToDom(pa, r.start.off);
    const [fn, fo] = posToDom(pf, r.end.off);
    const sel = window.getSelection()!;
    this.root.focus({ preventScroll: true });
    sel.setBaseAndExtent(an, ao, fn, fo);
    this.lastRange = r;
    if (scroll) {
      const el = fn.nodeType === 1 ? fn as HTMLElement : fn.parentElement;
      el?.scrollIntoView({ block: 'nearest' });
    }
  }

  collapsed() {
    const r = this.getRange();
    return !r || comparePos(r.start, r.end) === 0;
  }

  focus() { this.root.focus({ preventScroll: true }); }

  /** Remember the selection before focus moves to a toolbar input. */
  rememberSelection() { this.lastRange = this.getRange(); }

  // ---------------------------------------------------------------- commit

  /** Write pending DOM edits to the model. Returns true when something changed. */
  commit(): boolean {
    clearTimeout(this.commitTimer);
    const doc = this.hooks.doc();
    if (!doc || !this.pending) return false;
    this.pending = false;
    let changed = 0;
    let pairs: [HTMLElement, Element][] = [];
    this.committing = true;
    try {
      doc.transact('입력', [], () => {
        ({ pairs, changed } = commitDom(this.root, doc, this.reg));
      });
    } finally {
      this.committing = false;
    }
    for (const [el, x] of pairs) { this.reg.dom.set(el, x); this.reg.sig.set(el, el.innerHTML); }
    this.reindex();
    return changed > 0;
  }

  committing = false;

  private scheduleCommit() {
    clearTimeout(this.commitTimer);
    this.commitTimer = window.setTimeout(() => { if (!this.composing) { this.commit(); this.hooks.selectionChanged(); } }, 900);
  }

  // ---------------------------------------------------------------- formatting helpers

  /**
   * Apply character formatting. A collapsed caret inside a word formats the word (as Word does);
   * elsewhere it arms "pending" formatting for the next typed text.
   */
  applyRunProps(props: RunProps, label?: string): boolean {
    const doc = this.doc;
    this.commit();
    let r = this.getRange();
    if (!r) return false;
    r = normRange(r);
    if (comparePos(r.start, r.end) === 0) {
      const w = doc.wordAt(r.start);
      if (!w) { this.armPending(r.start, props); return true; }
      doc.applyRunProps(w, props, label);
      this.refresh(w.start.pi, 1, 1);
      this.setRange(r, false);
      return true;
    }
    doc.applyRunProps(r, props, label);
    this.refresh(r.start.pi, r.end.pi - r.start.pi + 1, r.end.pi - r.start.pi + 1);
    this.setRange(r, false);
    return true;
  }

  /** Insert an empty styled span at the caret so the next typed characters get `props`. */
  private armPending(pos: Pos, props: RunProps) {
    const doc = this.doc;
    const p = doc.para(pos.pi);
    if (!p) return;
    const prev = this.pendingFmt && comparePos(this.pendingFmt.pos, pos) === 0 ? this.pendingFmt.span : null;
    const base = prev ? this.reg.runs[Number(prev.dataset.r)] : doc.rPrAt(p, pos.off);
    const rPr = base ? base.cloneNode(true) as Element : create(doc.xml, 'w:rPr');
    setRunProps(rPr, props);
    if (prev) {
      this.reg.runs[Number(prev.dataset.r)] = rPr;
      runCss(prev, doc.runFmt(p, rPr));
      this.placeCaretIn(prev);
      this.hooks.selectionChanged();
      return;
    }
    const span = document.createElement('span');
    span.dataset.r = String(this.reg.runs.push(rPr) - 1);
    runCss(span, doc.runFmt(p, rPr));
    span.textContent = '​';
    const pel = this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${pos.pi}"]`);
    if (!pel) return;
    const [node, off] = posToDom(pel, pos.off);
    const range = document.createRange();
    range.setStart(node, off);
    range.collapse(true);
    range.insertNode(span);
    this.pendingFmt = { pos, span };
    this.pending = true;
    this.placeCaretIn(span);
    this.hooks.selectionChanged();
  }

  private placeCaretIn(span: HTMLElement) {
    const t = span.firstChild!;
    this.root.focus({ preventScroll: true });
    window.getSelection()!.setBaseAndExtent(t, t.textContent!.length, t, t.textContent!.length);
  }

  /** Run properties of the pending span, when the caret sits in it. */
  pendingRPr(): Element | null | undefined {
    const span = this.pendingFmt?.span;
    if (!span?.isConnected) return undefined;
    const sel = window.getSelection();
    if (!sel?.focusNode || !span.contains(sel.focusNode)) return undefined;
    return this.reg.runs[Number(span.dataset.r)];
  }

  /** Paragraph indices covered by the selection. */
  paraIndices(): number[] {
    const r = this.getRange();
    if (!r) return [];
    return this.doc.paraRange(r);
  }

  /** Run a paragraph-level command on the selection and refresh. */
  paraCommand(fn: (pis: number[]) => void) {
    this.commit();
    const r = this.getRange();
    if (!r) return;
    const pis = this.doc.paraRange(r);
    fn(pis);
    const before = pis.length;
    this.refresh(pis[0], before, before);
    this.setRange(r, false);
  }

  // ---------------------------------------------------------------- events

  private bind() {
    const root = this.root;
    root.addEventListener('beforeinput', (e) => this.onBeforeInput(e));
    root.addEventListener('input', (e) => {
      this.pending = true;
      // Safety net: a paragraph split the browser did natively is committed at once so that
      // paragraph indices never drift (normally Enter is handled in onBeforeInput).
      if ((e as InputEvent).inputType === 'insertParagraph' && !this.composing) { const r = this.getRange(); this.commit(); this.render(); if (r) this.setRange({ start: { pi: r.start.pi + 1, off: 0 }, end: { pi: r.start.pi + 1, off: 0 } }); return; }
      this.scheduleCommit();
      this.layoutSoon();
    });
    root.addEventListener('compositionstart', () => { this.composing = true; });
    root.addEventListener('compositionend', () => { this.composing = false; this.pending = true; this.scheduleCommit(); });
    root.addEventListener('blur', () => { if (!this.composing) this.commit(); });
    root.addEventListener('keydown', (e) => this.onKeyDown(e));
    root.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.host.addEventListener('mousemove', (e) => this.onHover(e));
    // A table's outer right border can be grabbed from the page margin too (outside #doc-root).
    this.host.addEventListener('mousedown', (e) => {
      if (this.root.contains(e.target as Node) || e.button !== 0) return;
      const edge = this.tableEdge(e);
      if (edge) this.startTableDrag(e, edge);
    });
    this.host.addEventListener('mousemove', (e) => this.updateTableHandle(e));
    this.host.addEventListener('mouseleave', () => this.hideTableHandle());
    // Word: double-click the top or bottom margin of a page to edit its header or footer.
    this.host.addEventListener('dblclick', (e) => {
      const doc = this.hooks.doc();
      if (!doc) return;
      const t = e.target as HTMLElement;
      if (this.hf && (this.hf.el.contains(t))) return;
      const page = t.closest('.page') as HTMLElement | null;
      if (!page) { if (this.hf && this.root.contains(t)) this.exitHF(); return; }
      const g = this.geometry();
      const r = page.getBoundingClientRect();
      const y = (e.clientY - r.top) / this.zoom / (4 / 3); // pt from the page top
      if (y < g.top) this.enterHF('header', Array.from(this.frames.children).indexOf(page) + 1);
      else if (y > g.H - g.bottom) this.enterHF('footer', Array.from(this.frames.children).indexOf(page) + 1);
    });
    this.root.addEventListener('dblclick', () => { if (this.hf) this.exitHF(); });
    root.addEventListener('click', (e) => {
      const a = (e.target as HTMLElement).closest('a.w-hyperlink') as HTMLElement | null;
      if (a?.dataset.href && (e.ctrlKey || e.metaKey)) window.open(a.dataset.href, '_blank');
    });
    root.addEventListener('dragstart', (e) => { if ((e.target as HTMLElement).tagName === 'IMG') e.preventDefault(); });
    document.addEventListener('selectionchange', () => {
      const sel = window.getSelection();
      if (sel?.anchorNode && root.contains(sel.anchorNode)) {
        this.lastRange = this.getRange();
        this.hooks.selectionChanged();
      }
    });
  }

  private onBeforeInput(e: InputEvent) {
    const t = e.inputType;
    if (t.startsWith('format') || t === 'insertFromDrop' || t === 'deleteByDrag') { e.preventDefault(); return; }
    if (t === 'insertParagraph') {
      e.preventDefault();
      // During IME composition the composed text must land before splitting.
      if (this.composing || e.isComposing) setTimeout(() => this.enter(), 0);
      else this.enter();
      return;
    }
  }

  /** Enter: split the paragraph in the model. */
  enter() {
    const doc = this.hooks.doc();
    if (!doc) return;
    if (!this.collapsed()) { document.execCommand('delete'); }
    this.commit();
    const r = this.getRange();
    if (!r) return;
    const pos = r.end;
    const n = doc.paragraphs().length;
    const next = doc.splitAt(pos);
    const added = doc.paragraphs().length - n;
    this.refresh(pos.pi, 1, 1 + added);
    this.setRange({ start: next, end: next });
  }

  private onKeyDown(e: KeyboardEvent) {
    if (e.isComposing || e.keyCode === 229) return;
    const doc = this.hooks.doc();
    if (!doc) return;
    if (this.selImg && (e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      const atom = this.selImg.closest('[data-a]');
      this.deselectImage();
      atom?.remove();
      this.pending = true;
      this.commit();
      this.render();
      return;
    }
    if (this.selImg) this.deselectImage();
    if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      this.onTab(e.shiftKey);
      return;
    }
    if (e.key === 'Backspace' && !e.shiftKey && this.collapsed()) {
      const r = this.getRange();
      const p = r && doc.para(r.start.pi);
      if (r && p && r.start.off === 0 && doc.labels().get(p)) {
        e.preventDefault();
        this.commit();
        const f = doc.paraFmt(p);
        doc.toggleList([r.start.pi], doc.numbering.isBullet(f.numId!, f.ilvl ?? 0) ? 'bullet' : 'number');
        this.refresh(r.start.pi, 1, 1);
        this.setRange(r);
      }
    }
  }

  private onTab(shift: boolean) {
    const doc = this.doc;
    this.commit();
    const r = this.getRange();
    if (!r) return;
    const p = doc.para(r.start.pi)!;
    const cell = doc.cellOf(r.start.pi);
    if (cell) {
      // Move between cells.
      const all = Array.from(this.root.querySelectorAll<HTMLElement>('td'));
      const td = this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${r.start.pi}"]`)?.closest('td');
      const i = td ? all.indexOf(td as HTMLElement) : -1;
      const target = all[i + (shift ? -1 : 1)];
      if (target) {
        const first = target.querySelector<HTMLElement>('p.w-p[data-pi]');
        if (first) {
          const pi = Number(first.dataset.pi);
          const len = doc.lengthOf(doc.para(pi)!);
          this.setRange({ start: { pi, off: 0 }, end: { pi, off: shift ? 0 : len } });
        }
      }
      return;
    }
    if (doc.labels().get(p) && r.start.off === 0 && comparePos(r.start, r.end) === 0) {
      doc.indent([r.start.pi], shift ? -1 : 1);
      this.refresh(r.start.pi, 1, 1);
      this.setRange(r);
      return;
    }
    if (shift) return;
    if (comparePos(r.start, r.end) !== 0) { document.execCommand('delete'); this.commit(); }
    const at = this.getRange()!.start;
    const next = doc.insertTab(at);
    this.refresh(at.pi, 1, 1);
    this.setRange({ start: next, end: next });
  }

  // ---------------------------------------------------------------- images

  // ---------------------------------------------------------------- table border dragging

  /** The table border under the pointer: a column boundary or a row bottom. */
  private tableEdge(e: MouseEvent): { kind: 'col' | 'row'; table: HTMLTableElement; index: number; pos: number } | null {
    const tol = 4;
    const cellAt = (x: number, y: number) => (document.elementFromPoint(x, y) as HTMLElement | null)?.closest?.('td') as HTMLTableCellElement | null;
    // Just outside the table's right or bottom border the pointer is not over a cell: look inward.
    let td = (e.target as HTMLElement).closest?.('td') as HTMLTableCellElement | null;
    if (!td) td = cellAt(e.clientX - tol - 1, e.clientY) ?? cellAt(e.clientX, e.clientY - tol - 1);
    if (!td || !this.root.contains(td) || td.dataset.col == null) return null;
    const table = td.closest('table') as HTMLTableElement;
    if (!this.reg.dom.get(table)) return null;
    const r = td.getBoundingClientRect();
    const col = Number(td.dataset.col);
    if (Math.abs(e.clientX - r.right) <= tol) return { kind: 'col', table, index: col + td.colSpan, pos: r.right };
    if (Math.abs(e.clientX - r.left) <= tol && col > 0) return { kind: 'col', table, index: col, pos: r.left };
    if (Math.abs(e.clientY - r.bottom) <= tol) return { kind: 'row', table, index: Number(td.dataset.row) + td.rowSpan - 1, pos: r.bottom };
    return null;
  }

  /** Current column widths in twips: the grid, or measured cells when the grid has none. */
  private measuredWidths(table: HTMLTableElement, tbl: Element): number[] {
    const grid = this.doc.columnWidths(tbl);
    if (grid.length && grid.every((w) => w > 0)) return grid;
    const n = Math.max(grid.length, table.querySelector('colgroup')?.children.length ?? 0);
    const out = new Array(n).fill(0);
    for (const td of Array.from(table.querySelectorAll<HTMLTableCellElement>(':scope > tbody > tr > td[data-col]'))) {
      if (td.colSpan === 1) out[Number(td.dataset.col)] = Math.max(out[Number(td.dataset.col)], td.offsetWidth * 15); // px → twips
    }
    const avg = out.filter((w) => w > 0).reduce((a, b, _i, arr) => a + b / arr.length, 0) || 1440;
    return out.map((w) => w || avg);
  }

  private onHover(e: MouseEvent) {
    if (e.buttons) return;
    const edge = this.tableEdge(e);
    this.host.style.cursor = this.root.style.cursor = edge ? (edge.kind === 'col' ? 'col-resize' : 'row-resize') : '';
  }

  // ---------------------------------------------------------------- whole-table resize handle

  private tblHandle: HTMLElement | null = null;
  private tblHandleFor: HTMLTableElement | null = null;

  /** Word: hovering a table shows a small square at its bottom-right corner; dragging it resizes the table. */
  private updateTableHandle(e: MouseEvent) {
    if (e.buttons || this.hf) return;
    const t = e.target as HTMLElement;
    if (t === this.tblHandle) return;
    let table = t.closest?.('table.w-tbl') as HTMLTableElement | null;
    if (table && (!this.root.contains(table) || !this.reg.dom.get(table))) table = null;
    if (!table && this.tblHandleFor) {
      // Keep the handle while the pointer is on its way to it, just outside the corner.
      const r = this.tblHandleFor.getBoundingClientRect();
      const pad = 16 * this.zoom;
      if (this.tblHandleFor.isConnected && e.clientX >= r.left && e.clientX <= r.right + pad && e.clientY >= r.top && e.clientY <= r.bottom + pad) return;
    }
    if (!table) { this.hideTableHandle(); return; }
    this.showTableHandle(table);
  }

  private showTableHandle(table: HTMLTableElement) {
    if (!this.tblHandle) {
      const h = document.createElement('div');
      h.className = 'tbl-handle';
      h.title = '끌어서 표 크기 조절';
      h.addEventListener('mousedown', (e) => { if (e.button === 0) this.startTableResize(e); });
      this.tblHandle = h;
    }
    this.tblHandleFor = table;
    const hr = this.host.getBoundingClientRect();
    const r = table.getBoundingClientRect();
    Object.assign(this.tblHandle.style, { left: (r.right - hr.left) / this.zoom + 1 + 'px', top: (r.bottom - hr.top) / this.zoom + 1 + 'px' });
    if (this.tblHandle.parentNode !== this.host) this.host.appendChild(this.tblHandle);
  }

  hideTableHandle() {
    this.tblHandle?.remove();
    this.tblHandleFor = null;
  }

  private startTableResize(e: MouseEvent) {
    const table = this.tblHandleFor;
    const tbl = table && this.reg.dom.get(table);
    if (!table || !tbl) return;
    e.preventDefault();
    e.stopPropagation();
    const r0 = table.getBoundingClientRect();
    const rows = Array.from(table.rows).filter((r) => !r.classList.contains('pg-spacer') && r.closest('table') === table);
    const rowPx = rows.map((r) => r.getBoundingClientRect().height);
    const h0 = rowPx.reduce((a, b) => a + b, 0) || r0.height;
    const w0 = r0.width;
    const widths = this.measuredWidths(table, tbl);
    const guide = document.createElement('div');
    guide.className = 'tbl-guide box';
    Object.assign(guide.style, { left: r0.left + 'px', top: r0.top + 'px', width: w0 + 'px', height: r0.height + 'px' });
    document.body.appendChild(guide);
    const minW = widths.length * 240 * ((this.zoom * 4) / 3 / 20); // ~0.4 cm per column
    const minH = rows.length * 8 * this.zoom;
    const x0 = e.clientX, y0 = e.clientY;
    let w = w0, h = h0;
    const move = (ev: MouseEvent) => {
      w = Math.max(minW, w0 + ev.clientX - x0);
      h = Math.max(minH, h0 + ev.clientY - y0);
      if (ev.shiftKey) { const f = Math.max(w / w0, h / h0); w = w0 * f; h = h0 * f; } // keep proportions
      guide.style.width = w + 'px';
      guide.style.height = r0.height - h0 + h + 'px';
      const cm = (px: number) => (px / this.zoom * 0.75 / 28.3465).toFixed(2);
      this.hooks.status(`표 크기 ${cm(w)} × ${cm(h)} cm — Shift: 비율 유지`);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide.remove();
      const fx = w / w0, fy = h / h0;
      const sameW = Math.abs(w - w0) < 1, sameH = Math.abs(h - h0) < 1;
      if (sameW && sameH) return;
      const r = this.getRange();
      this.commit();
      const ws = sameW ? widths : widths.map((x) => x * fx);
      const hs = sameH ? null : rowPx.map((px) => (px * fy / this.zoom) * 0.75 * 20);
      this.doc.resizeTable(tbl, ws, hs);
      this.hideTableHandle();
      this.render();
      if (r) this.setRange(r, false);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  private startTableDrag(e: MouseEvent, edge: NonNullable<ReturnType<DocEditor['tableEdge']>>) {
    e.preventDefault();
    const tbl = this.reg.dom.get(edge.table)!;
    const guide = document.createElement('div');
    guide.className = 'tbl-guide ' + edge.kind;
    const tr = edge.table.getBoundingClientRect();
    const sc = this.wrap.closest('#scroll-container')!.getBoundingClientRect();
    if (edge.kind === 'col') Object.assign(guide.style, { left: edge.pos + 'px', top: sc.top + 'px', height: sc.height + 'px' });
    else Object.assign(guide.style, { top: edge.pos + 'px', left: Math.min(tr.left, sc.left) + 'px', width: Math.max(tr.width, sc.width) + 'px' });
    document.body.appendChild(guide);
    const start = edge.kind === 'col' ? e.clientX : e.clientY;
    const widths = edge.kind === 'col' ? this.measuredWidths(edge.table, tbl) : [];
    const k = edge.index;
    // Limits so neither neighbouring column gets narrower than ~0.4 cm.
    const minTw = 240, pxPerTw = (this.zoom * 4) / 3 / 20;
    const lo = k > 0 ? -(widths[k - 1] - minTw) * pxPerTw : 0;
    const hi = k < widths.length ? (widths[k] - minTw) * pxPerTw : Infinity;
    const rowEl = edge.kind === 'row' ? Array.from(edge.table.rows).filter((r) => !r.classList.contains('pg-spacer'))[k] : null;
    const rowH = rowEl ? rowEl.getBoundingClientRect().height : 0;
    let delta = 0;
    const move = (ev: MouseEvent) => {
      if (edge.kind === 'col') {
        delta = Math.max(lo, Math.min(hi, ev.clientX - start));
        guide.style.left = edge.pos + delta + 'px';
      } else {
        delta = Math.max(-rowH + 8 * this.zoom, ev.clientY - start);
        guide.style.top = edge.pos + delta + 'px';
      }
      this.hooks.status(edge.kind === 'col' ? `열 너비 ${(delta / pxPerTw / 566.929 >= 0 ? '+' : '')}${(delta / pxPerTw / 566.929).toFixed(2)} cm` : `행 높이 ${((rowH + delta) / this.zoom * 0.75 / 28.3465).toFixed(2)} cm`);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide.remove();
      if (Math.abs(delta) < 1) return;
      const r = this.getRange();
      this.commit();
      if (edge.kind === 'col') {
        const tw = delta / pxPerTw;
        const ws = [...widths];
        if (k > 0) ws[k - 1] += tw;
        if (k < ws.length) ws[k] -= tw; // the right-hand neighbour gives up the space; the table keeps its width
        this.doc.setColumnWidths(tbl, ws);
      } else {
        const twips = ((rowH + delta) / this.zoom) * 0.75 * 20;
        this.doc.setRowHeight(tbl, k, twips);
      }
      this.render();
      if (r) this.setRange(r, false);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  private onMouseDown(e: MouseEvent) {
    const edge = this.tableEdge(e);
    if (edge && e.button === 0) { this.startTableDrag(e, edge); return; }
    const t = e.target as HTMLElement;
    if (t.classList.contains('img-handle')) { this.startResize(e); return; }
    if (t.tagName === 'IMG' && t.closest('#doc-root')) {
      e.preventDefault();
      this.selectImage(t);
      return;
    }
    if (this.selImg) this.deselectImage();
  }

  private selectImage(img: HTMLElement) {
    this.deselectImage();
    this.selImg = img;
    img.classList.add('selected');
    const atom = img.closest('[data-a]') as HTMLElement;
    const handle = document.createElement('span');
    handle.className = 'img-handle';
    handle.contentEditable = 'false';
    atom.appendChild(handle);
    // Put the caret just before the image so typing continues from there.
    const sel = window.getSelection()!;
    const parent = atom.parentNode!;
    sel.setBaseAndExtent(parent, Array.prototype.indexOf.call(parent.childNodes, atom), parent, Array.prototype.indexOf.call(parent.childNodes, atom));
    this.hooks.status(`그림 ${Math.round(img.offsetWidth * 0.75)} × ${Math.round(img.offsetHeight * 0.75)} pt — 모서리를 끌어서 크기 조절, Delete로 지우기`);
  }

  deselectImage() {
    if (!this.selImg) return;
    this.selImg.classList.remove('selected');
    this.selImg.closest('[data-a]')?.querySelector('.img-handle')?.remove();
    this.selImg = null;
  }

  private startResize(e: MouseEvent) {
    e.preventDefault();
    const img = this.selImg;
    if (!img) return;
    const atom = img.closest('[data-a]') as HTMLElement;
    const p = paraElOf(atom, this.root);
    if (!p) return;
    const parent = atom.parentNode!;
    const pos = domToPos(this.root, parent, Array.prototype.indexOf.call(parent.childNodes, atom));
    const w0 = img.offsetWidth, h0 = img.offsetHeight, x0 = e.clientX;
    const g = this.geometry();
    const move = (ev: MouseEvent) => {
      const w = Math.max(16, Math.min(g.CW * PX, w0 + (ev.clientX - x0) / this.zoom));
      img.style.width = w + 'px';
      img.style.height = (w * h0) / w0 + 'px';
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (!pos) return;
      this.commit();
      const wPt = img.offsetWidth * 0.75, hPt = img.offsetHeight * 0.75;
      this.doc.resizeImage(pos, wPt, hPt);
      this.refresh(pos.pi, 1, 1);
      this.setRange({ start: pos, end: pos });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  // ---------------------------------------------------------------- misc

  /** Highlight a find hit. */
  selectHit(h: { pi: number; off: number; len: number }) {
    this.setRange({ start: { pi: h.pi, off: h.off }, end: { pi: h.pi, off: h.off + h.len } });
  }

  isEditorTarget(t: EventTarget | null) {
    return t instanceof Node && this.root.contains(t);
  }

  paragraphElement(pi: number) {
    return this.root.querySelector<HTMLElement>(`p.w-p[data-pi="${pi}"]`);
  }

  isParaEl = isPara;
}
