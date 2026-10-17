// Document renderer: WordprocessingML → editable HTML (1 CSS pt = 20 twips = 12700 EMU).
// Like HOP's renderPage* layer it only reads the model. Every rendered paragraph, run and inline
// object is registered so the editor can map DOM positions and edits back to XML (see edit.ts).
import { DocxDocument, textDir } from './document';
import { Item } from './model';
import { REL } from './package';
import { bulletChar } from './numbering';
import { Border, RunFmt, parseBorders } from './styles';
import { NS, child, children, descendants, getRId, is, num, path, wa, wnum } from './xml';

export class ImageCache {
  private urls = new Map<string, string>();
  constructor(private doc: DocxDocument) {}

  url(part: string, rId: string | null): string | null {
    if (!rId) return null;
    const t = this.doc.pkg.rel(part, rId)?.target;
    return t ? this.urls.get(t) ?? null : null;
  }

  /** Load every image referenced by the given parts (document, headers, footers). */
  async ensure(parts: string[]) {
    const targets = parts.flatMap((p) => this.doc.imageTargets(p).map((r) => r.target));
    await Promise.all(targets.filter((t) => !this.urls.has(t)).map(async (t) => {
      const ext = t.split('.').pop()!.toLowerCase();
      if (['emf', 'wmf', 'tif', 'tiff'].includes(ext)) return;
      const data = await this.doc.pkg.getBinary(t);
      if (!data) return;
      const mime = ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
      this.urls.set(t, URL.createObjectURL(new Blob([data as BlobPart], { type: mime })));
    }));
  }

  /** Document plus every header/footer part. */
  async ensureAll() {
    const parts = [this.doc.docPath, ...this.doc.pkg.rels(this.doc.docPath).filter((r) => r.type === REL.header || r.type === REL.footer).map((r) => r.target)];
    await this.ensure(parts);
  }

  dispose() {
    for (const u of this.urls.values()) URL.revokeObjectURL(u);
    this.urls.clear();
  }
}

/** Links between the DOM and the XML it was rendered from. */
export interface Registry {
  runs: (Element | null)[]; // data-r → direct w:rPr of the run
  atoms: Element[][]; // data-a → paragraph-level nodes of an inline object, tab, break, field…
  containers: Element[]; // data-c → hyperlink / ins / smartTag element
  dom: WeakMap<Element, Element>; // rendered p / table / tr / td / block → XML element
  sig: WeakMap<Element, string>; // paragraph innerHTML as rendered (change detection)
}

export function newRegistry(): Registry {
  return { runs: [], atoms: [], containers: [], dom: new WeakMap(), sig: new WeakMap() };
}

export interface RenderCtx {
  doc: DocxDocument;
  images: ImageCache;
  part: string; // part whose relationships resolve r:id / r:embed
  reg: Registry;
  editable: boolean;
  page?: number;
  pages?: number;
  notes?: { n: number };
}

export const FALLBACK_FONTS = `'Apple SD Gothic Neo','Malgun Gothic','맑은 고딕','Noto Sans KR','Noto Sans CJK KR','Apple Color Emoji','Segoe UI Emoji','Noto Color Emoji',sans-serif`;

const FONT_ALIASES: Record<string, string> = {
  'Calibri': "'Calibri','Carlito'",
  'Cambria': "'Cambria','Caladea'",
  '맑은 고딕': "'맑은 고딕','Malgun Gothic','Apple SD Gothic Neo'",
  'Malgun Gothic': "'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo'",
  '바탕': "'바탕','Batang','AppleMyungjo'",
  'Batang': "'Batang','바탕','AppleMyungjo'",
  '굴림': "'굴림','Gulim','Apple SD Gothic Neo'",
  '돋움': "'돋움','Dotum','Apple SD Gothic Neo'",
  '궁서': "'궁서','Gungsuh','AppleMyungjo'",
};

export function fontStack(f: { ascii?: string; ea?: string }) {
  const names = [f.ascii, f.ea].filter((n, i, a): n is string => !!n && a.indexOf(n) === i);
  return [...names.map((n) => FONT_ALIASES[n] ?? `'${n.replace(/'/g, '')}'`), FALLBACK_FONTS].join(',');
}

export const pt = (twips: number) => `${+(twips / 20).toFixed(2)}pt`;

/** Apply resolved run formatting to an element (every property is set explicitly). */
export function runCss(el: HTMLElement, f: RunFmt) {
  const s = el.style;
  s.fontFamily = fontStack(f);
  let size = (f.sz ?? 20) / 2;
  s.fontWeight = f.b ? 'bold' : 'normal';
  s.fontStyle = f.i ? 'italic' : 'normal';
  const deco: string[] = [];
  if (f.u && f.u !== 'none') deco.push('underline');
  if (f.strike || f.dstrike) deco.push('line-through');
  s.textDecorationLine = deco.join(' ') || 'none';
  const u = f.u ?? '';
  s.textDecorationStyle = f.dstrike || u === 'double' ? 'double' : /dot/i.test(u) ? 'dotted' : /dash/i.test(u) ? 'dashed' : /wav/i.test(u) ? 'wavy' : 'solid';
  if (/thick|heavy/i.test(u)) s.textDecorationThickness = '2px';
  s.color = f.color && f.color !== 'auto' ? '#' + f.color : '#000000';
  const bg = f.highlight && f.highlight !== 'none' ? f.highlight : f.shd && f.shd !== 'auto' ? f.shd : null;
  s.backgroundColor = bg ? '#' + bg : 'transparent';
  if (f.vertAlign === 'superscript' || f.vertAlign === 'subscript') {
    s.verticalAlign = f.vertAlign === 'superscript' ? 'super' : 'sub';
    size *= 0.65;
  } else s.verticalAlign = 'baseline';
  s.fontSize = `${+size.toFixed(2)}pt`;
  s.textTransform = f.caps ? 'uppercase' : 'none';
  s.fontVariant = f.smallCaps ? 'small-caps' : 'normal';
  s.letterSpacing = f.spacing ? pt(f.spacing) : 'normal';
  if (f.vanish) s.display = 'none';
}

function cssBorder(b: Border | undefined): string {
  if (!b || b.style === 'none' || !b.sz) return 'none';
  const w = Math.max(0.5, b.sz / 8);
  const style = b.style === 'double' ? 'double' : /dot/i.test(b.style) ? 'dotted' : /dash/i.test(b.style) ? 'dashed' : 'solid';
  return `${b.style === 'double' ? Math.max(w * 3, 2.25) : w}pt ${style} #${b.color ?? '000000'}`;
}

const JC: Record<string, string> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', thaiDistribute: 'justify', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify' };

// ---------------------------------------------------------------- blocks

/** Render the block children of a body, table cell, header/footer or text box into `host`. */
export function renderBlocks(container: Element, ctx: RenderCtx, host: HTMLElement) {
  const kids = children(container);
  const styleOf = (el: Element | undefined) => el && is(el, 'w:p') ? ctx.doc.pStyleOf(el) ?? ctx.doc.styles.defaultPara : undefined;
  kids.forEach((c, i) => {
    if (c.namespaceURI !== NS.w) { host.appendChild(hiddenBlock(c, ctx)); return; }
    switch (c.localName) {
      case 'p': {
        const f = ctx.doc.paraFmt(c);
        const st = styleOf(c);
        // contextualSpacing: no space between paragraphs of the same style.
        const prev = kids[i - 1], next = kids[i + 1];
        host.appendChild(renderParagraph(c, ctx, {
          noBefore: !!f.contextual && styleOf(prev) === st,
          noAfter: !!f.contextual && styleOf(next) === st,
        }));
        break;
      }
      case 'tbl': host.appendChild(renderTable(c, ctx)); break;
      case 'sdt': host.appendChild(renderBlockSdt(c, ctx)); break;
      case 'sectPr': break;
      case 'customXml': {
        const div = document.createElement('div');
        div.className = 'w-block';
        div.contentEditable = 'false';
        renderBlocks(c, { ...ctx, editable: false }, div);
        if (ctx.editable) ctx.reg.dom.set(div, c);
        host.appendChild(div);
        break;
      }
      default: host.appendChild(hiddenBlock(c, ctx));
    }
  });
}

function hiddenBlock(c: Element, ctx: RenderCtx) {
  const d = document.createElement('div');
  d.hidden = true;
  d.className = 'w-hidden';
  d.contentEditable = 'false';
  if (ctx.editable) ctx.reg.dom.set(d, c);
  return d;
}

function renderBlockSdt(sdt: Element, ctx: RenderCtx) {
  const div = document.createElement('div');
  div.className = 'w-block w-sdt';
  div.contentEditable = 'false';
  const alias = wa(path(sdt, 'w:sdtPr', 'w:alias'), 'val');
  if (alias) div.title = alias;
  const content = child(sdt, 'w:sdtContent');
  if (content) renderBlocks(content, { ...ctx, editable: false }, div);
  if (ctx.editable) ctx.reg.dom.set(div, sdt);
  return div;
}

// ---------------------------------------------------------------- paragraphs

export function renderParagraph(p: Element, ctx: RenderCtx, opts: { noBefore?: boolean; noAfter?: boolean } = {}): HTMLElement {
  const { doc, reg } = ctx;
  const f = doc.paraFmt(p);
  const el = document.createElement('p');
  el.className = 'w-p';
  if (ctx.editable) {
    el.dataset.pi = String(doc.indexOf(p));
    reg.dom.set(el, p);
  }
  const s = el.style;
  s.textAlign = JC[f.jc ?? 'left'] ?? 'left';
  if (f.jc === 'distribute') s.textAlignLast = 'justify';
  const left = f.left ?? 0, right = f.right ?? 0;
  const hanging = f.hanging ?? 0, first = f.firstLine ?? 0;
  if (left) s.marginLeft = pt(left);
  if (right) s.marginRight = pt(right);
  if (hanging) s.textIndent = pt(-hanging); else if (first) s.textIndent = pt(first);
  const before = opts.noBefore ? 0 : f.before ?? 0;
  const after = opts.noAfter ? 0 : f.after ?? 0;
  s.setProperty('--mt', pt(before));
  s.marginBottom = pt(after);
  if (f.line != null) {
    if (!f.lineRule || f.lineRule === 'auto') s.lineHeight = String(+((f.line / 240) * 1.2).toFixed(3));
    else s.lineHeight = pt(f.line);
  }
  if (f.shd) s.backgroundColor = '#' + f.shd;
  if (f.bdr) {
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const b = f.bdr[side];
      if (!b || b.style === 'none') continue;
      s.setProperty(`border-${side}`, cssBorder(b));
      s.setProperty(`padding-${side}`, `${b.space}pt`);
    }
  }
  if (f.pageBreakBefore) el.dataset.breakBefore = '1';
  if (f.keepNext) el.dataset.keepNext = '1';
  if (path(p, 'w:pPr', 'w:sectPr')) {
    // Section break: the next section starts on a new page unless it is continuous.
    el.dataset.sect = '1';
    if (wa(path(p, 'w:pPr', 'w:sectPr', 'w:type'), 'val') !== 'continuous') el.dataset.breakAfter = '1';
  }
  if (f.outlineLvl != null && f.outlineLvl < 9) el.dataset.outline = String(f.outlineLvl);
  s.tabSize = pt(doc.defaultTabStop);

  // The paragraph mark's formatting sizes empty lines and labels.
  const mark = doc.markFmt(p);
  el.style.fontFamily = fontStack(mark);
  el.style.fontSize = `${(mark.sz ?? 20) / 2}pt`;

  const label = doc.labels().get(p);
  if (label) {
    const lf = { ...mark, ...doc.numbering.labelFmt(label.level) };
    el.dataset.label = label.text + (label.level.suff === 'space' ? ' ' : '');
    el.classList.add('w-li');
    s.setProperty('--lbl-w', label.level.suff === 'tab' ? pt(Math.max(hanging, label.text.length > 3 ? 0 : 360)) : '0pt');
    s.setProperty('--lbl-font', fontStack(lf));
    s.setProperty('--lbl-size', `${(lf.sz ?? 20) / 2}pt`);
    s.setProperty('--lbl-color', lf.color && lf.color !== 'auto' ? '#' + lf.color : '#000');
    s.setProperty('--lbl-weight', lf.b ? 'bold' : 'normal');
  }

  const items = doc.items(p);
  renderItems(p, items, ctx, el);
  if (ctx.editable) reg.sig.set(el, el.innerHTML);
  return el;
}

function renderItems(p: Element, items: Item[], ctx: RenderCtx, el: HTMLElement) {
  const { doc, reg } = ctx;
  // Open container wrappers (hyperlinks etc.), outermost first.
  const open: { xml: Element; dom: HTMLElement }[] = [];
  const parentFor = (containers: Element[]): HTMLElement => {
    let k = 0;
    while (k < open.length && k < containers.length && open[k].xml === containers[k]) k++;
    open.length = k;
    for (; k < containers.length; k++) {
      const c = containers[k];
      const w = document.createElement(c.localName === 'hyperlink' ? 'a' : 'span');
      w.className = 'w-' + c.localName;
      if (c.localName === 'ins' || c.localName === 'moveTo') w.title = `삽입: ${wa(c, 'author') ?? ''}`;
      if (c.localName === 'hyperlink') {
        const id = getRId(c, 'id');
        const url = id ? doc.pkg.rel(ctx.part, id)?.target : null;
        const anchor = wa(c, 'anchor');
        w.title = url ?? (anchor ? '#' + anchor : '');
        if (url) w.dataset.href = url;
      }
      if (ctx.editable) w.dataset.c = String(reg.containers.push(c) - 1);
      (open[k - 1]?.dom ?? el).appendChild(w);
      open.push({ xml: c, dom: w });
    }
    return open.at(-1)?.dom ?? el;
  };
  const atom = (it: Item, node: HTMLElement) => {
    if (ctx.editable) node.dataset.a = String(reg.atoms.push(it.nodes) - 1);
    if (!it.len) node.dataset.len = '0';
    return node;
  };
  let visible = 0;
  let last: Item | null = null;
  const comments = new Set<string>();
  for (const it of items) {
    const parent = parentFor(it.containers);
    switch (it.kind) {
      case 'text': {
        const span = document.createElement('span');
        if (ctx.editable) span.dataset.r = String(reg.runs.push(it.rPr) - 1);
        runCss(span, doc.runFmt(p, it.rPr));
        span.textContent = it.text;
        if (comments.size) span.classList.add('w-cmt-range');
        parent.appendChild(span);
        break;
      }
      case 'tab': {
        const span = atom(it, document.createElement('span'));
        span.className = 'w-tab';
        span.contentEditable = 'false';
        span.textContent = '\t';
        const f = doc.runFmt(p, it.rPr);
        span.style.fontSize = `${(f.sz ?? 20) / 2}pt`;
        if (f.u && f.u !== 'none') span.style.textDecoration = 'underline';
        parent.appendChild(span);
        break;
      }
      case 'br': {
        const br = atom(it, document.createElement('br'));
        parent.appendChild(br);
        break;
      }
      case 'page': {
        const span = atom(it, document.createElement('span'));
        span.className = 'w-pgbr';
        span.contentEditable = 'false';
        parent.appendChild(span);
        break;
      }
      case 'field': {
        const span = atom(it, document.createElement('span'));
        span.className = 'w-fld';
        span.contentEditable = 'false';
        runCss(span, doc.runFmt(p, it.rPr));
        const instr = ((it as Item & { instr?: string }).instr ?? '').trim();
        const kw = instr.split(/\s+/)[0]?.toUpperCase();
        let text = it.text;
        if (kw === 'PAGE' && ctx.page) text = String(ctx.page);
        else if ((kw === 'NUMPAGES' || kw === 'SECTIONPAGES') && ctx.pages) text = String(ctx.pages);
        span.textContent = text || (kw === 'PAGE' ? '#' : '');
        span.title = instr;
        if (kw === 'PAGE' || kw === 'NUMPAGES' || kw === 'SECTIONPAGES') span.dataset.fld = kw;
        parent.appendChild(span);
        break;
      }
      case 'atom': {
        const span = atom(it, document.createElement('span'));
        span.className = 'w-obj';
        span.contentEditable = 'false';
        renderAtom(p, it, ctx, span);
        parent.appendChild(span);
        break;
      }
      case 'hidden': {
        const span = atom(it, document.createElement('span'));
        const n = it.nodes[0];
        const local = n.namespaceURI === NS.w ? n.localName : '';
        if (local === 'commentRangeStart') comments.add(wa(n, 'id') ?? '');
        if (local === 'commentRangeEnd') comments.delete(wa(n, 'id') ?? '');
        const cref = local === 'r' ? child(n, 'w:commentReference') : null;
        if (local === 'del' || local === 'moveFrom') {
          // Tracked deletion: shown struck through, never editable, zero length in the model.
          span.className = 'w-del';
          span.contentEditable = 'false';
          span.textContent = Array.from(n.getElementsByTagNameNS(NS.w, 'delText')).map((t) => t.textContent).join('') || n.textContent || '';
          span.title = `삭제: ${wa(n, 'author') ?? ''}`;
        } else if (cref) {
          span.className = 'w-cmt';
          span.contentEditable = 'false';
          span.title = doc.commentText(wa(cref, 'id') ?? '');
        } else {
          span.hidden = true;
          span.className = 'w-hid';
        }
        parent.appendChild(span);
        break;
      }
    }
    if (it.len) { visible += it.len; last = it; }
  }
  // Empty paragraphs (and a trailing line break) need a placeholder <br> to keep their line.
  if (!visible || last?.kind === 'br') el.appendChild(document.createElement('br'));
}

// ---------------------------------------------------------------- inline objects

function extentPt(el: Element | null): { w: number; h: number } | null {
  if (!el) return null;
  return { w: num(el, 'cx') / 12700, h: num(el, 'cy') / 12700 };
}

function vmlSize(shape: Element | null): { w: number; h: number } | null {
  const style = shape?.getAttribute('style') ?? '';
  const get = (k: string) => {
    const m = new RegExp(`(?:^|;)\\s*${k}:\\s*([\\d.]+)(pt|in|px|cm|mm)?`).exec(style);
    if (!m) return null;
    const v = Number(m[1]);
    return ({ pt: v, in: v * 72, px: v * 0.75, cm: v * 28.3465, mm: v * 2.83465 } as Record<string, number>)[m[2] ?? 'px'];
  };
  const w = get('width'), h = get('height');
  return w && h ? { w, h } : null;
}

function placeholderBox(span: HTMLElement, size: { w: number; h: number } | null, label: string) {
  const box = document.createElement('span');
  box.className = 'w-objph';
  box.textContent = label;
  if (size) { box.style.width = size.w + 'pt'; box.style.height = size.h + 'pt'; }
  span.appendChild(box);
}

function renderDrawing(drawing: Element, ctx: RenderCtx, span: HTMLElement) {
  const frame = child(drawing, 'wp:inline') ?? child(drawing, 'wp:anchor');
  const size = extentPt(child(frame, 'wp:extent'));
  if (frame && is(frame, 'wp:anchor')) {
    span.classList.add('w-anchor');
    const wrap = children(frame).find((c) => c.localName.startsWith('wrap'))?.localName ?? 'wrapNone';
    const align = path(frame, 'wp:positionH', 'wp:align')?.textContent;
    const off = Number(path(frame, 'wp:positionH', 'wp:posOffset')?.textContent ?? 0) / 12700;
    const side = align === 'right' || align === 'outside' ? 'right' : align === 'center' ? 'center' : align === 'left' || align === 'inside' ? 'left' : off > 200 ? 'right' : 'left';
    const dist = (k: string) => num(frame, k) / 12700;
    if (wrap === 'wrapTopAndBottom' || side === 'center') {
      span.classList.add('w-anchor-block');
      span.style.textAlign = side === 'right' ? 'right' : side === 'center' ? 'center' : 'left';
      span.style.margin = `${dist('distT')}pt 0 ${dist('distB')}pt`;
    } else if (wrap !== 'wrapNone') {
      span.style.float = side;
      span.style.margin = side === 'right' ? `${dist('distT')}pt 0 ${dist('distB')}pt ${Math.max(6, dist('distL'))}pt` : `${dist('distT')}pt ${Math.max(6, dist('distR'))}pt ${dist('distB')}pt 0`;
    }
  }
  const blip = descendants(drawing, 'a:blip')[0];
  const txbx = descendants(drawing, 'w:txbxContent')[0];
  const uri = descendants(drawing, 'a:graphicData')[0]?.getAttribute('uri') ?? '';
  if (txbx && !blip) {
    const box = document.createElement('span');
    box.className = 'w-txbx';
    if (size) { box.style.width = size.w + 'pt'; box.style.minHeight = size.h + 'pt'; }
    renderBlocks(txbx, { ...ctx, editable: false }, box);
    span.appendChild(box);
    return;
  }
  if (blip) {
    const url = ctx.images.url(ctx.part, getRId(blip, 'embed') ?? getRId(blip, 'link'));
    if (url) {
      const img = document.createElement('img');
      img.src = url;
      img.draggable = false;
      if (size) { img.style.width = size.w + 'pt'; img.style.height = size.h + 'pt'; }
      span.appendChild(img);
      return;
    }
    placeholderBox(span, size, '그림');
    return;
  }
  placeholderBox(span, size, /chart/.test(uri) ? '차트' : /diagram/.test(uri) ? 'SmartArt' : '개체');
}

function renderVml(pict: Element, ctx: RenderCtx, span: HTMLElement) {
  const shape = descendants(pict, 'v:shape')[0] ?? descendants(pict, 'v:rect')[0] ?? null;
  const size = vmlSize(shape);
  const img = descendants(pict, 'v:imagedata')[0];
  const txbx = descendants(pict, 'w:txbxContent')[0];
  if (img) {
    const url = ctx.images.url(ctx.part, getRId(img, 'id'));
    if (url) {
      const el = document.createElement('img');
      el.src = url;
      el.draggable = false;
      if (size) { el.style.width = size.w + 'pt'; el.style.height = size.h + 'pt'; }
      span.appendChild(el);
      return;
    }
  }
  if (txbx) {
    const box = document.createElement('span');
    box.className = 'w-txbx';
    if (size) { box.style.width = size.w + 'pt'; box.style.minHeight = size.h + 'pt'; }
    renderBlocks(txbx, { ...ctx, editable: false }, box);
    span.appendChild(box);
    return;
  }
  placeholderBox(span, size, '개체');
}

function renderAtom(p: Element, it: Item, ctx: RenderCtx, span: HTMLElement) {
  const o = it.obj;
  if (!o) return;
  const fmt = it.rPr || is(o, 'w:sdt') ? ctx.doc.runFmt(p, it.rPr) : null;
  if (o.namespaceURI === NS.mc) {
    // Prefer the modern drawing in mc:Choice, fall back to VML.
    const d = descendants(o, 'w:drawing')[0];
    if (d) renderDrawing(d, ctx, span);
    else { const pict = descendants(o, 'w:pict')[0]; if (pict) renderVml(pict, ctx, span); else placeholderBox(span, null, '개체'); }
    return;
  }
  if (o.namespaceURI === NS.m) {
    span.classList.add('w-math');
    renderMath(o, span);
    return;
  }
  switch (o.localName) {
    case 'drawing': renderDrawing(o, ctx, span); return;
    case 'pict': case 'object': renderVml(o, ctx, span); return;
    case 'sym': {
      const code = parseInt(wa(o, 'char') ?? '20', 16);
      const font = wa(o, 'font');
      span.textContent = symbolChar(font, code);
      if (fmt) runCss(span, fmt);
      return;
    }
    case 'footnoteReference': case 'endnoteReference': {
      const n = ctx.notes ?? (ctx.notes = { n: 0 });
      span.textContent = String(++n.n);
      if (fmt) runCss(span, { ...fmt, vertAlign: 'superscript' });
      span.classList.add('w-noteref');
      return;
    }
    case 'noBreakHyphen': span.textContent = '‑'; if (fmt) runCss(span, fmt); return;
    case 'ruby': span.textContent = child(o, 'w:rubyBase')?.textContent ?? ''; if (fmt) runCss(span, fmt); return;
    case 'sdt': span.textContent = it.text; if (fmt) runCss(span, fmt); span.classList.add('w-sdt-inline'); return;
    default: placeholderBox(span, null, '개체');
  }
}

// ---------------------------------------------------------------- tables

interface CellInfo { tc: Element; col: number; span: number; vmerge: 'restart' | 'continue' | null }

export function tableGrid(tbl: Element): CellInfo[][] {
  return children(tbl, 'w:tr').map((tr) => {
    let col = wnum(path(tr, 'w:trPr', 'w:gridBefore'), 'val') ?? 0;
    return children(tr, 'w:tc').map((tc) => {
      const pr = child(tc, 'w:tcPr');
      const span = wnum(child(pr, 'w:gridSpan'), 'val') ?? 1;
      const vm = child(pr, 'w:vMerge');
      const info: CellInfo = { tc, col, span, vmerge: vm ? (wa(vm, 'val') === 'restart' ? 'restart' : 'continue') : null };
      col += span;
      return info;
    });
  });
}

function tableBorders(tbl: Element, ctx: RenderCtx): Record<string, Border> {
  const out: Record<string, Border> = {};
  const styleId = wa(path(tbl, 'w:tblPr', 'w:tblStyle'), 'val') ?? ctx.doc.styles.defaultTable;
  for (const s of ctx.doc.styles.chain(styleId)) Object.assign(out, parseBorders(path(s.el, 'w:tblPr', 'w:tblBorders')));
  Object.assign(out, parseBorders(path(tbl, 'w:tblPr', 'w:tblBorders')));
  return out;
}

function cellMargins(tbl: Element, ctx: RenderCtx) {
  const m = { top: 0, bottom: 0, left: 108, right: 108 };
  const styleId = wa(path(tbl, 'w:tblPr', 'w:tblStyle'), 'val') ?? ctx.doc.styles.defaultTable;
  const read = (el: Element | null) => {
    for (const c of children(el)) {
      const v = wnum(c, 'w');
      const k = c.localName === 'start' ? 'left' : c.localName === 'end' ? 'right' : c.localName;
      if (v != null && k in m) m[k as keyof typeof m] = v;
    }
  };
  for (const s of ctx.doc.styles.chain(styleId)) read(path(s.el, 'w:tblPr', 'w:tblCellMar'));
  read(path(tbl, 'w:tblPr', 'w:tblCellMar'));
  return m;
}

export function renderTable(tbl: Element, ctx: RenderCtx): HTMLElement {
  const table = document.createElement('table');
  table.className = 'w-tbl';
  if (ctx.editable) ctx.reg.dom.set(table, tbl);
  const tblPr = child(tbl, 'w:tblPr');
  const cols = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((g) => wnum(g, 'w') ?? 0);
  const cg = document.createElement('colgroup');
  for (const w of cols) { const c = document.createElement('col'); c.style.width = pt(w); cg.appendChild(c); }
  table.appendChild(cg);
  const total = cols.reduce((a, b) => a + b, 0);
  if (is(tbl.parentElement, 'w:tc')) {
    // Nested tables fit their cell (Word autofits them).
    table.style.width = '100%';
    if (total) cg.querySelectorAll('col').forEach((c, i) => { c.style.width = `${(cols[i] / total) * 100}%`; });
  } else if (total) table.style.width = pt(total);
  const jc = wa(child(tblPr, 'w:jc'), 'val');
  if (jc === 'center') table.style.alignSelf = 'center';
  else if (jc === 'right' || jc === 'end') table.style.alignSelf = 'flex-end';
  const ind = wnum(child(tblPr, 'w:tblInd'), 'w');
  if (ind) table.style.marginLeft = pt(ind);
  const borders = tableBorders(tbl, ctx);
  const hasBorders = Object.values(borders).some((b) => b.style !== 'none' && b.sz);
  if (!hasBorders) table.classList.add('nogrid');
  const mar = cellMargins(tbl, ctx);
  const grid = tableGrid(tbl);
  const rows = children(tbl, 'w:tr');
  const tbody = document.createElement('tbody');
  grid.forEach((row, ri) => {
    const tr = document.createElement('tr');
    const xtr = rows[ri];
    if (ctx.editable) ctx.reg.dom.set(tr, xtr);
    const h = path(xtr, 'w:trPr', 'w:trHeight');
    if (h) tr.style.height = pt(wnum(h, 'val') ?? 0);
    for (const cell of row) {
      if (cell.vmerge === 'continue') continue;
      const td = document.createElement('td');
      if (ctx.editable) ctx.reg.dom.set(td, cell.tc);
      td.dataset.col = String(cell.col);
      td.dataset.row = String(ri);
      if (cell.span > 1) td.colSpan = cell.span;
      if (cell.vmerge === 'restart') {
        let n = 1;
        for (let r = ri + 1; r < grid.length; r++) {
          const below = grid[r].find((c) => c.col === cell.col);
          if (below?.vmerge === 'continue') n++; else break;
        }
        if (n > 1) td.rowSpan = n;
      }
      const tcPr = child(cell.tc, 'w:tcPr');
      const own = parseBorders(child(tcPr, 'w:tcBorders'));
      const lastRow = ri + (td.rowSpan || 1) - 1 >= grid.length - 1;
      const edge = {
        top: own.top ?? (ri === 0 ? borders.top : borders.insideH),
        bottom: own.bottom ?? (lastRow ? borders.bottom : borders.insideH),
        left: own.left ?? own.start ?? (cell.col === 0 ? borders.left ?? borders.start : borders.insideV),
        right: own.right ?? own.end ?? (cell.col + cell.span >= cols.length ? borders.right ?? borders.end : borders.insideV),
      };
      for (const side of ['top', 'bottom', 'left', 'right'] as const) td.style.setProperty(`border-${side}`, cssBorder(edge[side]));
      const shd = wa(child(tcPr, 'w:shd'), 'fill');
      if (shd && shd !== 'auto') td.style.backgroundColor = '#' + shd;
      const va = wa(child(tcPr, 'w:vAlign'), 'val');
      td.style.verticalAlign = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
      const dir = textDir(wa(child(tcPr, 'w:textDirection'), 'val'));
      if (dir !== 'lrTb') td.classList.add(dir === 'tbRl' ? 'w-vert' : 'w-vert-up');
      const cm = { ...mar };
      for (const c of children(child(tcPr, 'w:tcMar'))) {
        const k = c.localName === 'start' ? 'left' : c.localName === 'end' ? 'right' : c.localName;
        const v = wnum(c, 'w');
        if (v != null && k in cm) cm[k as keyof typeof cm] = v;
      }
      td.style.padding = `${pt(cm.top)} ${pt(cm.right)} ${pt(cm.bottom)} ${pt(cm.left)}`;
      renderBlocks(cell.tc, ctx, td);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  return table;
}

/** Read-only rendering of a header/footer part for page `page` of `pages`. */
export function renderHeaderFooter(doc: DocxDocument, images: ImageCache, part: string, page: number, pages: number): HTMLElement {
  const host = document.createElement('div');
  const root = doc.pkg.getXml(part)?.documentElement;
  if (root) renderBlocks(root, { doc, images, part, reg: newRegistry(), editable: false, page, pages }, host);
  return host;
}

// ---------------------------------------------------------------- symbols & math

// Symbol font (Adobe Symbol encoding) letters → Greek.
const SYMBOL_GREEK = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';
const SYMBOL_greek = 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ';
const SYMBOL_MISC: Record<number, string> = { 0xb7: '•', 0xd7: '×', 0xb1: '±', 0xa5: '∞', 0xae: '→', 0xac: '←', 0xad: '↑', 0xaf: '↓', 0xde: '⇒', 0xdb: '⇔', 0xb9: '≠', 0xa3: '≤', 0xb3: '≥', 0xbb: '≈', 0xd6: '√', 0xe5: '∑', 0xf2: '∫', 0xb6: '∂', 0xa0: '€', 0xb0: '°', 0xd8: '¬', 0xd9: '∧', 0xda: '∨', 0xce: '∈', 0xc7: '∩', 0xc8: '∪', 0x22: '∀', 0x24: '∃' };

export function symbolChar(font: string | null, code: number): string {
  const c = code >= 0xf000 ? code - 0xf000 : code;
  if (/symbol/i.test(font ?? '')) {
    if (c >= 0x41 && c <= 0x5a) return SYMBOL_GREEK[c - 0x41];
    if (c >= 0x61 && c <= 0x7a) return SYMBOL_greek[c - 0x61];
    return SYMBOL_MISC[c] ?? String.fromCharCode(c);
  }
  if (/wingdings|webdings/i.test(font ?? '')) return bulletChar(String.fromCharCode(0xf000 + c));
  return String.fromCharCode(c);
}

/** Minimal OMML rendering: runs, scripts, fractions, radicals, delimiters, n-ary operators. */
function renderMath(el: Element, out: HTMLElement) {
  const M = NS.m;
  const kid = (e: Element, local: string) => Array.from(e.children).find((c) => c.localName === local && c.namespaceURI === M) ?? null;
  const box = (tag: string, cls?: string) => { const b = document.createElement(tag); if (cls) b.className = cls; return b; };
  const walk = (e: Element | null, host: HTMLElement) => {
    if (!e) return;
    for (const c of Array.from(e.children)) {
      if (c.namespaceURI !== M) continue;
      switch (c.localName) {
        case 'r': host.append(Array.from(c.children).filter((t) => t.localName === 't').map((t) => t.textContent).join('')); break;
        case 'sSup': case 'sSub': case 'sSubSup': {
          walk(kid(c, 'e'), host);
          if (c.localName !== 'sSup') { const b = box('sub'); walk(kid(c, 'sub'), b); host.appendChild(b); }
          if (c.localName !== 'sSub') { const b = box('sup'); walk(kid(c, 'sup'), b); host.appendChild(b); }
          break;
        }
        case 'f': {
          const f = box('span', 'w-frac');
          const n = box('span'), d = box('span');
          walk(kid(c, 'num'), n); walk(kid(c, 'den'), d);
          f.append(n, d);
          host.appendChild(f);
          break;
        }
        case 'rad': {
          host.append('√');
          const b = box('span', 'w-rad');
          walk(kid(c, 'e'), b);
          host.appendChild(b);
          break;
        }
        case 'd': {
          const pr = kid(c, 'dPr');
          const beg = pr && kid(pr, 'begChr') ? kid(pr, 'begChr')!.getAttributeNS(M, 'val') ?? '(' : '(';
          const end = pr && kid(pr, 'endChr') ? kid(pr, 'endChr')!.getAttributeNS(M, 'val') ?? ')' : ')';
          host.append(beg);
          Array.from(c.children).filter((x) => x.localName === 'e').forEach((x, i) => { if (i) host.append(','); walk(x, host); });
          host.append(end);
          break;
        }
        case 'nary': {
          const pr = kid(c, 'naryPr');
          const chr = pr && kid(pr, 'chr') ? kid(pr, 'chr')!.getAttributeNS(M, 'val') ?? '∫' : '∫';
          host.append(chr);
          const sub = box('sub'), sup = box('sup');
          walk(kid(c, 'sub'), sub); walk(kid(c, 'sup'), sup);
          if (sub.textContent) host.appendChild(sub);
          if (sup.textContent) host.appendChild(sup);
          walk(kid(c, 'e'), host);
          break;
        }
        case 'rPr': case 'ctrlPr': case 'oMathParaPr': break;
        default: walk(c, host);
      }
    }
  };
  if (el.localName === 'oMath') walk(el, out);
  else walk(el, out);
}
