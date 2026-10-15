// DocxDocument: the document model and every mutating command.
// As in Opptx (and HOP's rhwp core), the XML parts are the single source of truth: the renderer
// reads them, commands mutate them, and untouched content round-trips byte-for-byte.
import { CT, OpcPackage, REL, relsPathFor } from './package';
import { Label, Numbering } from './numbering';
import {
  Item, ParaProps, RunProps, ensurePPr, ensureRPr, insertAt, itemRuns, itemsInRange, makeRun, markRPr,
  mergeRuns, paraItems, paraLength, paraText, setParaProps, setRunProps, setText, splitParagraph, walkParagraphs,
} from './model';
import { ParaFmt, RunFmt, StyleSheet, ThemeInfo, parseTheme } from './styles';
import { NS, PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TBLPR_ORDER, TCPR_ORDER, TRPR_ORDER, child, children, create, descendants, ensureOrdered, escapeXml, is, parseXml, path, removeChild, serializeXml, setAttr, wa, wnum } from './xml';

export const TWIP_PER_PT = 20;
export const EMU_PER_PT = 12700;
export const TWIP_PER_CM = 566.929;

export interface Pos { pi: number; off: number }
export interface Range { start: Pos; end: Pos }

export interface PageSetup {
  w: number; h: number; // twips
  orient: 'portrait' | 'landscape';
  top: number; bottom: number; left: number; right: number; header: number; footer: number; gutter: number;
}

interface HistoryEntry { label: string; before: Map<string, string | null>; after: Map<string, string | null>; selBefore: unknown; selAfter: unknown }

export const XMLNS = `xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w14="${NS.w14}"`;

/** Parse a fragment of WordprocessingML into nodes owned by `doc`. */
export function fragment(doc: Document, xml: string): Element[] {
  const tmp = parseXml(`<root ${XMLNS}>${xml}</root>`);
  return children(tmp.documentElement).map((c) => doc.importNode(c, true) as Element);
}

export function comparePos(a: Pos, b: Pos) { return a.pi - b.pi || a.off - b.off; }
export function normRange(r: Range): Range { return comparePos(r.start, r.end) <= 0 ? r : { start: r.end, end: r.start }; }

// Built-in style definitions added on demand when a document lacks them (Korean Word defaults).
const BUILTIN_STYLES: Record<string, string> = {
  Heading1: '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>',
  Heading2: '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi"/><w:color w:val="2F5496"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>',
  Heading3: '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi"/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>',
  Title: '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>',
  Subtitle: '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>',
  Quote: '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>',
  ListParagraph: '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>',
  NoSpacing: '<w:style w:type="paragraph" w:styleId="NoSpacing"><w:name w:val="No Spacing"/><w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>',
  TableGrid: '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>',
};

const BULLETS = ['●', '○', '■', '●', '○', '■', '●', '○', '■'];
const NUM_FMTS: [string, string][] = [['decimal', '%1.'], ['ganada', '%2.'], ['decimal', '%3)'], ['ganada', '%4)'], ['decimalEnclosedCircle', '%5'], ['lowerLetter', '%6.'], ['lowerRoman', '%7.'], ['decimal', '(%8)'], ['lowerLetter', '(%9)']];

export class DocxDocument {
  readonly docPath: string;
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  dirty = false;
  private depth = 0;
  version = 0;
  onChange: ((label: string) => void) | null = null;
  /** Supplies the editor selection to store with history entries. */
  selection: (() => unknown) | null = null;
  private cache: { version: number; paras: Element[]; index: Map<Element, number>; labels: Map<Element, Label> | null } | null = null;
  private _styles: StyleSheet | null = null;
  private _numbering: Numbering | null = null;
  private _theme: ThemeInfo | null = null;
  private fmtCache = new Map<Element, ParaFmt>();

  private constructor(readonly pkg: OpcPackage) {
    const rel = pkg.relOfType('', REL.officeDocument);
    this.docPath = rel?.target ?? 'word/document.xml';
  }

  static async load(data: ArrayBuffer | Uint8Array): Promise<DocxDocument> {
    const pkg = await OpcPackage.load(data);
    const doc = new DocxDocument(pkg);
    if (!pkg.getXml(doc.docPath)) throw new Error('Word 문서(word/document.xml)를 찾을 수 없습니다.');
    if (!doc.body) throw new Error('문서 본문(w:body)이 없습니다.');
    // An editable document always has at least one paragraph to put the caret in.
    if (!child(doc.body, 'w:p')) doc.body.insertBefore(create(doc.xml, 'w:p'), child(doc.body, 'w:sectPr'));
    return doc;
  }

  // ---------------------------------------------------------------- parts

  get xml(): XMLDocument { return this.pkg.getXml(this.docPath)!; }
  get body(): Element { return child(this.xml.documentElement, 'w:body')!; }
  private partOf(type: string) { return this.pkg.relOfType(this.docPath, type)?.target ?? null; }
  get stylesPath() { return this.partOf(REL.styles) ?? 'word/styles.xml'; }
  get numberingPath() { return this.partOf(REL.numbering); }
  get settings(): XMLDocument | null { const p = this.partOf(REL.settings); return p ? this.pkg.getXml(p) : null; }

  get theme(): ThemeInfo {
    if (!this._theme) { const p = this.partOf(REL.theme); this._theme = parseTheme(p ? this.pkg.getXml(p) : null); }
    return this._theme;
  }

  get styles(): StyleSheet {
    if (!this._styles) this._styles = new StyleSheet(this.pkg.getXml(this.stylesPath), this.theme);
    return this._styles;
  }

  get numbering(): Numbering {
    if (!this._numbering) { const p = this.numberingPath; this._numbering = new Numbering(p ? this.pkg.getXml(p) : null, this.styles); }
    return this._numbering;
  }

  get defaultTabStop(): number { return wnum(child(this.settings?.documentElement, 'w:defaultTabStop'), 'val') ?? 720; }
  get evenAndOddHeaders(): boolean { return !!child(this.settings?.documentElement, 'w:evenAndOddHeaders'); }

  private invalidate() {
    this.version++;
    this.cache = null;
    this._styles = null;
    this._numbering = null;
    this.fmtCache.clear();
  }

  // ---------------------------------------------------------------- paragraphs

  private ensureCache() {
    if (this.cache && this.cache.version === this.version) return this.cache;
    const paras = walkParagraphs(this.body);
    this.cache = { version: this.version, paras, index: new Map(paras.map((p, i) => [p, i])), labels: null };
    return this.cache;
  }

  paragraphs(): Element[] { return this.ensureCache().paras; }
  para(pi: number): Element | null { return this.paragraphs()[pi] ?? null; }
  indexOf(p: Element): number { return this.ensureCache().index.get(p) ?? -1; }

  /** List labels for every numbered paragraph. */
  labels(): Map<Element, Label> {
    const c = this.ensureCache();
    if (!c.labels) c.labels = this.numbering.labels(c.paras.map((el) => { const f = this.paraFmt(el); return { el, numId: f.numId, ilvl: f.ilvl }; }));
    return c.labels;
  }

  tableStyleOf(p: Element): string | null {
    for (let n = p.parentElement; n; n = n.parentElement) {
      if (is(n, 'w:tbl')) return wa(path(n, 'w:tblPr', 'w:tblStyle'), 'val') ?? this.styles.defaultTable;
      if (is(n, 'w:body')) break;
    }
    return null;
  }

  pStyleOf(p: Element): string | null { return wa(path(p, 'w:pPr', 'w:pStyle'), 'val'); }

  /** Effective paragraph formatting (styles, numbering level, direct). */
  paraFmt(p: Element): ParaFmt {
    const hit = this.fmtCache.get(p);
    if (hit) return hit;
    const st = this.styles;
    const f = st.styleParaFmt(this.pStyleOf(p), this.tableStyleOf(p));
    const direct = st.parsePPr(child(p, 'w:pPr'));
    const numId = direct.numId ?? f.numId;
    const ilvl = direct.ilvl ?? f.ilvl ?? 0;
    if (numId && numId !== '0') {
      const lvl = this.numbering.level(numId, ilvl);
      if (lvl) Object.assign(f, st.parsePPr(lvl.pPr));
    }
    Object.assign(f, direct);
    f.numId = numId;
    f.ilvl = ilvl;
    this.fmtCache.set(p, f);
    return f;
  }

  runFmt(p: Element, rPr: Element | null): RunFmt {
    return this.styles.runFmt(this.pStyleOf(p), this.tableStyleOf(p), rPr);
  }

  markFmt(p: Element): RunFmt { return this.runFmt(p, path(p, 'w:pPr', 'w:rPr')); }

  /** Direct run properties in effect at a position (the item before it, as Word does). */
  rPrAt(p: Element, off: number): Element | null {
    let pos = 0;
    let last: Element | null | undefined;
    for (const it of paraItems(p)) {
      if (it.len && pos < off && (it.kind === 'text' || it.kind === 'tab' || it.kind === 'field')) last = it.rPr;
      if (it.len && pos === 0 && off === 0 && it.kind === 'text') { last = it.rPr; break; }
      pos += it.len;
      if (pos >= off && last !== undefined) break;
    }
    return last !== undefined ? last : markRPr(p);
  }

  /** Effective character formatting at a position (for the style bar). */
  fmtAt(pos: Pos): RunFmt | null {
    const p = this.para(pos.pi);
    return p ? this.runFmt(p, this.rPrAt(p, pos.off)) : null;
  }

  // ---------------------------------------------------------------- history

  /**
   * Run a mutation as one undoable step. `parts` lists the XML parts it may touch.
   */
  transact<T>(label: string, parts: string[], fn: () => T): T {
    if (this.depth > 0) return fn();
    this.depth++;
    try {
      return this.transactOuter(label, parts, fn);
    } finally {
      this.depth--;
    }
  }

  private transactOuter<T>(label: string, parts: string[], fn: () => T): T {
    const list = [...new Set([this.docPath, ...parts])];
    const before = this.pkg.snapshot(list);
    const selBefore = this.selection?.();
    let result: T;
    try {
      result = fn();
    } catch (e) {
      this.pkg.restore(before);
      this.invalidate();
      throw e;
    }
    const after = this.pkg.snapshot(list);
    const changed = [...after].some(([p, s]) => before.get(p) !== s);
    this.invalidate();
    if (changed) {
      this.undoStack.push({ label, before, after, selBefore, selAfter: this.selection?.() });
      if (this.undoStack.length > 300) this.undoStack.shift();
      this.redoStack = [];
      this.dirty = true;
      this.depth--;
      try { this.onChange?.(label); } finally { this.depth++; }
    }
    return result;
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }
  undoLabel() { return this.undoStack.at(-1)?.label; }
  redoLabel() { return this.redoStack.at(-1)?.label; }

  undo(): unknown {
    const e = this.undoStack.pop();
    if (!e) return undefined;
    this.pkg.restore(e.before);
    this.redoStack.push(e);
    this.dirty = true;
    this.invalidate();
    this._theme = null;
    this.onChange?.('되돌리기');
    return e.selBefore;
  }

  redo(): unknown {
    const e = this.redoStack.pop();
    if (!e) return undefined;
    this.pkg.restore(e.after);
    this.undoStack.push(e);
    this.dirty = true;
    this.invalidate();
    this.onChange?.('다시 실행');
    return e.selAfter;
  }

  async save(): Promise<Blob> {
    this.touchModified();
    const blob = await this.pkg.save();
    this.dirty = false;
    return blob;
  }

  async saveBytes(): Promise<Uint8Array> {
    this.touchModified();
    const bytes = await this.pkg.saveBytes();
    this.dirty = false;
    return bytes;
  }

  private touchModified() {
    const core = this.pkg.getXml('docProps/core.xml');
    const m = core?.getElementsByTagNameNS('http://purl.org/dc/terms/', 'modified')[0];
    if (m) m.textContent = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  }

  // ---------------------------------------------------------------- section / page

  sectPr(): Element {
    let s = child(this.body, 'w:sectPr');
    if (!s) {
      s = create(this.xml, 'w:sectPr');
      this.body.appendChild(s);
    }
    return s;
  }

  pageSetup(): PageSetup {
    const s = child(this.body, 'w:sectPr');
    const sz = child(s, 'w:pgSz'), mar = child(s, 'w:pgMar');
    const w = wnum(sz, 'w') ?? 11906, h = wnum(sz, 'h') ?? 16838;
    return {
      w, h, orient: wa(sz, 'orient') === 'landscape' || w > h ? 'landscape' : 'portrait',
      top: Math.abs(wnum(mar, 'top') ?? 1440), bottom: Math.abs(wnum(mar, 'bottom') ?? 1440),
      left: wnum(mar, 'left') ?? 1440, right: wnum(mar, 'right') ?? 1440,
      header: wnum(mar, 'header') ?? 720, footer: wnum(mar, 'footer') ?? 720, gutter: wnum(mar, 'gutter') ?? 0,
    };
  }

  setPageSetup(ps: PageSetup) {
    this.transact('쪽 설정', [], () => {
      const all = [...descendants(this.body, 'w:sectPr')];
      if (!all.length) all.push(this.sectPr());
      for (const s of all) {
        const sz = ensureOrdered(s, 'w:pgSz', SECTPR_ORDER);
        setAttr(sz, 'w:w', Math.round(ps.w));
        setAttr(sz, 'w:h', Math.round(ps.h));
        if (ps.orient === 'landscape') setAttr(sz, 'w:orient', 'landscape'); else sz.removeAttributeNS(NS.w, 'orient');
        const mar = ensureOrdered(s, 'w:pgMar', SECTPR_ORDER);
        for (const k of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'] as const) setAttr(mar, 'w:' + k, Math.round(ps[k]));
      }
    });
  }

  /** Header/footer parts for page `n` (1-based) of the final section. */
  headerFooter(n: number): { header: string | null; footer: string | null } {
    const s = child(this.body, 'w:sectPr');
    const titlePg = !!child(s, 'w:titlePg') && wa(child(s, 'w:titlePg'), 'val') !== '0';
    const type = titlePg && n === 1 ? 'first' : this.evenAndOddHeaders && n % 2 === 0 ? 'even' : 'default';
    const find = (tag: string) => {
      const ref = children(s, tag).find((r) => (wa(r, 'type') ?? 'default') === type);
      const id = ref?.getAttributeNS(NS.r, 'id');
      return id ? this.pkg.rel(this.docPath, id)?.target ?? null : null;
    };
    return { header: find('w:headerReference'), footer: find('w:footerReference') };
  }

  // ---------------------------------------------------------------- styles & numbering setup

  private stylesDoc(): XMLDocument {
    let d = this.pkg.getXml(this.stylesPath);
    if (!d) {
      d = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${XMLNS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`);
      this.pkg.setXml(this.stylesPath, d);
      this.pkg.addRel(this.docPath, REL.styles, this.stylesPath);
      this.pkg.addOverride(this.stylesPath, CT.styles);
    }
    return d;
  }

  /** Make sure a built-in style exists (call inside a transaction that lists stylesParts()). */
  ensureStyle(id: string): boolean {
    if (this.styles.style(id)) return true;
    const def = BUILTIN_STYLES[id];
    if (!def) return false;
    const d = this.stylesDoc();
    for (const el of fragment(d, def)) d.documentElement.appendChild(el);
    this._styles = null;
    return true;
  }

  stylesParts() { return [this.stylesPath, relsPathFor(this.docPath), '[Content_Types].xml']; }

  private numberingDoc(): { path: string; doc: XMLDocument } {
    let p = this.numberingPath;
    if (!p) {
      p = 'word/numbering.xml';
      this.pkg.setXml(p, parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${XMLNS}/>`));
      this.pkg.addRel(this.docPath, REL.numbering, p);
      this.pkg.addOverride(p, CT.numbering);
    }
    return { path: p, doc: this.pkg.getXml(p)! };
  }

  numberingParts() { return ['word/numbering.xml', ...(this.numberingPath ? [this.numberingPath] : []), relsPathFor(this.docPath), '[Content_Types].xml']; }

  /** A numbering instance for bullets or numbers; reuses one the document already has. */
  listNum(kind: 'bullet' | 'number', near?: Element): string {
    const nb = this.numbering;
    // Continue the list of the previous paragraph when it is the same kind.
    if (near) {
      const idx = this.indexOf(near);
      for (let i = idx - 1; i >= Math.max(0, idx - 1); i--) {
        const f = this.paraFmt(this.paragraphs()[i]);
        if (f.numId && f.numId !== '0' && nb.isBullet(f.numId, 0) === (kind === 'bullet')) return f.numId;
      }
    }
    const { doc } = this.numberingDoc();
    const root = doc.documentElement;
    for (const n of children(root, 'w:num')) {
      const id = wa(n, 'numId')!;
      const lvl = nb.level(id, 0);
      // Only reuse full multi-level lists that are not tied to a list style (e.g. "List Number").
      const abs = nb.abstractOf(id);
      const multi = children(abs, 'w:lvl').length >= 9 && !child(children(abs, 'w:lvl')[0], 'w:pStyle');
      if (lvl && multi && (lvl.fmt === 'bullet') === (kind === 'bullet') && (kind === 'bullet' || lvl.fmt === 'decimal')) return id;
    }
    const absIds = children(root, 'w:abstractNum').map((a) => Number(wa(a, 'abstractNumId')) || 0);
    const absId = Math.max(-1, ...absIds) + 1;
    const numId = Math.max(0, ...children(root, 'w:num').map((n) => Number(wa(n, 'numId')) || 0)) + 1;
    let lvls = '';
    for (let i = 0; i < 9; i++) {
      const [fmt, text] = kind === 'bullet' ? ['bullet', BULLETS[i]] : NUM_FMTS[i];
      const left = 800 + i * 400;
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="400"/></w:pPr>${kind === 'bullet' ? '<w:rPr><w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:eastAsia="맑은 고딕" w:hint="eastAsia"/></w:rPr>' : ''}</w:lvl>`;
    }
    const [abs] = fragment(doc, `<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`);
    const firstNum = child(root, 'w:num');
    root.insertBefore(abs, firstNum ?? children(root).find((c) => c.localName === 'numIdMacAtCleanup') ?? null);
    const [num] = fragment(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`);
    const after = children(root, 'w:num').at(-1) ?? abs;
    after.after(num);
    this._numbering = null;
    return String(numId);
  }

  // ---------------------------------------------------------------- range helpers

  private eachInRange(r: Range, fn: (p: Element, s: number, e: number, full: boolean) => void) {
    const { start, end } = normRange(r);
    const paras = this.paragraphs().slice(start.pi, end.pi + 1);
    paras.forEach((p, k) => {
      const pi = start.pi + k;
      const len = paraLength(p);
      const s = pi === start.pi ? start.off : 0;
      const e = pi === end.pi ? Math.min(end.off, len) : len;
      fn(p, s, e, s === 0 && e === len && (pi !== end.pi || start.pi === end.pi || e === len));
    });
  }

  /** Word boundaries around a collapsed position, or null when not inside a word. */
  wordAt(pos: Pos): Range | null {
    const p = this.para(pos.pi);
    if (!p) return null;
    const t = paraText(p);
    const w = /[\p{L}\p{N}_]/u;
    if (!(pos.off > 0 && w.test(t[pos.off - 1]) && pos.off < t.length && w.test(t[pos.off]))) return null;
    let s = pos.off, e = pos.off;
    while (s > 0 && w.test(t[s - 1])) s--;
    while (e < t.length && w.test(t[e])) e++;
    return { start: { pi: pos.pi, off: s }, end: { pi: pos.pi, off: e } };
  }

  // ---------------------------------------------------------------- character formatting

  applyRunProps(range: Range, props: RunProps, label = '글자 모양') {
    this.transact(label, [], () => {
      this.eachInRange(range, (p, s, e, full) => {
        if (s < e) {
          for (const it of itemsInRange(p, s, e)) for (const r of itemRuns(it)) setRunProps(ensureRPr(r), props);
          mergeRuns(p);
        }
        if (full || paraLength(p) === 0) {
          const pPr = ensurePPr(p);
          setRunProps(ensureOrdered(pPr, 'w:rPr', PPR_ORDER), props);
        }
      });
    });
  }

  clearRunProps(range: Range) {
    this.transact('글자 서식 지우기', [], () => {
      this.eachInRange(range, (p, s, e, full) => {
        if (s < e) {
          for (const it of itemsInRange(p, s, e)) for (const r of itemRuns(it)) removeChild(r, 'w:rPr');
          mergeRuns(p);
        }
        if (full) removeChild(child(p, 'w:pPr'), 'w:rPr');
      });
    });
  }

  // ---------------------------------------------------------------- paragraph formatting

  paraRange(r: Range): number[] {
    const { start, end } = normRange(r);
    return Array.from({ length: end.pi - start.pi + 1 }, (_, k) => start.pi + k);
  }

  applyParaProps(pis: number[], props: ParaProps | ((p: Element, f: ParaFmt) => ParaProps), label = '문단 모양') {
    this.transact(label, [], () => {
      const paras = pis.map((i) => this.para(i)).filter((p): p is Element => !!p);
      const fmts = paras.map((p) => this.paraFmt(p));
      paras.forEach((p, k) => setParaProps(p, typeof props === 'function' ? props(p, fmts[k]) : props));
    });
  }

  setStyle(pis: number[], styleId: string) {
    this.transact('스타일', this.stylesParts(), () => {
      this.ensureStyle(styleId);
      const def = this.styles.defaultPara;
      for (const i of pis) {
        const p = this.para(i);
        if (p) setParaProps(p, { style: styleId === def ? null : styleId });
      }
    });
  }

  toggleList(pis: number[], kind: 'bullet' | 'number') {
    const paras = pis.map((i) => this.para(i)).filter((p): p is Element => !!p);
    if (!paras.length) return;
    const nb = this.numbering;
    const isKind = (p: Element) => {
      const f = this.paraFmt(p);
      return !!f.numId && f.numId !== '0' && nb.has(f.numId) && nb.isBullet(f.numId, f.ilvl ?? 0) === (kind === 'bullet');
    };
    const allOn = paras.every(isKind);
    this.transact(kind === 'bullet' ? '글머리표' : '문단 번호', this.numberingParts(), () => {
      if (allOn) {
        for (const p of paras) {
          const styleNum = this.styles.styleParaFmt(this.pStyleOf(p), null).numId;
          setParaProps(p, { numId: styleNum ? '0' : null });
          if (!styleNum) removeChild(child(p, 'w:pPr'), 'w:numPr');
        }
        return;
      }
      const id = this.listNum(kind, paras[0]);
      for (const p of paras) {
        const f = this.paraFmt(p);
        setParaProps(p, { numId: id, ilvl: f.numId && f.numId !== '0' ? f.ilvl ?? 0 : 0 });
        // List levels carry their own indents; drop direct ones that would override them.
        const ind = path(p, 'w:pPr', 'w:ind');
        if (ind) ind.remove();
      }
    });
  }

  /** Indent: list level ±1 for list paragraphs, otherwise left indent ± one tab stop. */
  indent(pis: number[], dir: 1 | -1) {
    const step = this.defaultTabStop;
    this.applyParaProps(pis, (_p, f) => {
      if (f.numId && f.numId !== '0') return { ilvl: Math.max(0, Math.min(8, (f.ilvl ?? 0) + dir)) };
      const left = Math.max(0, Math.round(((f.left ?? 0) + dir * step) / step) * step);
      return { left: left || null };
    }, dir > 0 ? '들여쓰기' : '내어쓰기');
  }

  // ---------------------------------------------------------------- structure edits

  /** Enter: split the paragraph. An empty list item leaves the list instead. Returns the caret. */
  splitAt(pos: Pos): Pos {
    const p = this.para(pos.pi);
    if (!p) return pos;
    const f = this.paraFmt(p);
    if (paraLength(p) === 0 && f.numId && f.numId !== '0') {
      this.transact('목록 끝내기', [], () => {
        const styleNum = this.styles.styleParaFmt(this.pStyleOf(p), null).numId;
        if (styleNum) setParaProps(p, { numId: '0' }); else removeChild(child(p, 'w:pPr'), 'w:numPr');
        path(p, 'w:pPr', 'w:ind')?.remove();
      });
      return pos;
    }
    return this.transact('문단 나누기', [], () => {
      const atEnd = pos.off >= paraLength(p);
      const second = splitParagraph(p, pos.off);
      const style = this.styles.style(this.pStyleOf(p));
      // Headings and titles continue with the normal style even when w:next is missing.
      const heading = style && (this.styles.styleParaFmt(style.id, null).outlineLvl != null || /^(title|subtitle)$/i.test(style.name));
      const next = style?.next ?? (heading ? this.styles.defaultPara ?? undefined : undefined);
      if (atEnd && style && next && next !== style.id) {
        const nextNum = this.styles.styleParaFmt(next, null).numId;
        setParaProps(second, { style: next === this.styles.defaultPara ? null : next });
        if (!nextNum && f.numId && !path(p, 'w:pPr', 'w:numPr')) removeChild(child(second, 'w:pPr'), 'w:numPr');
      }
      if (atEnd) path(second, 'w:pPr', 'w:pageBreakBefore')?.remove();
      return { pi: pos.pi + 1, off: 0 };
    });
  }

  /** Insert plain text (newlines become paragraphs, tabs become tabs). Returns the caret after it. */
  insertText(pos: Pos, text: string, label = '입력'): Pos {
    return this.transact(label, [], () => {
      const lines = text.replace(/\r\n?/g, '\n').split('\n');
      let cur = pos;
      lines.forEach((line, k) => {
        if (k > 0) {
          const p = this.para(cur.pi)!;
          splitParagraph(p, cur.off);
          this.invalidate();
          cur = { pi: cur.pi + 1, off: 0 };
        }
        if (!line) return;
        const p = this.para(cur.pi)!;
        const rPr = this.rPrAt(p, cur.off);
        const nodes: Element[] = [];
        line.split('\t').forEach((seg, j) => {
          if (j > 0) nodes.push(makeRun(this.xml, rPr, create(this.xml, 'w:tab')));
          if (seg) nodes.push(makeRun(this.xml, rPr, seg));
        });
        insertAt(p, cur.off, nodes);
        mergeRuns(p);
        cur = { pi: cur.pi, off: cur.off + line.length };
      });
      return cur;
    });
  }

  insertTab(pos: Pos): Pos { return this.insertText(pos, '\t', '탭'); }

  insertPageBreak(pos: Pos): Pos {
    return this.transact('쪽 나누기', [], () => {
      const p = this.para(pos.pi)!;
      const rPr = this.rPrAt(p, pos.off);
      const second = splitParagraph(p, pos.off);
      const br = makeRun(this.xml, rPr, create(this.xml, 'w:br', { 'w:type': 'page' }));
      p.appendChild(br);
      path(second, 'w:pPr', 'w:pageBreakBefore')?.remove();
      return { pi: pos.pi + 1, off: 0 };
    });
  }

  /** Remove the text range [start,end) inside one paragraph and put `text` there. */
  private replaceIn(p: Element, s: number, e: number, text: string) {
    const items = itemsInRange(p, s, e);
    const rPr = items.find((it) => it.kind === 'text')?.rPr ?? this.rPrAt(p, s);
    const run = text ? makeRun(this.xml, rPr, text) : null;
    let placed = false;
    for (const it of items) {
      for (const n of it.nodes) {
        if (run && !placed) { n.before(run); placed = true; }
        n.remove();
      }
    }
    if (run && !placed) insertAt(p, s, [run]);
    mergeRuns(p);
  }

  /** Copy a range as standalone paragraphs (first and last trimmed). */
  extractRange(r: Range): string[] {
    const { start, end } = normRange(r);
    const out: string[] = [];
    for (let pi = start.pi; pi <= end.pi; pi++) {
      const orig = this.para(pi);
      if (!orig) continue;
      const p = orig.cloneNode(true) as Element;
      const len = paraLength(p);
      const s = pi === start.pi ? start.off : 0;
      const e = pi === end.pi ? end.off : len;
      if (e < len) for (const it of itemsInRange(p, e, len)) it.nodes.forEach((n) => n.remove());
      if (s > 0) for (const it of itemsInRange(p, 0, s)) it.nodes.forEach((n) => n.remove());
      out.push(new XMLSerializer().serializeToString(p));
    }
    return out;
  }

  /** Paste paragraphs produced by extractRange at a collapsed position. */
  insertFragment(pos: Pos, xmls: string[]): Pos {
    if (!xmls.length) return pos;
    return this.transact('붙이기', [], () => {
      const paras = xmls.map((x) => this.xml.importNode(parseXml(x).documentElement, true) as Element);
      for (const p of paras) for (const a of ['paraId', 'textId']) p.removeAttributeNS(NS.w14, a);
      const target = this.para(pos.pi)!;
      const content = (p: Element) => children(p).filter((c) => !is(c, 'w:pPr'));
      if (paras.length === 1) {
        const len = paraLength(paras[0]);
        insertAt(target, pos.off, content(paras[0]));
        mergeRuns(target);
        return { pi: pos.pi, off: pos.off + len };
      }
      const tail = splitParagraph(target, pos.off);
      // First fragment paragraph joins the current one, the last one joins the tail.
      insertAt(target, pos.off, content(paras[0]));
      let ref = target;
      for (const p of paras.slice(1, -1)) { ref.after(p); ref = p; }
      const last = paras[paras.length - 1];
      const lastLen = paraLength(last);
      insertAt(tail, 0, content(last));
      mergeRuns(target);
      mergeRuns(tail);
      this.invalidate();
      return { pi: this.indexOf(tail), off: lastLen };
    });
  }

  // ---------------------------------------------------------------- tables

  insertTable(pos: Pos, rows: number, cols: number): Pos {
    return this.transact('표 넣기', this.stylesParts(), () => {
      const hasGrid = this.ensureStyle('TableGrid');
      const ps = this.pageSetup();
      let p = this.para(pos.pi)!;
      const avail = this.cellWidth(p) ?? ps.w - ps.left - ps.right;
      const cw = Math.floor(avail / cols);
      const cells = Array.from({ length: cols }, () => `<w:tc><w:tcPr><w:tcW w:w="${cw}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`).join('');
      const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('');
      const xml = `<w:tbl><w:tblPr>${hasGrid ? '<w:tblStyle w:val="TableGrid"/>' : ''}<w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${cw}"/>`.repeat(cols)}</w:tblGrid>${`<w:tr>${cells}</w:tr>`.repeat(rows)}</w:tbl>`;
      const [tbl] = fragment(this.xml, xml);
      const len = paraLength(p);
      if (len === 0) p.before(tbl);
      else {
        if (pos.off > 0 && pos.off < len) { splitParagraph(p, pos.off); }
        else if (pos.off >= len) {
          const np = create(this.xml, 'w:p');
          const pPr = child(p, 'w:pPr');
          if (pPr) { const c = pPr.cloneNode(true) as Element; removeChild(c, 'w:sectPr'); removeChild(c, 'w:numPr'); np.appendChild(c); }
          p.after(np);
        }
        if (pos.off === 0) p.before(tbl); else { p = p.nextElementSibling as Element; p.before(tbl); }
      }
      this.invalidate();
      return { pi: this.indexOf(descendants(tbl, 'w:p')[0]), off: 0 };
    });
  }

  private cellWidth(p: Element): number | null {
    const tc = p.parentElement;
    if (!is(tc, 'w:tc')) return null;
    return wnum(path(tc, 'w:tcPr', 'w:tcW'), 'w') ?? null;
  }

  /** Cell, row and table containing paragraph `pi`. */
  cellOf(pi: number): { tc: Element; tr: Element; tbl: Element } | null {
    const p = this.para(pi);
    const tc = p?.parentElement;
    if (!tc || !is(tc, 'w:tc')) return null;
    const tr = tc.parentElement!, tbl = tr.parentElement!;
    return { tc, tr, tbl };
  }

  private gridIndex(tc: Element): number {
    let col = wnum(path(tc.parentElement!, 'w:trPr', 'w:gridBefore'), 'val') ?? 0;
    for (const c of children(tc.parentElement!, 'w:tc')) {
      if (c === tc) return col;
      col += wnum(path(c, 'w:tcPr', 'w:gridSpan'), 'val') ?? 1;
    }
    return col;
  }

  private cellAtGrid(tr: Element, col: number): Element | null {
    let g = wnum(path(tr, 'w:trPr', 'w:gridBefore'), 'val') ?? 0;
    for (const c of children(tr, 'w:tc')) {
      const span = wnum(path(c, 'w:tcPr', 'w:gridSpan'), 'val') ?? 1;
      if (col >= g && col < g + span) return c;
      g += span;
    }
    return null;
  }

  private blankCell(tc: Element): Element {
    const c = tc.cloneNode(false) as Element;
    const tcPr = child(tc, 'w:tcPr');
    if (tcPr) { const x = tcPr.cloneNode(true) as Element; removeChild(x, 'w:vMerge'); c.appendChild(x); }
    const p0 = child(tc, 'w:p');
    const p = create(this.xml, 'w:p');
    const pPr = child(p0, 'w:pPr');
    if (pPr) p.appendChild(pPr.cloneNode(true));
    c.appendChild(p);
    return c;
  }

  tableOp(pi: number, op: 'row-above' | 'row-below' | 'col-left' | 'col-right' | 'del-row' | 'del-col' | 'del-table'): Pos | null {
    const cell = this.cellOf(pi);
    if (!cell) return null;
    const { tc, tr, tbl } = cell;
    const labels = { 'row-above': '위에 행 삽입', 'row-below': '아래에 행 삽입', 'col-left': '왼쪽에 열 삽입', 'col-right': '오른쪽에 열 삽입', 'del-row': '행 삭제', 'del-col': '열 삭제', 'del-table': '표 삭제' };
    return this.transact(labels[op], [], () => {
      const firstP = (el: Element) => descendants(el, 'w:p')[0];
      const finish = (focus: Element | null): Pos => {
        this.invalidate();
        const p = focus ? (is(focus, 'w:p') ? focus : firstP(focus)) : null;
        return { pi: p ? Math.max(0, this.indexOf(p)) : Math.max(0, Math.min(pi, this.paragraphs().length - 1)), off: 0 };
      };
      const rows = children(tbl, 'w:tr');
      switch (op) {
        case 'row-above': case 'row-below': {
          const nr = tr.cloneNode(false) as Element;
          const trPr = child(tr, 'w:trPr');
          if (trPr) nr.appendChild(trPr.cloneNode(true));
          for (const c of children(tr, 'w:tc')) nr.appendChild(this.blankCell(c));
          if (op === 'row-above') tr.before(nr); else tr.after(nr);
          return finish(this.cellAtGrid(nr, this.gridIndex(tc)));
        }
        case 'col-left': case 'col-right': {
          const col = this.gridIndex(tc);
          const grid = child(tbl, 'w:tblGrid');
          const gc = children(grid, 'w:gridCol');
          const ref = gc[Math.min(col, gc.length - 1)];
          if (ref) { const n = ref.cloneNode(true) as Element; if (op === 'col-left') ref.before(n); else ref.after(n); }
          let focus: Element | null = null;
          for (const r of rows) {
            const c = this.cellAtGrid(r, col);
            if (!c) continue;
            const nc = this.blankCell(c);
            removeChild(child(nc, 'w:tcPr'), 'w:gridSpan');
            if (op === 'col-left') c.before(nc); else c.after(nc);
            if (r === tr) focus = nc;
          }
          return finish(focus);
        }
        case 'del-row': {
          if (rows.length <= 1) return this.deleteTable(tbl, finish);
          const next = tr.nextElementSibling && is(tr.nextElementSibling, 'w:tr') ? tr.nextElementSibling : tr.previousElementSibling;
          tr.remove();
          return finish(next);
        }
        case 'del-col': {
          const col = this.gridIndex(tc);
          const gc = children(child(tbl, 'w:tblGrid'), 'w:gridCol');
          if (gc.length <= 1) return this.deleteTable(tbl, finish);
          gc[Math.min(col, gc.length - 1)]?.remove();
          for (const r of rows) {
            const c = this.cellAtGrid(r, col);
            if (!c) continue;
            const span = path(c, 'w:tcPr', 'w:gridSpan');
            if (span && (wnum(span, 'val') ?? 1) > 1) setAttr(span, 'w:val', (wnum(span, 'val') ?? 2) - 1);
            else if (children(r, 'w:tc').length > 1) c.remove();
          }
          return finish(this.cellAtGrid(tr, Math.max(0, col - 1)) ?? this.cellAtGrid(tr, 0));
        }
        case 'del-table': return this.deleteTable(tbl, finish);
      }
    });
  }

  /** Change the case of the text in a range (run structure and formatting are kept). */
  changeCase(range: Range, mode: 'upper' | 'lower' | 'title' | 'sentence') {
    this.transact('대/소문자 바꾸기', [], () => {
      this.eachInRange(range, (p, s, e) => {
        if (s >= e) return;
        let first = true;
        let prevSpace = s === 0 || /\s/.test(paraText(p)[s - 1] ?? ' ');
        for (const it of itemsInRange(p, s, e)) {
          if (it.kind !== 'text') { prevSpace = it.kind !== 'atom'; continue; }
          const t = child(it.nodes[0], 'w:t');
          if (!t) continue;
          let out = '';
          for (const ch of t.textContent ?? '') {
            if (mode === 'upper') out += ch.toUpperCase();
            else if (mode === 'lower') out += ch.toLowerCase();
            else if (mode === 'title') out += prevSpace ? ch.toUpperCase() : ch.toLowerCase();
            else out += first && /\S/.test(ch) ? ch.toUpperCase() : ch.toLowerCase();
            if (/\S/.test(ch)) first = false;
            prevSpace = /\s/.test(ch);
          }
          setText(t, out);
        }
        mergeRuns(p);
      });
    });
  }

  /** Make the selected text (within one paragraph) a hyperlink; with no selection insert the URL. */
  insertHyperlink(range: Range, url: string, text?: string): Pos {
    const { start, end } = normRange(range);
    const p = this.para(start.pi)!;
    const e = start.pi === end.pi ? end.off : paraLength(p);
    return this.transact('하이퍼링크', [relsPathFor(this.docPath)], () => {
      const internal = url.startsWith('#');
      const link = create(this.xml, 'w:hyperlink');
      if (internal) setAttr(link, 'w:anchor', url.slice(1));
      else link.setAttributeNS(NS.r, 'r:id', this.pkg.addExternalRel(this.docPath, REL.hyperlink, url));
      setAttr(link, 'w:history', '1');
      const styled = (rPr: Element | null) => {
        const r = rPr ? rPr.cloneNode(true) as Element : create(this.xml, 'w:rPr');
        if (this.styles.style('Hyperlink')) setAttr(ensureOrdered(r, 'w:rStyle', RPR_ORDER), 'w:val', 'Hyperlink');
        else setRunProps(r, { color: '0563C1', u: 'single' });
        return r;
      };
      let len: number;
      if (start.off === e) {
        const label = text || url;
        link.appendChild(makeRun(this.xml, styled(this.rPrAt(p, start.off)), label));
        insertAt(p, start.off, [link]);
        len = label.length;
      } else {
        const items = itemsInRange(p, start.off, e).filter((it) => !it.containers.length);
        if (!items.length) return start;
        items[0].nodes[0].before(link);
        for (const it of items) for (const n of it.nodes) {
          if (is(n, 'w:r')) { const old = child(n, 'w:rPr'); const r = styled(old); if (old) old.replaceWith(r); else n.insertBefore(r, n.firstChild); }
          link.appendChild(n);
        }
        len = e - start.off;
      }
      mergeRuns(p);
      return { pi: start.pi, off: start.off + len };
    });
  }

  /** Line break (Shift+Enter) at a position. */
  insertLineBreak(pos: Pos): Pos {
    return this.transact('줄 바꿈', [], () => {
      const p = this.para(pos.pi)!;
      insertAt(p, pos.off, [makeRun(this.xml, this.rPrAt(p, pos.off), create(this.xml, 'w:br'))]);
      return { pi: pos.pi, off: pos.off + 1 };
    });
  }

  /** Number of tracked changes (insertions, deletions, moves, formatting changes). */
  revisionCount(): number {
    let n = 0;
    for (const l of ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange']) n += this.body.getElementsByTagNameNS(NS.w, l).length;
    return n;
  }

  /** Accept (or reject) every tracked change in the body. */
  resolveRevisions(accept: boolean) {
    this.transact(accept ? '변경 내용 모두 적용' : '변경 내용 모두 거부', [], () => {
      const all = (l: string) => Array.from(this.body.getElementsByTagNameNS(NS.w, l));
      const unwrap = (el: Element) => { while (el.firstChild) el.before(el.firstChild); el.remove(); };
      for (const el of all('ins')) {
        if (is(el.parentElement, 'w:rPr') || is(el.parentElement, 'w:trPr')) { el.remove(); continue; } // paragraph-mark / row markers
        if (accept) unwrap(el); else el.remove();
      }
      for (const el of all('moveTo')) { if (accept) unwrap(el); else el.remove(); }
      for (const tag of ['del', 'moveFrom']) {
        for (const el of all(tag)) {
          if (is(el.parentElement, 'w:rPr') || is(el.parentElement, 'w:trPr')) { el.remove(); continue; }
          if (accept) { el.remove(); continue; }
          for (const dt of Array.from(el.getElementsByTagNameNS(NS.w, 'delText'))) {
            const t = create(this.xml, 'w:t');
            setText(t, dt.textContent ?? '');
            dt.replaceWith(t);
          }
          unwrap(el);
        }
      }
      for (const tag of ['rPrChange', 'pPrChange']) {
        for (const el of all(tag)) {
          if (!accept) {
            // Reject: the properties recorded before the change come back.
            const old = el.firstElementChild;
            const owner = el.parentElement!;
            for (const c of children(owner)) if (c !== el) c.remove();
            if (old) for (const c of children(old)) owner.appendChild(c);
          }
          el.remove();
        }
      }
      for (const l of ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']) for (const el of all(l)) el.remove();
      for (const p of walkParagraphs(this.body)) mergeRuns(p);
    });
  }

  /** Paragraph indices inside the cells that contain `pis`. */
  cellParagraphs(pis: number[]): number[] {
    const cells = new Set(pis.map((i) => this.cellOf(i)?.tc).filter((c): c is Element => !!c));
    const out: number[] = [];
    for (const tc of cells) for (const p of children(tc, 'w:p')) out.push(this.indexOf(p));
    return out.filter((i) => i >= 0);
  }

  /** Word's 3×3 cell alignment: vertical (cell) and horizontal (paragraphs in the cell). */
  alignCells(pis: number[], v: 'top' | 'center' | 'bottom', h: 'left' | 'center' | 'right') {
    const paras = this.cellParagraphs(pis);
    if (!paras.length) return;
    this.transact('셀 맞춤', [], () => {
      this.setCellVAlign(pis, v);
      for (const i of paras) setParaProps(this.para(i)!, { jc: h });
    });
  }

  /** Vertical alignment of the cells that contain paragraphs `pis` (w:tcPr/w:vAlign). */
  setCellVAlign(pis: number[], v: 'top' | 'center' | 'bottom') {
    const cells = [...new Set(pis.map((i) => this.cellOf(i)?.tc).filter((c): c is Element => !!c))];
    if (!cells.length) return;
    this.transact('셀 세로 정렬', [], () => {
      for (const tc of cells) {
        let tcPr = child(tc, 'w:tcPr');
        if (!tcPr) { tcPr = create(this.xml, 'w:tcPr'); tc.insertBefore(tcPr, tc.firstChild); }
        if (v === 'top') removeChild(tcPr, 'w:vAlign');
        else setAttr(ensureOrdered(tcPr, 'w:vAlign', TCPR_ORDER), 'w:val', v);
      }
    });
  }

  cellPrXml(pi: number): string {
    const pr = path(this.cellOf(pi)?.tc, 'w:tcPr');
    return pr ? new XMLSerializer().serializeToString(pr) : '';
  }

  cellVAlign(pi: number): 'top' | 'center' | 'bottom' | null {
    const tc = this.cellOf(pi)?.tc;
    if (!tc) return null;
    const v = wa(path(tc, 'w:tcPr', 'w:vAlign'), 'val');
    return v === 'center' || v === 'bottom' ? v : 'top';
  }

  /** Grid column widths (twips) of a table. */
  columnWidths(tbl: Element): number[] {
    return children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((g) => wnum(g, 'w') ?? 0);
  }

  /** Set grid column widths (twips): updates w:gridCol, every cell's w:tcW and a fixed table width. */
  setColumnWidths(tbl: Element, widths: number[]) {
    this.transact('열 너비', [], () => {
      const ws = widths.map((w) => Math.max(120, Math.round(w)));
      let grid = child(tbl, 'w:tblGrid');
      if (!grid) { grid = create(this.xml, 'w:tblGrid'); child(tbl, 'w:tblPr')!.after(grid); }
      for (const g of children(grid, 'w:gridCol')) g.remove();
      for (const w of ws) grid.appendChild(create(this.xml, 'w:gridCol', { 'w:w': w }));
      for (const tr of children(tbl, 'w:tr')) {
        let col = wnum(path(tr, 'w:trPr', 'w:gridBefore'), 'val') ?? 0;
        for (const tc of children(tr, 'w:tc')) {
          let tcPr = child(tc, 'w:tcPr');
          if (!tcPr) { tcPr = create(this.xml, 'w:tcPr'); tc.insertBefore(tcPr, tc.firstChild); }
          const span = wnum(child(tcPr, 'w:gridSpan'), 'val') ?? 1;
          const w = ws.slice(col, col + span).reduce((a, b) => a + b, 0);
          const tcW = ensureOrdered(tcPr, 'w:tcW', TCPR_ORDER);
          setAttr(tcW, 'w:w', w);
          setAttr(tcW, 'w:type', 'dxa');
          col += span;
        }
      }
      const tblPr = child(tbl, 'w:tblPr') ?? (() => { const e = create(this.xml, 'w:tblPr'); tbl.insertBefore(e, tbl.firstChild); return e; })();
      const tblW = ensureOrdered(tblPr, 'w:tblW', TBLPR_ORDER);
      setAttr(tblW, 'w:w', ws.reduce((a, b) => a + b, 0));
      setAttr(tblW, 'w:type', 'dxa');
      // Dragged widths are exact: stop Word from autofitting them back.
      setAttr(ensureOrdered(tblPr, 'w:tblLayout', TBLPR_ORDER), 'w:type', 'fixed');
    });
  }

  /** Minimum height of row `row` (twips; null clears it). */
  setRowHeight(tbl: Element, row: number, twips: number | null) {
    const tr = children(tbl, 'w:tr')[row];
    if (!tr) return;
    this.transact('행 높이', [], () => {
      let trPr = child(tr, 'w:trPr');
      if (twips == null) { removeChild(trPr, 'w:trHeight'); return; }
      if (!trPr) {
        trPr = create(this.xml, 'w:trPr');
        const ex = child(tr, 'w:tblPrEx');
        tr.insertBefore(trPr, ex ? ex.nextSibling : tr.firstChild);
      }
      const h = ensureOrdered(trPr, 'w:trHeight', TRPR_ORDER);
      setAttr(h, 'w:val', Math.max(60, Math.round(twips)));
      setAttr(h, 'w:hRule', 'atLeast');
    });
  }

  private deleteTable(tbl: Element, finish: (f: Element | null) => Pos): Pos {
    let next = tbl.nextElementSibling;
    if (!next || !is(next, 'w:p')) {
      const np = create(this.xml, 'w:p');
      tbl.after(np);
      next = np;
    }
    tbl.remove();
    return finish(next);
  }

  // ---------------------------------------------------------------- images

  insertImage(pos: Pos, bytes: Uint8Array, mime: string, wPt: number, hPt: number, name = '그림'): Pos {
    const ext = ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg' } as Record<string, string>)[mime] ?? 'png';
    return this.transact('그림 넣기', [relsPathFor(this.docPath), '[Content_Types].xml'], () => {
      const media = this.pkg.uniqueName((n) => `word/media/image${n}.${ext}`);
      this.pkg.setBinary(media, bytes);
      this.pkg.ensureDefault(ext, mime === 'image/jpg' ? 'image/jpeg' : mime);
      const rId = this.pkg.addRel(this.docPath, REL.image, media);
      const ids = descendants(this.xml, 'wp:docPr').map((d) => Number(d.getAttribute('id')) || 0);
      const id = Math.max(0, ...ids) + 1;
      const cx = Math.round(wPt * EMU_PER_PT), cy = Math.round(hPt * EMU_PER_PT);
      const xml = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${escapeXml(name)} ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${escapeXml(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      const [run] = fragment(this.xml, xml);
      const p = this.para(pos.pi)!;
      const rPr = this.rPrAt(p, pos.off);
      if (rPr) run.insertBefore(rPr.cloneNode(true), run.firstChild);
      insertAt(p, pos.off, [run]);
      return { pi: pos.pi, off: pos.off + 1 };
    });
  }

  /** Resize the inline image item at `pos` (offset of the image). */
  resizeImage(pos: Pos, wPt: number, hPt: number) {
    const p = this.para(pos.pi);
    if (!p) return;
    this.transact('그림 크기', [], () => {
      let at = 0;
      for (const it of paraItems(p)) {
        if (at === pos.off && it.kind === 'atom') {
          const d = it.obj && is(it.obj, 'w:drawing') ? it.obj : null;
          if (!d) return;
          const cx = String(Math.round(wPt * EMU_PER_PT)), cy = String(Math.round(hPt * EMU_PER_PT));
          for (const e of [...descendants(d, 'wp:extent'), ...descendants(d, 'a:ext')]) { e.setAttribute('cx', cx); e.setAttribute('cy', cy); }
          return;
        }
        at += it.len;
      }
    });
  }

  // ---------------------------------------------------------------- find / replace

  findAll(q: string, matchCase: boolean): { pi: number; off: number; len: number }[] {
    if (!q) return [];
    const out: { pi: number; off: number; len: number }[] = [];
    const needle = matchCase ? q : q.toLowerCase();
    this.paragraphs().forEach((p, pi) => {
      const t = matchCase ? paraText(p) : paraText(p).toLowerCase();
      for (let i = t.indexOf(needle); i >= 0; i = t.indexOf(needle, i + Math.max(1, needle.length))) out.push({ pi, off: i, len: q.length });
    });
    return out;
  }

  replaceAll(q: string, r: string, matchCase: boolean): number {
    const hits = this.findAll(q, matchCase).filter((h) => {
      const t = paraText(this.para(h.pi)!).slice(h.off, h.off + h.len);
      return !t.includes('￼');
    });
    if (!hits.length) return 0;
    this.transact('모두 바꾸기', [], () => {
      const paras = this.paragraphs();
      // Right to left so earlier offsets stay valid.
      for (const h of [...hits].reverse()) this.replaceIn(paras[h.pi], h.off, h.off + h.len, r);
    });
    return hits.length;
  }

  replaceAt(hit: { pi: number; off: number; len: number }, r: string) {
    this.transact('바꾸기', [], () => this.replaceIn(this.para(hit.pi)!, hit.off, hit.off + hit.len, r));
  }

  // ---------------------------------------------------------------- info

  outline(): { pi: number; level: number; text: string }[] {
    const out: { pi: number; level: number; text: string }[] = [];
    this.paragraphs().forEach((p, pi) => {
      if (is(p.parentElement, 'w:tc')) return;
      const lvl = this.paraFmt(p).outlineLvl;
      if (lvl == null || lvl > 8) return;
      const text = paraText(p).replace(/[￼\t\n]+/g, ' ').trim();
      if (text) out.push({ pi, level: lvl, text });
    });
    return out;
  }

  stats() {
    const texts = this.paragraphs().map((p) => paraText(p).replace(/￼/g, ''));
    const all = texts.join('\n');
    return {
      paragraphs: texts.filter((t) => t.trim()).length,
      chars: all.replace(/\n/g, '').length,
      charsNoSpace: all.replace(/\s/g, '').length,
      words: (all.match(/\S+/g) ?? []).length,
    };
  }

  coreProps() {
    const core = this.pkg.getXml('docProps/core.xml');
    const get = (n: string) => core?.getElementsByTagNameNS('*', n)[0]?.textContent ?? '';
    return { title: get('title'), creator: get('creator'), modified: get('modified'), created: get('created') };
  }

  /** Text of a comment (word/comments.xml), for tooltips. */
  commentText(id: string): string {
    const rel = this.pkg.rels(this.docPath).find((r) => r.type.endsWith('/comments'));
    const doc = rel ? this.pkg.getXml(rel.target) : null;
    const c = doc ? children(doc.documentElement, 'w:comment').find((x) => wa(x, 'id') === id) : null;
    if (!c) return '메모';
    return `${wa(c, 'author') ?? ''}: ${children(c, 'w:p').map((p) => paraText(p)).join('\n')}`.trim();
  }

  /** Images referenced by a part (rels of type image), for preloading. */
  imageTargets(part: string) {
    return this.pkg.rels(part).filter((r) => r.type === REL.image && !r.external);
  }

  /** Paragraph items (exposed for the renderer and the editor). */
  items(p: Element): Item[] { return paraItems(p); }
  textOf(p: Element) { return paraText(p); }
  lengthOf(p: Element) { return paraLength(p); }

  /** Re-serialize the main part (handy for tests). */
  xmlText() { return serializeXml(this.xml); }
}
