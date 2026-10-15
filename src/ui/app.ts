// Application controller: document lifecycle, command registration, UI state sync.
import { getCurrentWindow } from '@tauri-apps/api/window';
import { DocxDocument, Pos, Range, TWIP_PER_CM, comparePos, normRange } from '../engine/document';
import { RunProps } from '../engine/model';
import { ImageCache } from '../engine/render';
import { ParaFmt, RunFmt, styleLabel } from '../engine/styles';
import { fontStack } from '../engine/render';
import { Commands } from './commands';
import { closeContextMenu, closePopover, colorPopover, contextMenu, dateDialog, dialog, emojiPopover, fontDialog, highlightPopover, html, hyperlinkDialog, menuPopover, pageSetupDialog, paragraphDialog, symbolPopover, tablePopover, toast, windowMenuItems, zoomDialog } from './dialogs';
import { DocEditor } from './editor';
import { installIcons } from './icons';
import { printDocument } from './print';
import { drawRulers } from './ruler';
import { focusWindow, isDesktop, listWindows, openNewWindow, pickImage, pickOpenFile, saveFile } from './shell';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
/** A popover anchor: the ribbon button when it is visible, else the menu bar. */
const anchor = (id: string) => { const e = $(id); return e && e.offsetParent ? e : $('menu-bar'); };

const HF_ALLOWED = /^(file:|view:|tool:|hf:|window:|edit:(undo|redo|cut|copy|paste|paste-text|select-all)$|format:(bold|italic|underline|strike|superscript|subscript|grow|shrink|text-color|highlight|font|clear|align-\w+)$|insert:(date|symbol|emoji)$|page:setup$)/;
const FONTS = ['맑은 고딕', 'Apple SD Gothic Neo', '나눔고딕', '나눔명조', '굴림', '돋움', '바탕', '궁서', 'Noto Sans KR', 'Pretendard', 'Arial', 'Calibri', 'Calibri Light', 'Cambria', 'Helvetica', 'Times New Roman', 'Georgia', 'Verdana', 'Tahoma', 'Segoe UI', 'Consolas', 'Courier New'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const ZOOMS = [0.25, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 3, 4];
const PX = 4 / 3;
const HANGUL = /[ᄀ-ᇿ㄰-㆏가-힯　-鿿＀-￯]/;

interface Clip { text: string; paras: string[] }

export class App {
  doc: DocxDocument | null = null;
  images!: ImageCache;
  fileName = '새 문서.docx';
  fileHandle: FileSystemFileHandle | undefined;
  filePath: string | undefined; // desktop shell
  fit: 'width' | 'page' | 'text' | null = null;
  showRuler = true;
  showOutline = true;
  showRibbon = true;
  showGridlines = true;
  clipboard: Clip | null = null;
  lastTextColor = 'C00000';
  lastHighlight = 'yellow';
  readonly cmd = new Commands();
  readonly editor: DocEditor;
  private syncQueued = false;
  private infoTimer = 0;

  constructor() {
    this.editor = new DocEditor({
      doc: () => this.doc,
      images: () => this.images,
      selectionChanged: () => this.queueSync(),
      hfChanged: () => this.onHFChanged(),
      layoutChanged: () => { this.syncStatus(); this.syncRulers(); },
      status: (m) => this.message(m),
    }, $('page-wrap'));
    this.cmd.onError = (e) => { console.error(e); toast('오류: ' + ((e as Error).message ?? e)); };
    // While a header/footer is being edited only commands that make sense there are available.
    this.cmd.guard = (id) => !this.editor.hf || HF_ALLOWED.test(id);
    this.registerCommands();
    this.bindUi();
  }

  // ------------------------------------------------------------------ document lifecycle

  async open(bytes: Uint8Array, name: string, handle?: FileSystemFileHandle, path?: string) {
    const doc = await DocxDocument.load(bytes);
    this.images?.dispose();
    this.doc = doc;
    this.images = new ImageCache(doc);
    await this.images.ensureAll();
    this.fileName = name;
    this.fileHandle = handle;
    this.filePath = path;
    doc.selection = () => this.editor.getRange();
    doc.onChange = () => this.onDocChange();
    this.editor.render();
    this.buildStyleList();
    this.zoomDefault();
    this.updateTitle();
    this.updateOutline();
    this.updateInfo();
    this.editor.setRange({ start: { pi: 0, off: 0 }, end: { pi: 0, off: 0 } });
    $('scroll-container').scrollTop = 0;
    this.message(`${name} 열기 완료 — ${this.editor.pages}쪽`);
  }

  async newDocument() {
    const res = await fetch('./blank.docx');
    const bytes = new Uint8Array(await res.arrayBuffer());
    await this.open(bytes, '새 문서.docx');
    this.message('새 문서');
  }

  private onDocChange() {
    this.updateTitle();
    clearTimeout(this.infoTimer);
    this.infoTimer = window.setTimeout(() => { this.updateOutline(); this.updateInfo(); }, 400);
  }

  private updateTitle() {
    document.title = `${this.doc?.dirty ? '● ' : ''}${this.fileName} — Oword`;
    if (isDesktop()) getCurrentWindow().setTitle(document.title).catch(() => {});
  }

  message(m: string) { $('sb-message').textContent = m; }

  // ------------------------------------------------------------------ zoom

  setZoom(z: number, fit: App['fit'] = null) {
    this.fit = fit;
    this.editor.setZoom(z);
    this.syncStatus();
    this.syncRulers();
  }

  private fitZoom(kind: 'width' | 'page' | 'text') {
    if (!this.doc) return 1;
    const sc = $('scroll-container');
    const g = this.editor.geometry();
    const zw = (sc.clientWidth - 48) / (g.W * PX);
    const zh = (sc.clientHeight - 32) / (g.H * PX);
    // Text width: the text column (page minus margins) fills the window.
    const zt = (sc.clientWidth - 48) / ((g.CW + 24) * PX);
    return Math.max(0.2, kind === 'width' ? zw : kind === 'text' ? zt : Math.min(zw, zh));
  }

  private zoomDefault() {
    const zw = this.fitZoom('width');
    if (zw < 1) this.setZoom(zw, 'width'); else this.setZoom(1);
  }

  private stepZoom(dir: 1 | -1) {
    const z = this.editor.zoom;
    const next = dir > 0 ? ZOOMS.find((v) => v > z + 0.001) ?? 5 : [...ZOOMS].reverse().find((v) => v < z - 0.001) ?? 0.2;
    this.setZoom(next);
  }

  // ------------------------------------------------------------------ status / style bar

  private queueSync() {
    if (this.syncQueued) return;
    this.syncQueued = true;
    requestAnimationFrame(() => { this.syncQueued = false; this.syncStyleBar(); this.syncRulers(); this.syncStatus(); this.syncRibbon(); });
  }

  /** Enable state of ribbon buttons and the view checkboxes. */
  syncRibbon() {
    document.querySelectorAll<HTMLElement>('#ribbon [data-cmd]').forEach((b) => b.classList.toggle('disabled', !this.cmd.enabled(b.dataset.cmd!)));
    document.querySelectorAll<HTMLElement>('[data-toggle]').forEach((b) => b.classList.toggle('active', this.cmd.checked(b.dataset.cmd!)));
    document.querySelectorAll<HTMLInputElement>('[data-check]').forEach((cb) => { cb.checked = this.cmd.checked(cb.dataset.check!); });
    $('tb-marks').classList.toggle('active', this.editor.showMarks);
  }

  syncStatus() {
    const n = this.editor.pages;
    $('sb-page').textContent = !this.doc ? '페이지 0/0' : this.editor.mode === 'web' ? '웹 모양' : `페이지 ${Math.min(n, this.editor.caretPage())}/${n}`;
    $('sb-zoom-val').textContent = Math.round(this.editor.zoom * 100) + '%';
    const zs = $<HTMLInputElement>('zoom-slider');
    if (document.activeElement !== zs) zs.value = String(Math.round(this.editor.zoom * 100));
  }

  private updateInfo() {
    if (!this.doc) return;
    const s = this.doc.stats();
    $('sb-count').textContent = `${s.words.toLocaleString()} 단어 · ${s.chars.toLocaleString()} 글자`;
  }

  /** Character formatting at the caret: the pending span, the run under the caret, or the model. */
  private caretFmt(): { fmt: RunFmt; pi: number } | null {
    const doc = this.doc;
    const hc = this.editor.hfCaret();
    if (doc && hc) return { fmt: doc.runFmt(hc.p, hc.rPr), pi: -1 };
    const r = this.editor.getRange();
    if (!doc || !r) return null;
    const pos = normRange(r).start;
    const p = doc.para(pos.pi);
    if (!p) return null;
    const pend = this.editor.pendingRPr();
    if (pend !== undefined) return { fmt: doc.runFmt(p, pend), pi: pos.pi };
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    const span = node && (node.nodeType === 3 ? node.parentElement : node as HTMLElement)?.closest?.('[data-r]') as HTMLElement | null;
    if (span && this.editor.root.contains(span)) {
      const rPr = this.editor.reg.runs[Number(span.dataset.r)];
      if (rPr !== undefined) return { fmt: doc.runFmt(p, rPr), pi: pos.pi };
    }
    return { fmt: doc.fmtAt(pos) ?? doc.markFmt(p), pi: pos.pi };
  }

  private charBeforeCaret(): string {
    const sel = window.getSelection();
    const n = sel?.anchorNode;
    if (n?.nodeType === 3) return (n.textContent ?? '').slice(Math.max(0, sel!.anchorOffset - 1), sel!.anchorOffset) || (n.textContent ?? '')[0] || '';
    return n?.textContent?.[0] ?? '';
  }

  syncStyleBar() {
    const doc = this.doc;
    const st = this.caretFmt();
    const f = st?.fmt;
    const fn = $<HTMLInputElement>('font-name'), fs = $<HTMLInputElement>('font-size');
    const ea = HANGUL.test(this.charBeforeCaret());
    if (document.activeElement !== fn) fn.value = f ? (ea ? f.ea ?? f.ascii : f.ascii ?? f.ea) ?? '' : '';
    if (document.activeElement !== fs) fs.value = f ? String((f.sz ?? 20) / 2) : '';
    $('btn-bold').classList.toggle('active', !!f?.b);
    $('btn-italic').classList.toggle('active', !!f?.i);
    $('btn-underline').classList.toggle('active', !!f?.u && f.u !== 'none');
    $('btn-strike').classList.toggle('active', !!f?.strike);
    $('btn-sup').classList.toggle('active', f?.vertAlign === 'superscript');
    $('btn-sub').classList.toggle('active', f?.vertAlign === 'subscript');
    const p = st && doc ? doc.para(st.pi) : null;
    const pf = p ? doc!.paraFmt(p) : null;
    const jc = pf?.jc === 'start' ? 'left' : pf?.jc === 'end' ? 'right' : pf?.jc ?? 'left';
    for (const [id, a] of [['btn-align-left', 'left'], ['btn-align-center', 'center'], ['btn-align-right', 'right'], ['btn-align-justify', 'both']]) {
      $(id).classList.toggle('active', !!pf && (jc === a || (a === 'both' && jc === 'distribute')));
    }
    if (p) {
      const id = doc!.pStyleOf(p) ?? doc!.styles.defaultPara ?? '';
      const gal = $('style-gallery');
      if (!gal.querySelector(`[data-style="${CSS.escape(id)}"]`)) this.buildStyleList(id);
      gal.querySelectorAll<HTMLElement>('.style-card').forEach((c) => c.classList.toggle('active', c.dataset.style === id));
    }
    const ls = $<HTMLSelectElement>('line-spacing');
    if (pf && document.activeElement !== ls) {
      const v = pf.line != null && (!pf.lineRule || pf.lineRule === 'auto') ? (pf.line / 240).toFixed(2) : '';
      ls.value = [...ls.options].some((o) => o.value === v) ? v : '';
    }
    const inList = !!(p && doc!.labels().get(p));
    $('tb-bullet').classList.toggle('active', inList && !!pf?.numId && doc!.numbering.isBullet(pf.numId, pf.ilvl ?? 0));
    $('tb-numbering').classList.toggle('active', inList && !!pf?.numId && !doc!.numbering.isBullet(pf.numId, pf.ilvl ?? 0));
    const cell = st && doc?.cellOf(st.pi);
    const va = st && doc ? doc.cellVAlign(st.pi) : null;
    for (const v of ['top', 'center', 'bottom']) {
      const b = $(`btn-valign-${v}`);
      b.classList.toggle('active', va === v);
      b.classList.toggle('disabled', !va);
      for (const h of ['left', 'center', 'right']) $(`btn-cell-${v}-${h}`).classList.toggle('active', va === v && (jc === h || (h === 'left' && jc === 'both')));
    }
    // The contextual "표 레이아웃" tab exists only while the caret is in a table (as in Word).
    const tabBtn = document.querySelector<HTMLElement>('.rb-tab[data-tab="table"]')!;
    if (tabBtn.hidden === !!va) {
      tabBtn.hidden = !va;
      if (!va && tabBtn.classList.contains('active')) this.showTab('home');
    }
    this.syncLayoutFields(pf);
    $('hl-bar').style.background = '#' + ({ yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000' } as Record<string, string>)[this.lastHighlight];
    $('sb-mode').textContent = cell ? '표' : this.editor.collapsed() ? '입력' : '선택';
    this.markOutline(st?.pi ?? -1);
  }

  private syncLayoutFields(pf: ParaFmt | null) {
    const set = (id: string, v: number | undefined, k: number) => { const e = $<HTMLInputElement>(id); if (document.activeElement !== e) e.value = pf ? String(+((v ?? 0) / k).toFixed(2)) : ''; };
    set('lay-left', pf?.left, TWIP_PER_CM);
    set('lay-right', pf?.right, TWIP_PER_CM);
    set('lay-before', pf?.before, 20);
    set('lay-after', pf?.after, 20);
  }

  showTab(tab: string) {
    document.querySelectorAll<HTMLElement>('.rb-tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === tab));
    document.querySelectorAll<HTMLElement>('.rb-panel').forEach((p) => p.classList.toggle('active', p.dataset.panel === tab));
    if (!this.showRibbon) this.cmd.run('view:toggle-ribbon');
  }

  /** Style gallery in the Home tab: each card previews its style. */
  private buildStyleList(extra?: string) {
    const doc = this.doc;
    const gal = $('style-gallery');
    gal.innerHTML = '';
    if (!doc) return;
    const used = new Set(doc.paragraphs().map((p) => doc.pStyleOf(p)).filter((s): s is string => !!s));
    if (extra) used.add(extra);
    if (doc.styles.defaultPara) used.add(doc.styles.defaultPara);
    const list = doc.styles.paragraphStyles(used);
    // Offer the headings even when the document has no definition yet (added on first use).
    const ids = new Set(list.map((s) => s.id));
    const extraBuiltins = ['Heading1', 'Heading2', 'Heading3', 'Title'].filter((i) => !ids.has(i) && ![...doc.styles.styles.values()].some((s) => s.name.toLowerCase() === (i === 'Title' ? 'title' : 'heading ' + i.slice(-1))));
    const cards: [string, string][] = [...list.map((s) => [s.id, styleLabel(s)] as [string, string]), ...extraBuiltins.map((i) => [i, i === 'Title' ? '제목' : '제목 ' + i.slice(-1)] as [string, string])];
    const order = (id: string) => { const i = [doc.styles.defaultPara, 'NoSpacing', 'Heading1', 'Heading2', 'Heading3', 'Title', 'Subtitle'].indexOf(id); return i < 0 ? 99 : i; };
    cards.sort((a, b) => order(a[0]) - order(b[0]));
    for (const [id, label] of cards) {
      const b = document.createElement('button');
      b.className = 'style-card';
      b.dataset.style = id;
      b.title = label;
      const prev = document.createElement('span');
      prev.className = 'style-prev';
      prev.textContent = '가나다Aa';
      const f = doc.styles.style(id) ? doc.styles.styleRunFmt(id, null) : { sz: id === 'Title' ? 56 : id === 'Heading1' ? 32 : 26, color: '2F5496' } as RunFmt;
      prev.style.fontFamily = fontStack(f);
      prev.style.fontSize = Math.min(15, Math.max(10, (f.sz ?? 20) / 2 * 0.55 + 4)) + 'pt';
      prev.style.fontWeight = f.b ? 'bold' : 'normal';
      prev.style.fontStyle = f.i ? 'italic' : 'normal';
      prev.style.color = f.color && f.color !== 'auto' ? '#' + f.color : '#000';
      const name = document.createElement('span');
      name.className = 'style-name';
      name.textContent = label;
      b.append(prev, name);
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => this.applyStyle(id));
      gal.appendChild(b);
    }
  }

  syncRulers() {
    $('editor-area').classList.toggle('no-ruler', !this.showRuler || this.editor.mode === 'web');
    if (!this.doc || !this.showRuler || this.editor.mode === 'web') return;
    const g = this.editor.geometry();
    const st = this.caretFmt();
    const p = st ? this.doc.para(st.pi) : null;
    const pf = p ? this.doc.paraFmt(p) : null;
    const page = this.editor.caretPage() - 1;
    drawRulers($<HTMLCanvasElement>('h-ruler'), $<HTMLCanvasElement>('v-ruler'), this.editor.host, this.editor.zoom, {
      pageW: g.W, pageH: g.H, left: g.left, right: g.right, top: g.top, bottom: g.bottom,
      indLeft: (pf?.left ?? 0) / 20, indRight: (pf?.right ?? 0) / 20, first: ((pf?.firstLine ?? 0) - (pf?.hanging ?? 0)) / 20,
      pageTop: page * (g.H + this.editor.gapPt) * PX,
    });
  }

  // ------------------------------------------------------------------ outline (탐색 창)

  private updateOutline() {
    const list = $('outline-list');
    const doc = this.doc;
    if (!doc) { list.innerHTML = ''; return; }
    const items = doc.outline();
    if (!items.length) { list.innerHTML = '<div class="outline-empty">제목 스타일(제목 1, 제목 2…)을 쓴 문단이 여기에 나타납니다.</div>'; return; }
    list.innerHTML = '';
    for (const h of items) {
      const d = document.createElement('div');
      d.className = 'outline-item';
      d.style.paddingLeft = 8 + h.level * 14 + 'px';
      d.textContent = h.text;
      d.title = h.text;
      d.dataset.pi = String(h.pi);
      d.onclick = () => this.editor.scrollToPara(h.pi);
      list.appendChild(d);
    }
  }

  private markOutline(pi: number) {
    let cur: HTMLElement | null = null;
    for (const d of Array.from($('outline-list').children) as HTMLElement[]) {
      if (d.dataset.pi && Number(d.dataset.pi) <= pi) cur = d;
      d.classList.remove('current');
    }
    cur?.classList.add('current');
  }

  // ------------------------------------------------------------------ helpers for commands

  private hasDoc = () => !!this.doc;
  private inTable = () => { const r = this.editor.getRange(); return !!this.doc && !!r && !!this.doc.cellOf(r.start.pi); };

  private runProps(props: RunProps, label?: string) {
    if (!this.doc) return;
    if (this.editor.hf) { this.editor.hfFormat({ run: props }); this.queueSync(); return; }
    this.editor.applyRunProps(props, label);
    this.queueSync();
  }

  private toggleChar(key: 'b' | 'i' | 'strike') {
    const f = this.caretFmt()?.fmt;
    this.runProps({ [key]: !(f?.[key] ?? false) } as RunProps);
  }

  private toggleVert(v: 'superscript' | 'subscript') {
    const f = this.caretFmt()?.fmt;
    this.runProps({ vertAlign: f?.vertAlign === v ? 'baseline' : v });
  }

  private stepSize(dir: 1 | -1) {
    const cur = (this.caretFmt()?.fmt.sz ?? 20) / 2;
    const next = dir > 0 ? SIZES.find((s) => s > cur) ?? cur + 10 : [...SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
    this.runProps({ sz: next }, '글자 크기');
  }

  private para(fn: (pis: number[]) => void) {
    if (!this.doc) return;
    this.editor.paraCommand(fn);
    this.queueSync();
  }

  /** Commit, run a structural command that returns the new caret, then re-render. */
  private structural(fn: (pos: Pos, range: Range) => Pos | null | void) {
    if (!this.doc) return;
    this.editor.commit();
    const r = this.editor.getRange() ?? { start: { pi: 0, off: 0 }, end: { pi: 0, off: 0 } };
    if (comparePos(r.start, r.end) !== 0) { this.editor.focus(); this.editor.setRange(r, false); document.execCommand('delete'); this.editor.commit(); }
    const pos = normRange(this.editor.getRange() ?? r).start;
    const next = fn(pos, r);
    this.editor.render();
    if (next) this.editor.setRange({ start: next, end: next });
    this.queueSync();
  }

  private async insertImageFile(file: File | Blob, name = '그림') {
    if (!this.doc) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(file);
    const img = new Image();
    await new Promise<void>((res) => { img.onload = () => res(); img.onerror = () => res(); img.src = url; });
    URL.revokeObjectURL(url);
    let w = (img.naturalWidth || 400) * 0.75, h = (img.naturalHeight || 300) * 0.75; // 96 dpi → pt
    const g = this.editor.geometry();
    const k = Math.min(1, g.CW / w, (g.H - g.top - g.bottom) * 0.9 / h);
    w *= k; h *= k;
    const doc = this.doc;
    let target: Pos | null = null;
    this.structural((pos) => { target = doc.insertImage(pos, bytes, file.type || 'image/png', w, h, name.replace(/\.[^.]+$/, '')); return target; });
    await this.images.ensure([doc.docPath]);
    this.editor.render();
    if (target) this.editor.setRange({ start: target, end: target });
  }

  private copyToClipboard() {
    const doc = this.doc;
    const r = this.editor.getRange();
    if (!doc || !r || comparePos(r.start, r.end) === 0) return;
    this.editor.commit();
    const range = this.editor.getRange()!;
    this.clipboard = { text: window.getSelection()?.toString() ?? '', paras: doc.extractRange(range) };
  }

  private async onPaste(e: ClipboardEvent) {
    if (this.editor.hf?.el.contains(e.target as Node)) {
      // Header/footer: plain text only.
      e.preventDefault();
      const t = e.clipboardData?.getData('text/plain');
      if (t) document.execCommand('insertText', false, t.replace(/\r?\n/g, ' '));
      return;
    }
    if (!this.doc || !this.editor.isEditorTarget(e.target)) return;
    e.preventDefault();
    const items = Array.from(e.clipboardData?.items ?? []);
    const img = items.find((i) => i.type.startsWith('image/'));
    const text = e.clipboardData?.getData('text/plain') ?? '';
    const norm = (s: string) => s.replace(/[\s​]+/g, '');
    if (this.clipboard && norm(text) === norm(this.clipboard.text) && this.clipboard.paras.length) {
      const paras = this.clipboard.paras;
      this.structural((pos) => this.doc!.insertFragment(pos, paras));
      return;
    }
    if (img && !text) { const f = img.getAsFile(); if (f) await this.insertImageFile(f); return; }
    if (text) this.structural((pos) => this.doc!.insertText(pos, text, '붙이기'));
  }

  private async save(saveAs: boolean) {
    if (!this.doc) return;
    this.editor.commitHF();
    this.editor.commit();
    const blob = await this.doc.save();
    const name = this.fileName.toLowerCase().endsWith('.docx') ? this.fileName : this.fileName.replace(/\.[^.]*$/, '') + '.docx';
    const res = await saveFile(blob, name, this.fileHandle, saveAs, this.filePath);
    if (!res) { this.doc.dirty = true; return; }
    this.fileName = res.name;
    this.fileHandle = res.handle;
    this.filePath = res.path;
    this.updateTitle();
    this.message(`저장됨: ${res.name}`);
    toast('저장했습니다');
  }

  private findDialog(replace = false) {
    this.editor.commit();
    const body = html(`<div class="form" style="grid-template-columns:auto 1fr">
      <label>찾을 내용</label><input type="text" name="f">
      <label>바꿀 내용</label><input type="text" name="r">
      <label></label><label style="color:inherit"><input type="checkbox" name="mc"> 대/소문자 구분</label>
      <label></label><span id="find-result" style="font-size:12px;color:#5c6370"></span></div>`);
    const q = (n: string) => body.querySelector(`[name="${n}"]`) as HTMLInputElement;
    const sel = window.getSelection()?.toString();
    if (sel && !sel.includes('\n')) q('f').value = sel;
    if (replace && sel) setTimeout(() => q('r').focus(), 0);
    let pos = -1;
    let current: { pi: number; off: number; len: number } | null = null;
    const result = body.querySelector('#find-result')!;
    const findNext = () => {
      const hits = this.doc!.findAll(q('f').value, q('mc').checked);
      if (!hits.length) { result.textContent = '찾는 내용이 없습니다.'; current = null; return false; }
      pos = (pos + 1) % hits.length;
      current = hits[pos];
      this.editor.selectHit(current);
      result.textContent = `${hits.length}개 중 ${pos + 1}번째`;
      return false;
    };
    dialog(replace ? '바꾸기' : '찾기 및 바꾸기', body, [
      { label: '다음 찾기', primary: true, action: findNext },
      {
        label: '바꾸기', action: () => {
          if (current) {
            this.doc!.replaceAt(current, q('r').value);
            this.editor.refresh(current.pi, 1, 1);
            pos--;
          }
          return findNext();
        },
      },
      {
        label: '모두 바꾸기', action: () => {
          const n = this.doc!.replaceAll(q('f').value, q('r').value, q('mc').checked);
          if (n) this.editor.render();
          result.textContent = n ? `${n}개를 바꿨습니다.` : '찾는 내용이 없습니다.';
          return false;
        },
      },
      { label: '닫기' },
    ]);
  }

  // ------------------------------------------------------------------ commands

  private registerCommands() {
    const c = this.cmd;
    const d = () => this.doc!;
    const ed = this.editor;
    // file
    c.register('file:new', () => this.confirmDiscard() && this.newDocument(), 'ctrl+n');
    c.register('file:open', async () => {
      if (!this.confirmDiscard()) return;
      const f = await pickOpenFile($<HTMLInputElement>('file-input'));
      if (f) await this.open(f.bytes, f.name, f.handle, f.path);
    }, 'ctrl+o');
    c.register('file:save', { run: () => this.save(false), enabled: this.hasDoc }, 'ctrl+s');
    c.register('file:save-as', { run: () => this.save(true), enabled: this.hasDoc }, 'ctrl+shift+s');
    c.register('file:print', { run: () => { ed.commit(); printDocument(ed); }, enabled: this.hasDoc }, 'ctrl+p');
    c.register('file:info', { run: () => this.infoDialog(), enabled: this.hasDoc });
    c.register('file:about', () => { dialog('제품 정보', html(`<div style="line-height:1.7"><b>Oword</b> 0.2.2<br>DOCX 문서 편집기<br><span style="color:#5c6370;font-size:12px">Opptx(PPTX 편집기)와 같은 구조 — 문서 엔진 · 렌더러 · 명령 기반 UI · 데스크톱 셸 — 로 만든 Word 문서 편집기입니다. 이 구조는 HOP(HWP 편집기, MIT)을 참고했습니다.<br>MIT License · Copyright (c) 2026 moonlightpark<br>서드파티 라이선스: THIRD_PARTY_NOTICES.md<br>Microsoft와 관계없음. Microsoft Word는 Microsoft Corporation의 상표입니다.</span></div>`), [{ label: '확인', primary: true }]); });
    // edit
    c.register('edit:undo', {
      run: () => { ed.exitHF(); ed.commit(); const sel = d().undo() as Range | null; ed.render(); if (sel) ed.setRange(sel); this.buildStyleList(); this.queueSync(); },
      enabled: () => !!this.doc && (d().canUndo() || true),
    }, 'ctrl+z');
    c.register('edit:redo', {
      run: () => { ed.exitHF(); ed.commit(); const sel = d().redo() as Range | null; ed.render(); if (sel) ed.setRange(sel); this.queueSync(); },
      enabled: () => !!this.doc && d().canRedo(),
    }, 'ctrl+shift+z', 'ctrl+y');
    c.register('edit:cut', { run: () => { ed.focus(); this.copyToClipboard(); document.execCommand('cut'); }, enabled: () => !!this.doc && !ed.collapsed() }, 'ctrl+x');
    c.register('edit:copy', { run: () => { ed.focus(); this.copyToClipboard(); document.execCommand('copy'); }, enabled: () => !!this.doc && !ed.collapsed() }, 'ctrl+c');
    c.register('edit:paste', {
      run: async () => {
        // Menu paste: the async clipboard API (keyboard paste goes through the paste event).
        const text = await navigator.clipboard?.readText?.().catch(() => '') ?? '';
        const norm = (s: string) => s.replace(/[\s​]+/g, '');
        if (this.clipboard && (!text || norm(text) === norm(this.clipboard.text))) { const paras = this.clipboard.paras; this.structural((pos) => d().insertFragment(pos, paras)); return; }
        if (text) this.structural((pos) => d().insertText(pos, text, '붙이기'));
      },
      enabled: this.hasDoc,
    }, 'ctrl+v');
    c.register('edit:select-all', { run: () => { ed.focus(); document.execCommand('selectAll'); }, enabled: this.hasDoc }, 'ctrl+a');
    c.register('edit:find', { run: () => this.findDialog(), enabled: this.hasDoc }, 'ctrl+f');
    c.register('edit:replace', { run: () => this.findDialog(true), enabled: this.hasDoc }, 'ctrl+h');
    c.register('edit:paste-text', {
      run: async () => {
        const text = await navigator.clipboard?.readText?.().catch(() => '') ?? this.clipboard?.text ?? '';
        if (text) this.structural((pos) => d().insertText(pos, text, '텍스트만 붙여넣기'));
      },
      enabled: this.hasDoc,
    }, 'ctrl+shift+v');
    c.register('edit:clear-content', { run: () => { ed.focus(); document.execCommand('delete'); }, enabled: () => !!this.doc && !ed.collapsed() });
    // view
    c.register('view:zoom-in', { run: () => this.stepZoom(1), enabled: this.hasDoc });
    c.register('view:zoom-out', { run: () => this.stepZoom(-1), enabled: this.hasDoc });
    c.register('view:zoom-width', { run: () => this.setZoom(this.fitZoom('width'), 'width'), enabled: () => this.hasDoc() && ed.mode === 'print', checked: () => this.fit === 'width' });
    c.register('view:zoom-page', { run: () => this.setZoom(this.fitZoom('page'), 'page'), enabled: () => this.hasDoc() && ed.mode === 'print', checked: () => this.fit === 'page' });
    c.register('view:zoom-100', { run: () => this.setZoom(1), enabled: this.hasDoc });
    c.register('view:toggle-ruler', { run: () => { this.showRuler = !this.showRuler; $('editor-area').classList.toggle('no-ruler', !this.showRuler); this.syncRulers(); }, checked: () => this.showRuler });
    c.register('view:toggle-outline', { run: () => { this.showOutline = !this.showOutline; $('outline-panel').classList.toggle('hidden', !this.showOutline); if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); }, checked: () => this.showOutline });
    c.register('view:toggle-marks', { run: () => { ed.showMarks = !ed.showMarks; ed.paginate(); $('tb-marks').classList.toggle('active', ed.showMarks); }, checked: () => ed.showMarks }, 'ctrl+shift+8');
    // insert
    c.register('insert:table', { run: () => tablePopover(anchor('tb-table'), (r, col) => this.structural((pos) => d().insertTable(pos, r, col))), enabled: this.hasDoc });
    c.register('insert:image', {
      run: async () => { const f = await pickImage($<HTMLInputElement>('image-input')); if (f) await this.insertImageFile(f, f.name); },
      enabled: this.hasDoc,
    }, 'ctrl+shift+i');
    c.register('insert:page-break', { run: () => this.structural((pos) => d().insertPageBreak(pos)), enabled: this.hasDoc }, 'ctrl+enter');
    // format: characters
    c.register('format:bold', { run: () => this.toggleChar('b'), enabled: this.hasDoc }, 'ctrl+b');
    c.register('format:italic', { run: () => this.toggleChar('i'), enabled: this.hasDoc }, 'ctrl+i');
    c.register('format:underline', { run: () => { const f = this.caretFmt()?.fmt; this.runProps({ u: !(f?.u && f.u !== 'none') }); }, enabled: this.hasDoc }, 'ctrl+u');
    c.register('format:strike', { run: () => this.toggleChar('strike'), enabled: this.hasDoc }, 'ctrl+shift+x');
    c.register('format:superscript', { run: () => this.toggleVert('superscript'), enabled: this.hasDoc }, 'ctrl+shift+=');
    c.register('format:subscript', { run: () => this.toggleVert('subscript'), enabled: this.hasDoc }, 'ctrl+=');
    c.register('format:grow', { run: () => this.stepSize(1), enabled: this.hasDoc }, 'ctrl+]');
    c.register('format:shrink', { run: () => this.stepSize(-1), enabled: this.hasDoc }, 'ctrl+[');
    c.register('format:text-color', { run: () => colorPopover(anchor('btn-text-color'), { allowNone: '자동' }, (hex) => { if (hex) { this.lastTextColor = hex; $('color-bar').style.background = '#' + hex; } this.runProps({ color: hex }, '글자 색'); }), enabled: this.hasDoc });
    c.register('format:highlight', { run: () => highlightPopover(anchor('btn-highlight'), (name) => { if (name) this.lastHighlight = name; this.runProps({ highlight: name }, '형광펜'); }), enabled: this.hasDoc });
    c.register('format:clear', {
      run: () => {
        ed.commit();
        const r = ed.getRange();
        if (!r || comparePos(r.start, r.end) === 0) return;
        d().clearRunProps(r);
        const n = r ? Math.abs(r.end.pi - r.start.pi) + 1 : 1;
        ed.refresh(Math.min(r.start.pi, r.end.pi), n, n);
        ed.setRange(r, false);
        this.queueSync();
      },
      enabled: this.hasDoc,
    });
    // format: paragraphs
    for (const [id, jc, key] of [['left', 'left', 'ctrl+l'], ['center', 'center', 'ctrl+e'], ['right', 'right', 'ctrl+r'], ['justify', 'both', 'ctrl+j']]) {
      c.register(`format:align-${id}`, { run: () => ed.hf ? ed.hfFormat({ jc }) : this.para((pis) => d().applyParaProps(pis, { jc }, '정렬')), enabled: this.hasDoc }, key);
    }
    c.register('format:bullet', { run: () => this.para((pis) => d().toggleList(pis, 'bullet')), enabled: this.hasDoc }, 'ctrl+shift+l');
    c.register('format:numbering', { run: () => this.para((pis) => d().toggleList(pis, 'number')), enabled: this.hasDoc });
    c.register('format:indent', { run: () => this.para((pis) => d().indent(pis, 1)), enabled: this.hasDoc }, 'ctrl+m');
    c.register('format:outdent', { run: () => this.para((pis) => d().indent(pis, -1)), enabled: this.hasDoc }, 'ctrl+shift+m');
    for (const [v, key] of [['1', 'ctrl+1'], ['1.15', ''], ['1.5', 'ctrl+5'], ['2', 'ctrl+2'], ['2.5', ''], ['3', '']]) {
      c.register(`format:line-${v}`, { run: () => this.para((pis) => d().applyParaProps(pis, { line: Math.round(Number(v) * 240), lineRule: 'auto' }, '줄 간격')), enabled: this.hasDoc }, ...(key ? [key] : []));
    }
    c.register('format:paragraph-dialog-alias', { run: () => this.cmd.run('format:paragraph') }, 'ctrl+alt+m');
    c.register('format:paragraph', {
      run: () => {
        ed.commit();
        const r = ed.getRange();
        const p = r && d().para(normRange(r).start.pi);
        if (!p) return;
        paragraphDialog(d().paraFmt(p), (v) => {
          ed.setRange(r, false);
          this.para((pis) => d().applyParaProps(pis, {
            jc: v.jc, left: v.left || null, right: v.right || null,
            firstLine: v.special === 'first' ? v.by : v.special === 'none' ? null : undefined,
            hanging: v.special === 'hanging' ? v.by : v.special === 'none' ? null : undefined,
            before: v.before, after: v.after,
            line: v.lineMode === 'auto' ? Math.round(v.line * 240) : Math.round(v.line * 20), lineRule: v.lineMode,
          }));
        });
      },
      enabled: this.hasDoc,
    }, 'ctrl+shift+p');
    for (const [id, key] of [['Heading1', 'ctrl+alt+1'], ['Heading2', 'ctrl+alt+2'], ['Heading3', 'ctrl+alt+3']]) {
      c.register(`style:${id}`, { run: () => this.applyStyle(id), enabled: this.hasDoc }, key);
    }
    c.register('style:normal', { run: () => this.applyStyle(d().styles.defaultPara ?? 'Normal'), enabled: this.hasDoc }, 'ctrl+shift+n');
    // page
    c.register('page:setup', {
      run: () => pageSetupDialog(d().pageSetup(), (ps) => { ed.commit(); d().setPageSetup(ps); ed.render(); if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); else this.syncRulers(); }),
      enabled: this.hasDoc,
    });
    // table
    for (const op of ['row-above', 'row-below', 'col-left', 'col-right', 'del-row', 'del-col', 'del-table'] as const) {
      c.register(`table:${op}`, { run: () => this.structural((pos) => d().tableOp(pos.pi, op)), enabled: this.inTable });
    }
    for (const v of ['top', 'center', 'bottom'] as const) {
      c.register(`table:valign-${v}`, {
        run: () => {
          ed.commit();
          const r = ed.getRange();
          if (!r) return;
          d().setCellVAlign(d().paraRange(r), v);
          ed.render();
          ed.setRange(r, false);
          this.queueSync();
        },
        enabled: this.inTable,
        checked: () => { const r = ed.getRange(); return !!r && !!this.doc && d().cellVAlign(r.start.pi) === v; },
      });
    }
    // Word-style additions: views, insert, font dialog, case, review, layout, table alignment
    const setMode = (m: 'print' | 'web') => {
      ed.setMode(m);
      if (m === 'print' && this.fit) this.setZoom(this.fitZoom(this.fit), this.fit);
      this.syncRulers();
      this.syncStatus();
      this.syncRibbon();
      this.message(m === 'web' ? '웹 모양: 페이지 구분 없이 창 너비에 맞춰 보여 줍니다' : '인쇄 모양');
    };
    c.register('view:print-layout', { run: () => setMode('print'), enabled: this.hasDoc, checked: () => ed.mode === 'print' }, 'ctrl+alt+p');
    c.register('view:web-layout', { run: () => setMode('web'), enabled: this.hasDoc, checked: () => ed.mode === 'web' }, 'ctrl+alt+w');
    c.register('view:zoom-dialog', {
      run: () => zoomDialog(ed.zoom, ed.mode === 'web', (z) => {
        if (z.kind === 'percent') this.setZoom(z.value);
        else this.setZoom(this.fitZoom(z.kind), z.kind);
      }),
      enabled: this.hasDoc,
    });
    c.register('view:zoom-text', { run: () => this.setZoom(this.fitZoom('text'), 'text'), enabled: () => this.hasDoc() && ed.mode === 'print', checked: () => this.fit === 'text' });
    const switchWindows = async (at: HTMLElement) => {
      const list = await listWindows();
      menuPopover(at, windowMenuItems(list, (id) => focusWindow(id).catch((e) => toast('창 전환: ' + e)), () => this.cmd.run('window:new')));
    };
    c.register('window:switch', { run: () => switchWindows(anchor('tb-windows')) });
    c.register('window:new', { run: () => openNewWindow().catch((e) => toast('새 창: ' + e)) }, 'ctrl+alt+n');
    c.register('view:toggle-ribbon', { run: () => { this.showRibbon = !this.showRibbon; $('ribbon').classList.toggle('collapsed', !this.showRibbon); if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); }, checked: () => this.showRibbon }, 'ctrl+alt+r');
    c.register('view:gridlines', { run: () => { this.showGridlines = !this.showGridlines; ed.root.classList.toggle('no-gridlines', !this.showGridlines); }, checked: () => this.showGridlines });
    c.register('insert:line-break', { run: () => this.structural((pos) => d().insertLineBreak(pos)), enabled: this.hasDoc });
    c.register('insert:date', { run: () => { this.keepCaret(); dateDialog((t) => this.insertPlain(t, '날짜 및 시간')); }, enabled: this.hasDoc });
    c.register('insert:symbol', { run: () => { this.keepCaret(); symbolPopover(anchor('tb-symbol'), (ch) => this.insertPlain(ch, '기호')); }, enabled: this.hasDoc });
    c.register('insert:emoji', { run: () => { this.keepCaret(); emojiPopover(anchor('tb-emoji'), (ch) => this.insertPlain(ch, '이모지')); }, enabled: this.hasDoc }, 'ctrl+alt+e');
    // Headers and footers
    const curPage = () => Math.max(1, this.editor.caretPage());
    c.register('hf:edit-header', { run: () => { this.editor.enterHF('header', ed.hf?.page ?? curPage()); }, enabled: this.hasDoc });
    c.register('hf:edit-footer', { run: () => { this.editor.enterHF('footer', ed.hf?.page ?? curPage()); }, enabled: this.hasDoc });
    c.register('hf:goto-other', { run: () => { const h = ed.hf; if (h) ed.enterHF(h.kind === 'header' ? 'footer' : 'header', h.page); }, enabled: () => !!ed.hf });
    c.register('hf:close', { run: () => ed.exitHF(), enabled: () => !!ed.hf });
    for (const kind of ['header', 'footer'] as const) {
      c.register(`hf:remove-${kind}`, {
        run: () => {
          const page = ed.hf?.page ?? curPage();
          ed.exitHF();
          d().removeHeaderFooter(kind, d().hfType(page));
          ed.paginate();
          toast(kind === 'header' ? '머리글을 삭제했습니다' : '바닥글을 삭제했습니다');
        },
        enabled: () => !!this.doc && !!d().hfPart(kind, d().hfType(ed.hf?.page ?? curPage())),
      });
    }
    c.register('hf:first-different', { run: () => { ed.commitHF(); d().setHFOptions({ titlePg: !d().hfOptions().titlePg }); this.reopenHF(); }, enabled: this.hasDoc, checked: () => !!this.doc && d().hfOptions().titlePg });
    c.register('hf:odd-even', { run: () => { ed.commitHF(); d().setHFOptions({ evenOdd: !d().hfOptions().evenOdd }); this.reopenHF(); }, enabled: this.hasDoc, checked: () => !!this.doc && d().hfOptions().evenOdd });
    c.register('hf:insert-page', { run: () => ed.hfInsertField('PAGE'), enabled: () => !!ed.hf });
    c.register('hf:insert-pages', { run: () => ed.hfInsertField('NUMPAGES'), enabled: () => !!ed.hf });
    for (const kind of ['header', 'footer'] as const) {
      for (const jc of ['left', 'center', 'right'] as const) {
        c.register(`hf:page-number-${kind}-${jc}`, {
          run: () => { ed.exitHF(); d().insertPageNumber(kind, jc, this.pageNumberTotal); ed.paginate(); },
          enabled: this.hasDoc,
        });
      }
    }
    c.register('hf:page-number-menu', {
      run: () => menuPopover(anchor('tb-pagenum'), [
        { label: '페이지 위쪽 - 왼쪽', action: () => this.cmd.run('hf:page-number-header-left') },
        { label: '페이지 위쪽 - 가운데', action: () => this.cmd.run('hf:page-number-header-center') },
        { label: '페이지 위쪽 - 오른쪽', action: () => this.cmd.run('hf:page-number-header-right') },
        { sep: true },
        { label: '페이지 아래쪽 - 왼쪽', action: () => this.cmd.run('hf:page-number-footer-left') },
        { label: '페이지 아래쪽 - 가운데', action: () => this.cmd.run('hf:page-number-footer-center') },
        { label: '페이지 아래쪽 - 오른쪽', action: () => this.cmd.run('hf:page-number-footer-right') },
        { sep: true },
        { label: (this.pageNumberTotal ? '✓ ' : '') + '"n / 전체" 형식', action: () => { this.pageNumberTotal = !this.pageNumberTotal; } },
      ]),
      enabled: this.hasDoc,
    });
    for (const kind of ['header', 'footer'] as const) {
      c.register(`hf:${kind}-menu`, {
        run: () => menuPopover(anchor(kind === 'header' ? 'tb-header' : 'tb-footer'), [
          { label: kind === 'header' ? '머리글 편집' : '바닥글 편집', action: () => this.cmd.run(`hf:edit-${kind}`) },
          { label: kind === 'header' ? '머리글 삭제' : '바닥글 삭제', disabled: !this.cmd.enabled(`hf:remove-${kind}`), action: () => this.cmd.run(`hf:remove-${kind}`) },
        ]),
        enabled: this.hasDoc,
      });
    }
    c.register('insert:hyperlink', {
      run: () => {
        ed.commit();
        const r = ed.getRange();
        if (!r) return;
        const sel = window.getSelection()?.toString() ?? '';
        const bookmarks = Array.from(d().body.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'bookmarkStart')).map((b) => b.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'name') ?? '').filter((n) => n && n !== '_GoBack');
        hyperlinkDialog(sel.includes('\n') ? '' : sel, bookmarks, (url, text) => {
          const next = d().insertHyperlink(r, url, text);
          ed.render();
          ed.setRange({ start: next, end: next });
        });
      },
      enabled: this.hasDoc,
    }, 'ctrl+k');
    c.register('format:font', {
      run: () => {
        ed.commit();
        const r = ed.getRange();
        const f = this.caretFmt()?.fmt;
        if (!r || !f) return;
        fontDialog(f, FONTS, (v) => {
          ed.setRange(r, false);
          this.runProps({ ...v, sz: v.sz, u: v.u === 'none' ? false : v.u }, '글꼴');
        });
      },
      enabled: this.hasDoc,
    }, 'ctrl+d');
    for (const mode of ['upper', 'lower', 'title', 'sentence'] as const) {
      c.register(`format:case-${mode}`, {
        run: () => {
          ed.commit();
          const r = ed.getRange();
          if (!r || comparePos(r.start, r.end) === 0) { this.message('바꿀 글자를 먼저 선택하세요'); return; }
          d().changeCase(r, mode);
          const n = Math.abs(r.end.pi - r.start.pi) + 1;
          ed.refresh(Math.min(r.start.pi, r.end.pi), n, n);
          ed.setRange(r, false);
        },
        enabled: this.hasDoc,
      });
    }
    for (const accept of [true, false]) {
      c.register(`review:${accept ? 'accept' : 'reject'}-all`, {
        run: () => {
          ed.commit();
          const n = d().revisionCount();
          if (!n) { toast('변경 내용이 없습니다'); return; }
          d().resolveRevisions(accept);
          ed.render();
          toast(`변경 내용 ${n}개를 ${accept ? '적용' : '거부'}했습니다`);
        },
        enabled: this.hasDoc,
      });
    }
    const pageMenu = (at: string, items: [string, () => void][]) => menuPopover(anchor(at), items.map(([label, f]) => ({ label, action: () => { ed.commit(); f(); ed.render(); if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); this.syncRulers(); } })));
    const cm = TWIP_PER_CM;
    c.register('layout:margins', {
      run: () => pageMenu('tb-margins', [
        ['보통 (위아래 2.54, 좌우 2.54 cm)', () => d().setPageSetup({ ...d().pageSetup(), top: 1440, bottom: 1440, left: 1440, right: 1440 })],
        ['좁게 (1.27 cm)', () => d().setPageSetup({ ...d().pageSetup(), top: 720, bottom: 720, left: 720, right: 720 })],
        ['보통 (한국 기본: 위 3, 아래 2.5, 좌우 3 cm)', () => d().setPageSetup({ ...d().pageSetup(), top: 3 * cm, bottom: 2.5 * cm, left: 3 * cm, right: 3 * cm })],
        ['넓게 (위아래 2.54, 좌우 5.08 cm)', () => d().setPageSetup({ ...d().pageSetup(), top: 1440, bottom: 1440, left: 2880, right: 2880 })],
        ['사용자 지정 여백…', () => this.cmd.run('page:setup')],
      ]),
      enabled: this.hasDoc,
    });
    c.register('layout:orientation', {
      run: () => pageMenu('tb-orient', [
        ['세로', () => { const ps = d().pageSetup(); if (ps.orient === 'landscape') d().setPageSetup({ ...ps, orient: 'portrait', w: Math.min(ps.w, ps.h), h: Math.max(ps.w, ps.h) }); }],
        ['가로', () => { const ps = d().pageSetup(); if (ps.orient !== 'landscape') d().setPageSetup({ ...ps, orient: 'landscape', w: Math.max(ps.w, ps.h), h: Math.min(ps.w, ps.h) }); }],
      ]),
      enabled: this.hasDoc,
    });
    c.register('layout:size', {
      run: () => pageMenu('tb-size', ([['A4', 210, 297], ['A3', 297, 420], ['A5', 148, 210], ['B5 (JIS)', 182, 257], ['Letter', 215.9, 279.4], ['Legal', 215.9, 355.6]] as [string, number, number][]).map(([n, w, h]) => [
        `${n} (${w} × ${h} mm)`, () => { const ps = d().pageSetup(); const [a, b] = [w * cm / 10, h * cm / 10]; d().setPageSetup({ ...ps, w: ps.orient === 'landscape' ? b : a, h: ps.orient === 'landscape' ? a : b }); },
      ] as [string, () => void]).concat([['기타 용지 크기…', () => this.cmd.run('page:setup')]])),
      enabled: this.hasDoc,
    });
    c.register('table:delete-menu', {
      run: () => menuPopover(anchor('tb-tdelete'), [
        { label: '열 삭제', action: () => this.cmd.run('table:del-col') },
        { label: '행 삭제', action: () => this.cmd.run('table:del-row') },
        { label: '표 삭제', action: () => this.cmd.run('table:del-table') },
      ]),
      enabled: this.inTable,
    });
    for (const v of ['top', 'center', 'bottom'] as const) {
      for (const h of ['left', 'center', 'right'] as const) {
        c.register(`table:align-${v}-${h}`, {
          run: () => {
            ed.commit();
            const r = ed.getRange();
            if (!r) return;
            d().alignCells(d().paraRange(r), v, h);
            ed.render();
            ed.setRange(r, false);
            this.queueSync();
          },
          enabled: this.inTable,
        });
      }
    }
    c.register('style:Title', { run: () => this.applyStyle('Title'), enabled: this.hasDoc });
    // tools
    c.register('tool:stats', { run: () => this.statsDialog(), enabled: this.hasDoc });
    c.register('tool:shortcuts', () => {
      const labels = new Map<string, string>();
      document.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((m) => labels.set(m.dataset.cmd!, m.querySelector('.md-label')!.textContent!));
      const rows = c.shortcutList().filter(([, id]) => labels.has(id)).map(([k, id]) => `<tr><td>${labels.get(id)}</td><td><kbd>${k.replace(/\+/g, '</kbd> + <kbd>')}</kbd></td></tr>`).join('');
      dialog('단축키 목록', html(`<table class="kbd-table">${rows}<tr><td>문단 나누기 / 줄 바꿈</td><td><kbd>Enter</kbd> / <kbd>shift</kbd> + <kbd>Enter</kbd></td></tr><tr><td>탭 · 목록 수준 · 다음 셀</td><td><kbd>Tab</kbd> (<kbd>shift</kbd>: 반대로)</td></tr><tr><td>하이퍼링크 열기</td><td><kbd>ctrl</kbd> + 클릭</td></tr></table>`), [{ label: '닫기', primary: true }]);
    }, 'f1');
  }

  pageNumberTotal = false;

  /** View ▸ 창 전환 submenu: one entry per open window. */
  private async fillWindowList(host: HTMLElement, close: () => void) {
    const list = await listWindows();
    host.innerHTML = '';
    for (const it of windowMenuItems(list, (id) => focusWindow(id).catch((e) => toast('창 전환: ' + e)), () => this.cmd.run('window:new'))) {
      if (it.sep) { const d = document.createElement('div'); d.className = 'md-sep'; host.appendChild(d); continue; }
      const row = document.createElement('div');
      row.className = 'md-item';
      row.innerHTML = '<span class="md-label"></span><span class="md-shortcut"></span>';
      row.querySelector('.md-label')!.textContent = it.label ?? '';
      row.querySelector('.md-shortcut')!.textContent = /Mac/.test(navigator.platform) ? (it.shortcut ?? '').replace('Ctrl+', '⌘').replace('Alt+', '⌥') : it.shortcut ?? '';
      row.addEventListener('mousedown', (e) => e.preventDefault());
      row.addEventListener('click', () => { close(); it.action?.(); });
      host.appendChild(row);
    }
  }

  /** Keep the caret where it is while a dialog or popover has focus. */
  private keepCaret() { if (!this.editor.hf) this.editor.rememberSelection(); }

  /** Insert plain text at the caret: in the header/footer being edited, or in the body. */
  private insertPlain(text: string, label: string) {
    const hf = this.editor.hf;
    if (hf) { hf.el.focus(); document.execCommand('insertText', false, text); return; }
    this.structural((pos) => this.doc!.insertText(pos, text, label));
  }

  /** After the page-type options change, edit the header/footer that now applies to the page. */
  private reopenHF() {
    const hf = this.editor.hf;
    if (hf) this.editor.enterHF(hf.kind, hf.page); else this.editor.paginate();
    this.syncRibbon();
  }

  /** Show the "머리글 및 바닥글" tab while a header/footer is edited. */
  private onHFChanged() {
    const on = !!this.editor.hf;
    const tab = document.querySelector<HTMLElement>('.rb-tab[data-tab="hf"]')!;
    tab.hidden = !on;
    document.getElementById('studio-root')!.classList.toggle('hf-active', on);
    if (on) this.showTab('hf');
    else if (tab.classList.contains('active')) this.showTab('home');
    this.message(on ? `${this.editor.hf!.kind === 'header' ? '머리글' : '바닥글'} 편집 중 — Esc 또는 본문을 두 번 눌러 닫기` : '');
    this.syncRibbon();
    this.queueSync();
  }

  private applyStyle(id: string) {
    this.para((pis) => this.doc!.setStyle(pis, id));
    this.buildStyleList(id);
    this.updateOutline();
  }

  private statsDialog() {
    this.editor.commit();
    const s = this.doc!.stats();
    dialog('글자 수', html(`<table class="kbd-table">
      <tr><td>쪽</td><td>${this.editor.pages}</td></tr>
      <tr><td>단어</td><td>${s.words.toLocaleString()}</td></tr>
      <tr><td>글자 (공백 포함)</td><td>${s.chars.toLocaleString()}</td></tr>
      <tr><td>글자 (공백 제외)</td><td>${s.charsNoSpace.toLocaleString()}</td></tr>
      <tr><td>문단</td><td>${s.paragraphs.toLocaleString()}</td></tr></table>`), [{ label: '확인', primary: true }]);
  }

  private infoDialog() {
    const doc = this.doc!;
    const ps = doc.pageSetup();
    const core = doc.coreProps();
    const esc = (s: string) => s.replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]!));
    dialog('문서 정보', html(`<table class="kbd-table">
      <tr><td>파일</td><td>${esc(this.fileName)}</td></tr>
      <tr><td>쪽</td><td>${this.editor.pages}쪽</td></tr>
      <tr><td>용지</td><td>${(ps.w / 566.929).toFixed(1)} × ${(ps.h / 566.929).toFixed(1)} cm</td></tr>
      <tr><td>문단</td><td>${doc.paragraphs().length}개</td></tr>
      <tr><td>제목</td><td>${esc(core.title)}</td></tr>
      <tr><td>만든 이</td><td>${esc(core.creator)}</td></tr>
      <tr><td>수정</td><td>${esc(core.modified)}</td></tr></table>`), [{ label: '확인', primary: true }]);
  }

  private confirmDiscard() {
    return !this.doc?.dirty || confirm('저장하지 않은 변경 내용이 있습니다. 계속할까요?');
  }

  // ------------------------------------------------------------------ DOM bindings

  private bindUi() {
    installIcons();
    $('font-list').innerHTML = FONTS.map((f) => `<option value="${f}">`).join('');
    $('size-list').innerHTML = SIZES.map((s) => `<option value="${s}">`).join('');

    // Menus: click to open, hover to switch (like HOP's menu bar).
    const menus = document.querySelectorAll<HTMLElement>('.menu-item');
    const closeMenus = () => menus.forEach((m) => m.classList.remove('open'));
    const refreshMenu = (m: HTMLElement) => {
      m.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((it) => {
        it.classList.toggle('disabled', !this.cmd.enabled(it.dataset.cmd!));
        it.classList.toggle('checked', this.cmd.checked(it.dataset.cmd!));
      });
      const wl = m.querySelector<HTMLElement>('#window-list');
      if (wl) this.fillWindowList(wl, closeMenus);
    };
    menus.forEach((m) => {
      m.querySelector('.menu-title')!.addEventListener('mousedown', (e) => {
        e.preventDefault();
        this.editor.rememberSelection();
        const was = m.classList.contains('open');
        closeMenus();
        if (!was) { refreshMenu(m); m.classList.add('open'); }
      });
      m.addEventListener('mouseenter', () => {
        if ([...menus].some((x) => x.classList.contains('open')) && !m.classList.contains('open')) { closeMenus(); refreshMenu(m); m.classList.add('open'); }
      });
    });
    document.addEventListener('mousedown', (e) => { if (!(e.target as Element).closest?.('.menu-item')) closeMenus(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenus(); });

    // Any [data-cmd] element runs its command. mousedown is suppressed to keep the text caret.
    document.querySelectorAll<HTMLElement>('[data-cmd]').forEach((el) => {
      el.addEventListener('mousedown', (e) => e.preventDefault());
      el.addEventListener('click', () => { closeMenus(); this.cmd.run(el.dataset.cmd!); });
    });

    // Style bar inputs keep the document selection while they have focus.
    const fn = $<HTMLInputElement>('font-name'), fs = $<HTMLInputElement>('font-size');
    const lineSel = $<HTMLSelectElement>('line-spacing');
    for (const inp of [fn, fs, lineSel]) inp.addEventListener('mousedown', () => this.editor.rememberSelection());
    for (const inp of [fn, fs]) inp.addEventListener('focus', () => inp.select());
    const restore = () => { const r = this.editor.getRange(); this.editor.setRange(r, false); };
    fn.addEventListener('change', () => { const v = fn.value.trim(); if (v) { restore(); this.runProps({ font: v }, '글꼴'); } });
    fn.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); fn.blur(); } if (e.key === 'Escape') { restore(); } });
    const setSize = (v: number) => { if (v > 0 && v <= 1638) { restore(); this.runProps({ sz: Math.round(v * 2) / 2 }, '글자 크기'); } };
    fs.addEventListener('change', () => setSize(parseFloat(fs.value)));
    fs.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); fs.blur(); } if (e.key === 'Escape') restore(); });
    for (const [id, dir] of [['btn-size-up', 1], ['btn-size-down', -1]] as const) {
      $(id).addEventListener('mousedown', (e) => e.preventDefault());
      $(id).addEventListener('click', () => this.stepSize(dir));
    }
    lineSel.addEventListener('change', () => { if (lineSel.value) { restore(); this.cmd.run(`format:line-${Number(lineSel.value)}`); } });

    const splitButton = (id: string, apply: () => void, open: () => void) => {
      const b = $(id);
      b.addEventListener('mousedown', (e) => { e.preventDefault(); this.editor.rememberSelection(); });
      b.addEventListener('click', (e) => {
        const r = b.getBoundingClientRect();
        // Left part applies the last value, the arrow opens the palette.
        if (e.clientX - r.left < r.width * 0.6 && e.detail) apply(); else open();
      });
      b.addEventListener('contextmenu', (e) => { e.preventDefault(); open(); });
    };
    splitButton('btn-text-color', () => this.runProps({ color: this.lastTextColor }, '글꼴 색'), () => this.cmd.run('format:text-color'));
    splitButton('btn-highlight', () => this.runProps({ highlight: this.lastHighlight }, '텍스트 강조 색'), () => this.cmd.run('format:highlight'));

    // Ribbon: tabs, collapse on double-click (as in Word), checkboxes, case menu, layout fields.
    document.querySelectorAll<HTMLElement>('.rb-tab').forEach((t) => {
      t.addEventListener('mousedown', (e) => e.preventDefault());
      t.addEventListener('click', () => this.showTab(t.dataset.tab!));
      t.addEventListener('dblclick', () => this.cmd.run('view:toggle-ribbon'));
    });
    document.querySelectorAll<HTMLInputElement>('[data-check]').forEach((cb) => {
      cb.addEventListener('mousedown', (e) => e.preventDefault());
      cb.addEventListener('change', () => { this.cmd.run(cb.dataset.check!); this.syncRibbon(); });
    });
    $('btn-case').addEventListener('mousedown', (e) => { e.preventDefault(); this.editor.rememberSelection(); });
    $('btn-case').addEventListener('click', () => menuPopover($('btn-case'), [
      { label: '문장의 첫 글자를 대문자로', action: () => this.cmd.run('format:case-sentence') },
      { label: '모두 소문자로', action: () => this.cmd.run('format:case-lower') },
      { label: '모두 대문자로', action: () => this.cmd.run('format:case-upper') },
      { label: '각 단어의 첫 글자를 대문자로', action: () => this.cmd.run('format:case-title') },
    ]));
    for (const [id, key, k] of [['lay-left', 'left', TWIP_PER_CM], ['lay-right', 'right', TWIP_PER_CM], ['lay-before', 'before', 20], ['lay-after', 'after', 20]] as const) {
      const inp = $<HTMLInputElement>(id);
      inp.addEventListener('mousedown', () => this.editor.rememberSelection());
      inp.addEventListener('change', () => {
        const v = Math.max(0, Number(inp.value) || 0) * k;
        restore();
        this.para((pis) => this.doc!.applyParaProps(pis, { [key]: Math.round(v) }, '단락'));
      });
      inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); inp.blur(); } });
    }
    const zs = $<HTMLInputElement>('zoom-slider');
    zs.addEventListener('input', () => this.setZoom(Number(zs.value) / 100));
    // Shortcut labels in Mac notation (⌘ ⇧ ⌥) on macOS, like Word.app.
    if (/Mac|iPhone|iPad/.test(navigator.platform)) {
      document.querySelectorAll<HTMLElement>('.md-shortcut').forEach((e) => {
        e.textContent = (e.textContent ?? '').replace(/Ctrl\+/g, '⌘').replace(/Shift\+/g, '⇧').replace(/Alt\+/g, '⌥').replace(/Enter/, '↩').replace(/Delete/, '⌫');
      });
      document.querySelectorAll<HTMLElement>('[title]').forEach((e) => { e.title = e.title.replace(/Ctrl\+/g, '⌘').replace(/Shift\+/g, '⇧').replace(/Alt\+/g, '⌥'); });
    }

    // Keyboard.
    window.addEventListener('keydown', (e) => this.onKey(e));
    document.addEventListener('paste', (e) => this.onPaste(e));
    document.addEventListener('copy', (e) => { if (this.editor.isEditorTarget(e.target)) this.copyToClipboard(); });
    document.addEventListener('cut', (e) => { if (this.editor.isEditorTarget(e.target)) this.copyToClipboard(); });

    // Table context menu.
    this.editor.root.addEventListener('contextmenu', (e) => {
      if (!this.doc) return;
      const inTable = (e.target as HTMLElement).closest('td');
      e.preventDefault();
      const run = (id: string) => () => this.cmd.run(id);
      const mac = /Mac/.test(navigator.platform);
      const k = (s: string) => mac ? s.replace(/Ctrl\+/g, '⌘').replace(/Shift\+/g, '⇧').replace(/Alt\+/g, '⌥') : s;
      const items = [
        { label: '잘라내기', shortcut: k('Ctrl+X'), disabled: this.editor.collapsed(), action: run('edit:cut') },
        { label: '복사', shortcut: k('Ctrl+C'), disabled: this.editor.collapsed(), action: run('edit:copy') },
        { label: '붙여넣기', shortcut: k('Ctrl+V'), action: run('edit:paste') },
        { sep: true },
        { label: '글꼴…', shortcut: k('Ctrl+D'), action: run('format:font') },
        { label: '단락…', shortcut: k('Ctrl+Alt+M'), action: run('format:paragraph') },
        { label: '글머리 기호', action: run('format:bullet') },
        { label: '번호 매기기', action: run('format:numbering') },
        { sep: true },
        { label: '하이퍼링크…', shortcut: k('Ctrl+K'), action: run('insert:hyperlink') },
      ];
      if (inTable) items.push(
        { sep: true },
        { label: '위에 행 삽입', action: run('table:row-above') },
        { label: '아래에 행 삽입', action: run('table:row-below') },
        { label: '왼쪽에 열 삽입', action: run('table:col-left') },
        { label: '오른쪽에 열 삽입', action: run('table:col-right') },
        { sep: true },
        { label: '행 삭제', action: run('table:del-row') },
        { label: '열 삭제', action: run('table:del-col') },
        { label: '표 삭제', action: run('table:del-table') },
        { sep: true },
        { label: '셀 맞춤: 위쪽', action: run('table:valign-top') },
        { label: '셀 맞춤: 가운데', action: run('table:valign-center') },
        { label: '셀 맞춤: 아래쪽', action: run('table:valign-bottom') },
      );
      contextMenu(e.clientX, e.clientY, items as Parameters<typeof contextMenu>[2]);
    });

    // Drag & drop of .docx files / images.
    const sc = $('scroll-container');
    sc.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('Files')) { e.preventDefault(); sc.classList.add('dragover'); } });
    sc.addEventListener('dragleave', () => sc.classList.remove('dragover'));
    sc.addEventListener('drop', async (e) => {
      if (!e.dataTransfer?.files.length) return;
      e.preventDefault();
      sc.classList.remove('dragover');
      const f = e.dataTransfer.files[0];
      if (/\.docx$/i.test(f.name)) { if (this.confirmDiscard()) await this.open(new Uint8Array(await f.arrayBuffer()), f.name); }
      else if (f.type.startsWith('image/') && this.doc) await this.insertImageFile(f, f.name);
    });
    sc.addEventListener('scroll', () => this.syncRulers());
    sc.addEventListener('wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey) || !this.doc) return;
      e.preventDefault();
      this.setZoom(this.editor.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    // Clicking the grey area below the last page puts the caret at the end.
    sc.addEventListener('mousedown', (e) => {
      if (e.target !== sc && e.target !== $('scroll-content')) return;
      if (!this.doc) return;
      e.preventDefault();
      const last = this.doc.paragraphs().length - 1;
      const pos = { pi: last, off: this.doc.lengthOf(this.doc.para(last)!) };
      this.editor.setRange({ start: pos, end: pos });
    });
    new ResizeObserver(() => {
      if (this.editor.mode === 'web') { this.editor.paginate(); return; } // text re-wraps to the window
      if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); else this.syncRulers();
    }).observe(sc);
    window.addEventListener('beforeunload', (e) => { this.editor.commit(); if (this.doc?.dirty) e.preventDefault(); });
  }

  private onKey(e: KeyboardEvent) {
    if (document.querySelector('.dlg-backdrop')) return;
    const t = e.target as HTMLElement;
    if (e.key === 'Escape') { closePopover(); closeContextMenu(); this.editor.deselectImage(); }
    if (t.matches('input, textarea, select')) return;
    if (e.isComposing) return;
    const id = this.cmd.lookup(e);
    if (!id) return;
    // Clipboard and select-all use the browser's own handling (paste arrives as a paste event).
    if (['edit:cut', 'edit:copy', 'edit:paste', 'edit:select-all'].includes(id) && this.editor.isEditorTarget(t)) return;
    e.preventDefault();
    this.cmd.run(id);
  }
}
