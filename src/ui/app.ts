// Application controller: document lifecycle, command registration, UI state sync.
import { getCurrentWindow } from '@tauri-apps/api/window';
import { DocxDocument, Pos, Range, comparePos, normRange } from '../engine/document';
import { RunProps } from '../engine/model';
import { ImageCache } from '../engine/render';
import { RunFmt, styleLabel } from '../engine/styles';
import { Commands } from './commands';
import { closeContextMenu, closePopover, colorPopover, contextMenu, dialog, highlightPopover, html, pageSetupDialog, paragraphDialog, tablePopover, toast } from './dialogs';
import { DocEditor } from './editor';
import { installIcons } from './icons';
import { printDocument } from './print';
import { drawRulers } from './ruler';
import { isDesktop, pickImage, pickOpenFile, saveFile } from './shell';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

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
  fit: 'width' | 'page' | null = null;
  showRuler = true;
  showOutline = true;
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
      layoutChanged: () => { this.syncStatus(); this.syncRulers(); },
      status: (m) => this.message(m),
    }, $('page-wrap'));
    this.cmd.onError = (e) => { console.error(e); toast('오류: ' + ((e as Error).message ?? e)); };
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

  private fitZoom(kind: 'width' | 'page') {
    if (!this.doc) return 1;
    const sc = $('scroll-container');
    const g = this.editor.geometry();
    const zw = (sc.clientWidth - 48) / (g.W * PX);
    const zh = (sc.clientHeight - 32) / (g.H * PX);
    return Math.max(0.2, kind === 'width' ? zw : Math.min(zw, zh));
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
    requestAnimationFrame(() => { this.syncQueued = false; this.syncStyleBar(); this.syncRulers(); this.syncStatus(); });
  }

  syncStatus() {
    const n = this.editor.pages;
    $('sb-page').textContent = this.doc ? `쪽 ${Math.min(n, this.editor.caretPage())} / ${n}` : '쪽 0 / 0';
    $('sb-zoom-val').textContent = Math.round(this.editor.zoom * 100) + '%';
  }

  private updateInfo() {
    if (!this.doc) return;
    const s = this.doc.stats();
    $('sb-count').textContent = `글자 ${s.chars.toLocaleString()} · 단어 ${s.words.toLocaleString()}`;
  }

  /** Character formatting at the caret: the pending span, the run under the caret, or the model. */
  private caretFmt(): { fmt: RunFmt; pi: number } | null {
    const doc = this.doc;
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
    const sel = $<HTMLSelectElement>('para-style');
    if (p && document.activeElement !== sel) {
      const id = doc!.pStyleOf(p) ?? doc!.styles.defaultPara ?? '';
      if (![...sel.options].some((o) => o.value === id)) this.buildStyleList(id);
      sel.value = id;
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
    }
    $('sb-mode').textContent = cell ? '표' : this.editor.collapsed() ? '입력' : '선택';
    this.markOutline(st?.pi ?? -1);
  }

  private buildStyleList(extra?: string) {
    const doc = this.doc;
    const sel = $<HTMLSelectElement>('para-style');
    if (!doc) { sel.innerHTML = ''; return; }
    const used = new Set(doc.paragraphs().map((p) => doc.pStyleOf(p)).filter((s): s is string => !!s));
    if (extra) used.add(extra);
    if (doc.styles.defaultPara) used.add(doc.styles.defaultPara);
    const list = doc.styles.paragraphStyles(used);
    // Offer the headings even when the document has no definition yet (added on first use).
    const ids = new Set(list.map((s) => s.id));
    const extraBuiltins = ['Heading1', 'Heading2', 'Heading3', 'Title'].filter((i) => !ids.has(i) && ![...doc.styles.styles.values()].some((s) => s.name.toLowerCase() === (i === 'Title' ? 'title' : 'heading ' + i.slice(-1))));
    sel.innerHTML = list.map((s) => `<option value="${s.id}">${styleLabel(s)}</option>`).join('') +
      extraBuiltins.map((i) => `<option value="${i}">${i === 'Title' ? '제목' : '제목 ' + i.slice(-1)}</option>`).join('');
  }

  syncRulers() {
    if (!this.doc || !this.showRuler) return;
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

  private findDialog() {
    this.editor.commit();
    const body = html(`<div class="form" style="grid-template-columns:auto 1fr">
      <label>찾을 내용</label><input type="text" name="f">
      <label>바꿀 내용</label><input type="text" name="r">
      <label></label><label style="color:inherit"><input type="checkbox" name="mc"> 대/소문자 구분</label>
      <label></label><span id="find-result" style="font-size:12px;color:#5c6370"></span></div>`);
    const q = (n: string) => body.querySelector(`[name="${n}"]`) as HTMLInputElement;
    const sel = window.getSelection()?.toString();
    if (sel && !sel.includes('\n')) q('f').value = sel;
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
    dialog('찾기 / 바꾸기', body, [
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
    c.register('file:about', () => { dialog('제품 정보', html(`<div style="line-height:1.7"><b>Oword</b> 0.1.0<br>DOCX 문서 편집기<br><span style="color:#5c6370;font-size:12px">Opptx(PPTX 편집기)와 같은 구조 — 문서 엔진 · 렌더러 · 명령 기반 UI · 데스크톱 셸 — 로 만든 Word 문서 편집기입니다. 이 구조는 HOP(HWP 편집기, MIT)을 참고했습니다.<br>MIT License · Copyright (c) 2026 moonlightpark<br>서드파티 라이선스: THIRD_PARTY_NOTICES.md<br>Microsoft와 관계없음. Microsoft Word는 Microsoft Corporation의 상표입니다.</span></div>`), [{ label: '확인', primary: true }]); });
    // edit
    c.register('edit:undo', {
      run: () => { ed.commit(); const sel = d().undo() as Range | null; ed.render(); if (sel) ed.setRange(sel); this.buildStyleList(); this.queueSync(); },
      enabled: () => !!this.doc && (d().canUndo() || true),
    }, 'ctrl+z');
    c.register('edit:redo', {
      run: () => { ed.commit(); const sel = d().redo() as Range | null; ed.render(); if (sel) ed.setRange(sel); this.queueSync(); },
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
    c.register('edit:find', { run: () => this.findDialog(), enabled: this.hasDoc }, 'ctrl+f', 'ctrl+h');
    // view
    c.register('view:zoom-in', { run: () => this.stepZoom(1), enabled: this.hasDoc });
    c.register('view:zoom-out', { run: () => this.stepZoom(-1), enabled: this.hasDoc });
    c.register('view:zoom-width', { run: () => this.setZoom(this.fitZoom('width'), 'width'), enabled: this.hasDoc, checked: () => this.fit === 'width' });
    c.register('view:zoom-page', { run: () => this.setZoom(this.fitZoom('page'), 'page'), enabled: this.hasDoc, checked: () => this.fit === 'page' });
    c.register('view:zoom-100', { run: () => this.setZoom(1), enabled: this.hasDoc });
    c.register('view:toggle-ruler', { run: () => { this.showRuler = !this.showRuler; $('editor-area').classList.toggle('no-ruler', !this.showRuler); this.syncRulers(); }, checked: () => this.showRuler });
    c.register('view:toggle-outline', { run: () => { this.showOutline = !this.showOutline; $('outline-panel').classList.toggle('hidden', !this.showOutline); if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); }, checked: () => this.showOutline });
    c.register('view:toggle-marks', { run: () => { ed.showMarks = !ed.showMarks; ed.paginate(); $('tb-marks').classList.toggle('active', ed.showMarks); }, checked: () => ed.showMarks }, 'ctrl+shift+8');
    // insert
    c.register('insert:table', { run: () => tablePopover($('tb-table'), (r, col) => this.structural((pos) => d().insertTable(pos, r, col))), enabled: this.hasDoc });
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
    c.register('format:text-color', { run: () => colorPopover($('btn-text-color'), { allowNone: '자동' }, (hex) => { if (hex) { this.lastTextColor = hex; $('color-bar').style.background = '#' + hex; } this.runProps({ color: hex }, '글자 색'); }), enabled: this.hasDoc });
    c.register('format:highlight', { run: () => highlightPopover($('btn-highlight'), (name) => { if (name) this.lastHighlight = name; this.runProps({ highlight: name }, '형광펜'); }), enabled: this.hasDoc });
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
      c.register(`format:align-${id}`, { run: () => this.para((pis) => d().applyParaProps(pis, { jc }, '정렬')), enabled: this.hasDoc }, key);
    }
    c.register('format:bullet', { run: () => this.para((pis) => d().toggleList(pis, 'bullet')), enabled: this.hasDoc }, 'ctrl+shift+l');
    c.register('format:numbering', { run: () => this.para((pis) => d().toggleList(pis, 'number')), enabled: this.hasDoc });
    c.register('format:indent', { run: () => this.para((pis) => d().indent(pis, 1)), enabled: this.hasDoc }, 'ctrl+m');
    c.register('format:outdent', { run: () => this.para((pis) => d().indent(pis, -1)), enabled: this.hasDoc }, 'ctrl+shift+m');
    for (const [v, key] of [['1', 'ctrl+1'], ['1.15', ''], ['1.5', 'ctrl+5'], ['2', 'ctrl+2'], ['2.5', ''], ['3', '']]) {
      c.register(`format:line-${v}`, { run: () => this.para((pis) => d().applyParaProps(pis, { line: Math.round(Number(v) * 240), lineRule: 'auto' }, '줄 간격')), enabled: this.hasDoc }, ...(key ? [key] : []));
    }
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
    // tools
    c.register('tool:stats', { run: () => this.statsDialog(), enabled: this.hasDoc });
    c.register('tool:shortcuts', () => {
      const labels = new Map<string, string>();
      document.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((m) => labels.set(m.dataset.cmd!, m.querySelector('.md-label')!.textContent!));
      const rows = c.shortcutList().filter(([, id]) => labels.has(id)).map(([k, id]) => `<tr><td>${labels.get(id)}</td><td><kbd>${k.replace(/\+/g, '</kbd> + <kbd>')}</kbd></td></tr>`).join('');
      dialog('단축키 목록', html(`<table class="kbd-table">${rows}<tr><td>문단 나누기 / 줄 바꿈</td><td><kbd>Enter</kbd> / <kbd>shift</kbd> + <kbd>Enter</kbd></td></tr><tr><td>탭 · 목록 수준 · 다음 셀</td><td><kbd>Tab</kbd> (<kbd>shift</kbd>: 반대로)</td></tr><tr><td>하이퍼링크 열기</td><td><kbd>ctrl</kbd> + 클릭</td></tr></table>`), [{ label: '닫기', primary: true }]);
    }, 'f1');
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
    const refreshMenu = (m: HTMLElement) => m.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((it) => {
      it.classList.toggle('disabled', !this.cmd.enabled(it.dataset.cmd!));
      it.classList.toggle('checked', this.cmd.checked(it.dataset.cmd!));
    });
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
    document.addEventListener('mousedown', (e) => { if (!(e.target as HTMLElement).closest('.menu-item')) closeMenus(); });

    // Any [data-cmd] element runs its command. mousedown is suppressed to keep the text caret.
    document.querySelectorAll<HTMLElement>('[data-cmd]').forEach((el) => {
      el.addEventListener('mousedown', (e) => e.preventDefault());
      el.addEventListener('click', () => { closeMenus(); this.cmd.run(el.dataset.cmd!); });
    });

    // Style bar inputs keep the document selection while they have focus.
    const fn = $<HTMLInputElement>('font-name'), fs = $<HTMLInputElement>('font-size');
    const styleSel = $<HTMLSelectElement>('para-style'), lineSel = $<HTMLSelectElement>('line-spacing');
    for (const inp of [fn, fs, styleSel, lineSel]) inp.addEventListener('mousedown', () => this.editor.rememberSelection());
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
    styleSel.addEventListener('change', () => { restore(); this.applyStyle(styleSel.value); });
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
    splitButton('btn-text-color', () => this.runProps({ color: this.lastTextColor }, '글자 색'), () => this.cmd.run('format:text-color'));
    splitButton('btn-highlight', () => this.runProps({ highlight: this.lastHighlight }, '형광펜'), () => this.cmd.run('format:highlight'));

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
      const items = [
        { label: '오려두기', shortcut: 'Ctrl+X', disabled: this.editor.collapsed(), action: run('edit:cut') },
        { label: '복사하기', shortcut: 'Ctrl+C', disabled: this.editor.collapsed(), action: run('edit:copy') },
        { label: '붙이기', shortcut: 'Ctrl+V', action: run('edit:paste') },
        { sep: true },
        { label: '문단 모양…', shortcut: 'Ctrl+Shift+P', action: run('format:paragraph') },
        { label: '글머리표', action: run('format:bullet') },
        { label: '문단 번호', action: run('format:numbering') },
      ];
      if (inTable) items.push(
        { sep: true },
        { label: '위에 줄 넣기', action: run('table:row-above') },
        { label: '아래에 줄 넣기', action: run('table:row-below') },
        { label: '왼쪽에 칸 넣기', action: run('table:col-left') },
        { label: '오른쪽에 칸 넣기', action: run('table:col-right') },
        { sep: true },
        { label: '줄 지우기', action: run('table:del-row') },
        { label: '칸 지우기', action: run('table:del-col') },
        { label: '표 지우기', action: run('table:del-table') },
        { sep: true },
        { label: '셀 위쪽 정렬', action: run('table:valign-top') },
        { label: '셀 가운데 정렬', action: run('table:valign-center') },
        { label: '셀 아래쪽 정렬', action: run('table:valign-bottom') },
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
    new ResizeObserver(() => { if (this.fit) this.setZoom(this.fitZoom(this.fit), this.fit); else this.syncRulers(); }).observe(sc);
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
