// Popovers (color palette, highlight, table grid) and modal dialogs.
import { PageSetup, TWIP_PER_CM } from '../engine/document';
import { ParaFmt } from '../engine/styles';

let openPop: HTMLElement | null = null;

export function closePopover() {
  openPop?.remove();
  openPop = null;
}

function popover(anchor: HTMLElement, content: HTMLElement) {
  closePopover();
  const pop = document.createElement('div');
  pop.className = 'popover';
  pop.appendChild(content);
  pop.addEventListener('mousedown', (e) => { if (!(e.target as HTMLElement).matches('input')) e.preventDefault(); });
  document.body.appendChild(pop);
  const r = anchor.getBoundingClientRect();
  const pw = pop.offsetWidth;
  pop.style.left = Math.min(r.left, window.innerWidth - pw - 8) + 'px';
  pop.style.top = r.bottom + 4 + 'px';
  openPop = pop;
  setTimeout(() => {
    const off = (e: MouseEvent) => {
      if (!pop.contains(e.target as Node)) { closePopover(); document.removeEventListener('mousedown', off, true); }
    };
    document.addEventListener('mousedown', off, true);
  });
  return pop;
}

const PALETTE = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
  'F2F2F2', '7F7F7F', 'D0CECE', 'D6DCE4', 'D9E2F3', 'FBE5D5', 'EDEDED', 'FFF2CC', 'DEEBF6', 'E2EFD9',
  'D8D8D8', '595959', 'AEABAB', 'ADB9CA', 'B4C6E7', 'F7CBAC', 'DBDBDB', 'FEE599', 'BDD7EE', 'C5E0B3',
  'BFBFBF', '3F3F3F', '757070', '8496B0', '8EAADB', 'F4B183', 'C9C9C9', 'FFD965', '9CC3E5', 'A8D08D',
  'A5A5A5', '262626', '3A3838', '323F4F', '2F5496', 'C55A11', '7B7B7B', 'BF9000', '2E75B5', '538135',
  'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0',
];

export function colorPopover(anchor: HTMLElement, opts: { allowNone?: string }, pick: (hex: string | null) => void) {
  const wrap = document.createElement('div');
  const grid = document.createElement('div');
  grid.className = 'palette';
  for (const c of PALETTE) {
    const b = document.createElement('button');
    b.className = 'swatch';
    b.style.background = '#' + c;
    b.title = '#' + c;
    b.onclick = () => { closePopover(); pick(c); };
    grid.appendChild(b);
  }
  wrap.appendChild(grid);
  const row = document.createElement('div');
  row.className = 'pop-row';
  if (opts.allowNone) {
    const none = document.createElement('button');
    none.textContent = opts.allowNone;
    none.onclick = () => { closePopover(); pick(null); };
    row.appendChild(none);
  }
  const custom = document.createElement('input');
  custom.type = 'color';
  custom.title = '다른 색';
  custom.onchange = () => { closePopover(); pick(custom.value.slice(1).toUpperCase()); };
  const lbl = document.createElement('span');
  lbl.textContent = '다른 색:';
  row.append(lbl, custom);
  wrap.appendChild(row);
  popover(anchor, wrap);
}

export function tablePopover(anchor: HTMLElement, pick: (rows: number, cols: number) => void) {
  const wrap = document.createElement('div');
  const grid = document.createElement('div');
  grid.className = 'table-grid';
  const label = document.createElement('div');
  label.className = 'table-label';
  label.textContent = '표 크기 선택';
  const cells: HTMLElement[] = [];
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const d = document.createElement('div');
    d.onmouseenter = () => {
      cells.forEach((x, i) => x.classList.toggle('on', Math.floor(i / 10) <= r && i % 10 <= c));
      label.textContent = `${r + 1} × ${c + 1} 표`;
    };
    d.onclick = () => { closePopover(); pick(r + 1, c + 1); };
    cells.push(d);
    grid.appendChild(d);
  }
  wrap.append(grid, label);
  popover(anchor, wrap);
}

// ------------------------------------------------------------------ modal dialogs

export function dialog(title: string, body: HTMLElement, buttons: { label: string; primary?: boolean; action?: () => boolean | void }[]): () => void {
  const back = document.createElement('div');
  back.className = 'dlg-backdrop';
  const dlg = document.createElement('div');
  dlg.className = 'dlg';
  const t = document.createElement('div');
  t.className = 'dlg-title';
  t.textContent = title;
  const b = document.createElement('div');
  b.className = 'dlg-body';
  b.appendChild(body);
  const f = document.createElement('div');
  f.className = 'dlg-foot';
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey, true); };
  for (const btn of buttons) {
    const e = document.createElement('button');
    e.textContent = btn.label;
    if (btn.primary) e.className = 'primary';
    e.onclick = () => { if (btn.action?.() !== false) close(); };
    f.appendChild(e);
  }
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    else if (e.key === 'Enter' && !(e.target as HTMLElement).matches('textarea')) {
      const primary = buttons.find((x) => x.primary);
      if (primary) { e.preventDefault(); e.stopPropagation(); if (primary.action?.() !== false) close(); }
    }
    e.stopPropagation();
  };
  document.addEventListener('keydown', onKey, true);
  dlg.append(t, b);
  if (buttons.length) dlg.appendChild(f);
  back.appendChild(dlg);
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
  document.body.appendChild(back);
  setTimeout(() => (b.querySelector('input,button') as HTMLElement | null)?.focus());
  return close;
}

export function html(s: string): HTMLElement {
  const d = document.createElement('div');
  d.innerHTML = s;
  return d;
}

export function toast(msg: string) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2200);
}

export interface MenuItem { label?: string; shortcut?: string; disabled?: boolean; action?: () => void; sep?: boolean }

let openMenu: HTMLElement | null = null;

export function closeContextMenu() {
  openMenu?.remove();
  openMenu = null;
}

/** HOP-style dropdown used as a right-click menu. */
export function contextMenu(x: number, y: number, items: MenuItem[]) {
  closeContextMenu();
  const m = document.createElement('div');
  m.className = 'menu-dropdown ctx-menu';
  for (const it of items) {
    if (it.sep) { const s = document.createElement('div'); s.className = 'md-sep'; m.appendChild(s); continue; }
    const row = document.createElement('div');
    row.className = 'md-item' + (it.disabled ? ' disabled' : '');
    row.innerHTML = `<span class="md-label"></span><span class="md-shortcut"></span>`;
    row.querySelector('.md-label')!.textContent = it.label ?? '';
    row.querySelector('.md-shortcut')!.textContent = it.shortcut ?? '';
    row.addEventListener('mousedown', (e) => e.preventDefault());
    row.addEventListener('click', () => { closeContextMenu(); it.action?.(); });
    m.appendChild(row);
  }
  document.body.appendChild(m);
  m.style.left = Math.min(x, window.innerWidth - m.offsetWidth - 4) + 'px';
  m.style.top = Math.min(y, window.innerHeight - m.offsetHeight - 4) + 'px';
  openMenu = m;
  setTimeout(() => {
    const off = (e: Event) => {
      if (e instanceof MouseEvent && m.contains(e.target as Node)) return;
      closeContextMenu();
      document.removeEventListener('mousedown', off, true);
      window.removeEventListener('blur', off);
    };
    document.addEventListener('mousedown', off, true);
    window.addEventListener('blur', off);
  });
}

const HIGHLIGHTS: [string, string][] = [
  ['yellow', 'FFFF00'], ['green', '00FF00'], ['cyan', '00FFFF'], ['magenta', 'FF00FF'], ['blue', '0000FF'], ['red', 'FF0000'], ['darkBlue', '000080'], ['darkCyan', '008080'],
  ['darkGreen', '008000'], ['darkMagenta', '800080'], ['darkRed', '800000'], ['darkYellow', '808000'], ['darkGray', '808080'], ['lightGray', 'C0C0C0'], ['black', '000000'],
];

export function highlightPopover(anchor: HTMLElement, pick: (name: string | null) => void) {
  const wrap = document.createElement('div');
  const grid = document.createElement('div');
  grid.className = 'palette';
  grid.style.gridTemplateColumns = 'repeat(5, 22px)';
  for (const [name, hex] of HIGHLIGHTS) {
    const b = document.createElement('button');
    b.className = 'swatch';
    b.style.cssText = `background:#${hex};width:22px;height:22px`;
    b.title = name;
    b.onclick = () => { closePopover(); pick(name); };
    grid.appendChild(b);
  }
  const row = document.createElement('div');
  row.className = 'pop-row';
  const none = document.createElement('button');
  none.textContent = '형광펜 없음';
  none.onclick = () => { closePopover(); pick(null); };
  row.appendChild(none);
  wrap.append(grid, row);
  popover(anchor, wrap);
}

const cm = (tw: number) => (tw / TWIP_PER_CM).toFixed(2);
const fromCm = (v: string) => Math.round(Number(v) * TWIP_PER_CM);

export interface ParaDialogResult {
  jc: string; left: number; right: number; special: 'none' | 'first' | 'hanging'; by: number;
  before: number; after: number; lineMode: 'auto' | 'exact' | 'atLeast'; line: number;
}

/** 문단 모양: alignment, indents (cm), spacing (pt) and line spacing. */
export function paragraphDialog(f: ParaFmt, apply: (r: ParaDialogResult) => void) {
  const special = f.hanging ? 'hanging' : f.firstLine ? 'first' : 'none';
  const lineMode = !f.lineRule || f.lineRule === 'auto' ? 'auto' : f.lineRule === 'exact' ? 'exact' : 'atLeast';
  const lineVal = f.line == null ? 1 : lineMode === 'auto' ? +(f.line / 240).toFixed(2) : +(f.line / 20).toFixed(1);
  const body = html(`
    <div class="form">
      <label>정렬</label><select name="jc">
        <option value="left">왼쪽</option><option value="center">가운데</option><option value="right">오른쪽</option><option value="both">양쪽</option><option value="distribute">배분</option>
      </select>
      <label></label><span></span>
      <label>왼쪽 여백 (cm)</label><input type="number" step="0.1" name="left" value="${cm(f.left ?? 0)}">
      <label>오른쪽 여백 (cm)</label><input type="number" step="0.1" name="right" value="${cm(f.right ?? 0)}">
      <label>첫 줄</label><select name="special"><option value="none">없음</option><option value="first">들여쓰기</option><option value="hanging">내어쓰기</option></select>
      <label>값 (cm)</label><input type="number" step="0.1" min="0" name="by" value="${cm(f.hanging || f.firstLine || 0)}">
      <label>문단 위 (pt)</label><input type="number" step="1" min="0" name="before" value="${(f.before ?? 0) / 20}">
      <label>문단 아래 (pt)</label><input type="number" step="1" min="0" name="after" value="${(f.after ?? 0) / 20}">
      <label>줄 간격</label><select name="lineMode"><option value="auto">배수</option><option value="atLeast">최소</option><option value="exact">고정</option></select>
      <label>값</label><input type="number" step="0.05" min="0" name="line" value="${lineVal}">
    </div>`);
  const q = <T extends HTMLElement = HTMLInputElement>(n: string) => body.querySelector(`[name="${n}"]`) as T;
  q<HTMLSelectElement>('jc').value = ({ start: 'left', end: 'right' } as Record<string, string>)[f.jc ?? ''] ?? f.jc ?? 'left';
  q<HTMLSelectElement>('special').value = special;
  q<HTMLSelectElement>('lineMode').value = lineMode;
  dialog('문단 모양', body, [
    { label: '취소' },
    {
      label: '적용', primary: true, action: () => {
        apply({
          jc: q<HTMLSelectElement>('jc').value,
          left: fromCm(q('left').value), right: fromCm(q('right').value),
          special: q<HTMLSelectElement>('special').value as ParaDialogResult['special'], by: fromCm(q('by').value),
          before: Math.round(Number(q('before').value) * 20), after: Math.round(Number(q('after').value) * 20),
          lineMode: q<HTMLSelectElement>('lineMode').value as ParaDialogResult['lineMode'], line: Number(q('line').value) || 1,
        });
      },
    },
  ]);
}

const PAPERS: [string, number, number][] = [['A4', 210, 297], ['A3', 297, 420], ['A5', 148, 210], ['B5 (JIS)', 182, 257], ['B4 (JIS)', 257, 364], ['Letter', 215.9, 279.4], ['Legal', 215.9, 355.6]];

/** 쪽 설정: paper size, orientation and margins. */
export function pageSetupDialog(ps: PageSetup, apply: (ps: PageSetup) => void) {
  const mm = (tw: number) => tw / TWIP_PER_CM * 10;
  const [w, h] = [mm(Math.min(ps.w, ps.h)), mm(Math.max(ps.w, ps.h))];
  const match = PAPERS.find(([, a, b]) => Math.abs(a - w) < 1.5 && Math.abs(b - h) < 1.5);
  const body = html(`
    <div class="form">
      <label>용지</label><select name="paper">${PAPERS.map(([n, a, b]) => `<option value="${n}">${n} (${a} × ${b} mm)</option>`).join('')}<option value="custom">사용자 정의</option></select>
      <label>방향</label><select name="orient"><option value="portrait">세로</option><option value="landscape">가로</option></select>
      <label>너비 (mm)</label><input type="number" step="0.1" name="w" value="${mm(ps.w).toFixed(1)}">
      <label>높이 (mm)</label><input type="number" step="0.1" name="h" value="${mm(ps.h).toFixed(1)}">
      <label>위쪽 (cm)</label><input type="number" step="0.1" name="top" value="${cm(ps.top)}">
      <label>아래쪽 (cm)</label><input type="number" step="0.1" name="bottom" value="${cm(ps.bottom)}">
      <label>왼쪽 (cm)</label><input type="number" step="0.1" name="left" value="${cm(ps.left)}">
      <label>오른쪽 (cm)</label><input type="number" step="0.1" name="right" value="${cm(ps.right)}">
      <label>머리글 (cm)</label><input type="number" step="0.1" name="header" value="${cm(ps.header)}">
      <label>바닥글 (cm)</label><input type="number" step="0.1" name="footer" value="${cm(ps.footer)}">
    </div>`);
  const q = <T extends HTMLElement = HTMLInputElement>(n: string) => body.querySelector(`[name="${n}"]`) as T;
  const paper = q<HTMLSelectElement>('paper'), orient = q<HTMLSelectElement>('orient');
  paper.value = match ? match[0] : 'custom';
  orient.value = ps.orient;
  const sync = () => {
    const p = PAPERS.find(([n]) => n === paper.value);
    if (!p) return;
    const [a, b] = orient.value === 'landscape' ? [p[2], p[1]] : [p[1], p[2]];
    q('w').value = a.toFixed(1);
    q('h').value = b.toFixed(1);
  };
  paper.onchange = sync;
  orient.onchange = () => {
    if (paper.value === 'custom') { const t = q('w').value; q('w').value = q('h').value; q('h').value = t; } else sync();
  };
  dialog('쪽 설정', body, [
    { label: '취소' },
    {
      label: '적용', primary: true, action: () => {
        const tw = (v: string) => Math.round((Number(v) / 10) * TWIP_PER_CM);
        apply({
          ...ps,
          w: tw(q('w').value), h: tw(q('h').value), orient: orient.value as PageSetup['orient'],
          top: fromCm(q('top').value), bottom: fromCm(q('bottom').value), left: fromCm(q('left').value), right: fromCm(q('right').value),
          header: fromCm(q('header').value), footer: fromCm(q('footer').value),
        });
      },
    },
  ]);
}
