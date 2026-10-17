// Popovers (color palette, highlight, table grid) and modal dialogs.
import { PageSetup, TWIP_PER_CM } from '../engine/document';
import { PAPERS, Prefs } from './prefs';
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

export interface MenuItem { label?: string; shortcut?: string; disabled?: boolean; checked?: boolean; action?: () => void; sep?: boolean; sub?: MenuItem[] }

let openMenu: HTMLElement | null = null;

export function closeContextMenu() {
  openMenu?.remove();
  openMenu = null;
}

/** HOP-style dropdown used as a right-click menu. */
export function contextMenu(x: number, y: number, items: MenuItem[]) {
  closeContextMenu();
  const build = (list: MenuItem[], m: HTMLElement) => {
    for (const it of list) {
      if (it.sep) { const s = document.createElement('div'); s.className = 'md-sep'; m.appendChild(s); continue; }
      const row = document.createElement('div');
      row.className = 'md-item' + (it.disabled ? ' disabled' : '') + (it.checked ? ' checked' : '') + (it.sub ? ' has-sub' : '');
      row.innerHTML = `<span class="md-check"></span><span class="md-label"></span><span class="${it.sub ? 'md-arrow' : 'md-shortcut'}"></span>`;
      row.querySelector('.md-label')!.textContent = it.label ?? '';
      row.lastElementChild!.textContent = it.sub ? '▸' : it.shortcut ?? '';
      row.addEventListener('mousedown', (e) => e.preventDefault());
      if (it.sub) {
        const sub = document.createElement('div');
        sub.className = 'menu-dropdown';
        build(it.sub, sub);
        row.appendChild(sub);
        // Open to the left when there is no room on the right.
        row.addEventListener('mouseenter', () => row.classList.toggle('sub-left', row.getBoundingClientRect().right + 200 > window.innerWidth));
      } else {
        row.addEventListener('click', (e) => { e.stopPropagation(); closeContextMenu(); it.action?.(); });
      }
      m.appendChild(row);
    }
  };
  const m = document.createElement('div');
  m.className = 'menu-dropdown ctx-menu';
  build(items, m);
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


/** 쪽 설정: paper size, orientation and margins. */
export function pageSetupDialog(ps: PageSetup, apply: (ps: PageSetup, scope: string) => void, scopes: [string, string][] = [['all', '전체 문서']], scope0 = 'all') {
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
      <label>적용 대상</label><select name="scope">${scopes.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select>
    </div>`);
  const q = <T extends HTMLElement = HTMLInputElement>(n: string) => body.querySelector(`[name="${n}"]`) as T;
  const paper = q<HTMLSelectElement>('paper'), orient = q<HTMLSelectElement>('orient');
  paper.value = match ? match[0] : 'custom';
  orient.value = ps.orient;
  q<HTMLSelectElement>('scope').value = scope0;
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
        }, q<HTMLSelectElement>('scope').value);
      },
    },
  ]);
}

/** 옵션: default paper for new documents and the layout documents open in. */
export function optionsDialog(p: Prefs, apply: (p: Prefs) => void) {
  const body = html(`
    <div class="form cols2">
      <label>새 문서 용지</label><select name="paper">${PAPERS.map(([n, a, b]) => `<option value="${n}">${n} (${a} × ${b} mm)</option>`).join('')}</select>
      <label>문서를 열 때</label><select name="view"><option value="print">인쇄 모양 (용지 크기의 페이지)</option><option value="web">웹 모양 (창 너비의 웹 페이지)</option></select>
    </div>
    <p class="dlg-note">보기는 언제든 보기 탭 또는 상태 표시줄에서 바꿀 수 있습니다. 문서에 저장된 용지 크기는 바꾸지 않습니다(레이아웃 › 크기).</p>`);
  const paper = body.querySelector('[name="paper"]') as HTMLSelectElement, view = body.querySelector('[name="view"]') as HTMLSelectElement;
  paper.value = p.paper;
  view.value = p.view;
  dialog('옵션', body, [
    { label: '취소' },
    { label: '확인', primary: true, action: () => apply({ paper: paper.value, view: view.value === 'web' ? 'web' : 'print' }) },
  ]);
}

// ------------------------------------------------------------------ Word-style dialogs

export interface FontDialogResult {
  font?: string; ascii?: string; ea?: string; sz?: number; b: boolean; i: boolean; u: string; color: string | null;
  strike: boolean; dstrike: boolean; vertAlign: 'superscript' | 'subscript' | 'baseline'; smallCaps: boolean; caps: boolean; vanish: boolean;
}

const UNDERLINES: [string, string][] = [['none', '(없음)'], ['single', '실선'], ['words', '단어만'], ['double', '이중선'], ['thick', '굵은 선'], ['dotted', '점선'], ['dash', '파선'], ['wave', '물결선']];

/** 글꼴 dialog (Format ▸ Font…, Ctrl+D). */
export function fontDialog(f: { ascii?: string; ea?: string; sz?: number; b?: boolean; i?: boolean; u?: string; color?: string; strike?: boolean; dstrike?: boolean; vertAlign?: string; smallCaps?: boolean; caps?: boolean; vanish?: boolean }, fonts: string[], apply: (r: FontDialogResult) => void) {
  const style = f.b && f.i ? 'bi' : f.b ? 'b' : f.i ? 'i' : 'r';
  const body = html(`
    <div class="form">
      <label>한글 글꼴</label><input type="text" name="ea" list="dlg-fonts" value="${f.ea ?? ''}">
      <label>영어 글꼴</label><input type="text" name="ascii" list="dlg-fonts" value="${f.ascii ?? ''}">
      <label>글꼴 스타일</label><select name="style"><option value="r">보통</option><option value="i">기울임꼴</option><option value="b">굵게</option><option value="bi">굵은 기울임꼴</option></select>
      <label>크기 (pt)</label><input type="number" name="sz" step="0.5" min="1" value="${(f.sz ?? 20) / 2}">
      <label>글꼴 색</label><span><input type="color" name="color" value="#${f.color && f.color !== 'auto' ? f.color : '000000'}"> <label><input type="checkbox" name="auto" ${!f.color || f.color === 'auto' ? 'checked' : ''}> 자동</label></span>
      <label>밑줄 스타일</label><select name="u">${UNDERLINES.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select>
    </div>
    <div class="dlg-sect">효과</div>
    <div class="dlg-checks">
      <label><input type="checkbox" name="strike"> 취소선</label>
      <label><input type="checkbox" name="dstrike"> 이중 취소선</label>
      <label><input type="checkbox" name="sup"> 위 첨자</label>
      <label><input type="checkbox" name="sub"> 아래 첨자</label>
      <label><input type="checkbox" name="smallCaps"> 작은 대문자</label>
      <label><input type="checkbox" name="caps"> 모두 대문자</label>
      <label><input type="checkbox" name="vanish"> 숨김</label>
    </div>
    <div class="dlg-sect">미리 보기</div>
    <div class="font-preview">가나다 AaBbCc 123</div>
    <datalist id="dlg-fonts">${fonts.map((x) => `<option value="${x}">`).join('')}</datalist>`);
  const q = <T extends HTMLElement = HTMLInputElement>(n: string) => body.querySelector(`[name="${n}"]`) as T;
  q<HTMLSelectElement>('style').value = style;
  q<HTMLSelectElement>('u').value = UNDERLINES.some(([v]) => v === f.u) ? f.u! : f.u ? 'single' : 'none';
  for (const k of ['strike', 'dstrike', 'smallCaps', 'caps', 'vanish'] as const) q(k).checked = !!f[k];
  q('sup').checked = f.vertAlign === 'superscript';
  q('sub').checked = f.vertAlign === 'subscript';
  // Mutually exclusive pairs, like Word.
  for (const [a, b] of [['strike', 'dstrike'], ['sup', 'sub'], ['smallCaps', 'caps']]) {
    q(a).addEventListener('change', () => { if (q(a).checked) q(b).checked = false; });
    q(b).addEventListener('change', () => { if (q(b).checked) q(a).checked = false; });
  }
  const preview = body.querySelector<HTMLElement>('.font-preview')!;
  const update = () => {
    const st = q<HTMLSelectElement>('style').value;
    Object.assign(preview.style, {
      fontFamily: `'${q('ascii').value}','${q('ea').value}',sans-serif`,
      fontSize: Math.min(36, Number(q('sz').value) || 10) + 'pt',
      fontWeight: st.includes('b') ? 'bold' : 'normal', fontStyle: st.includes('i') ? 'italic' : 'normal',
      color: q('auto').checked ? '#000' : q('color').value,
      textDecorationLine: [q<HTMLSelectElement>('u').value !== 'none' ? 'underline' : '', q('strike').checked || q('dstrike').checked ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
      textTransform: q('caps').checked ? 'uppercase' : 'none', fontVariant: q('smallCaps').checked ? 'small-caps' : 'normal',
      opacity: q('vanish').checked ? '.4' : '1',
    });
  };
  body.addEventListener('input', update);
  body.addEventListener('change', update);
  q('color').addEventListener('input', () => { q('auto').checked = false; });
  update();
  dialog('글꼴', body, [
    { label: '취소' },
    {
      label: '확인', primary: true, action: () => {
        const st = q<HTMLSelectElement>('style').value;
        const ea = q('ea').value.trim(), ascii = q('ascii').value.trim();
        apply({
          font: ascii && ascii === ea ? ascii : undefined,
          ...(ascii !== ea ? { ascii, ea } as object : {}),
          sz: Number(q('sz').value) || undefined, b: st.includes('b'), i: st.includes('i'),
          u: q<HTMLSelectElement>('u').value, color: q('auto').checked ? null : q('color').value.slice(1).toUpperCase(),
          strike: q('strike').checked, dstrike: q('dstrike').checked,
          vertAlign: q('sup').checked ? 'superscript' : q('sub').checked ? 'subscript' : 'baseline',
          smallCaps: q('smallCaps').checked, caps: q('caps').checked, vanish: q('vanish').checked,
        } as FontDialogResult);
      },
    },
  ]);
}

/** 하이퍼링크 dialog (Insert ▸ Hyperlink…, Ctrl+K). */
export function hyperlinkDialog(text: string, bookmarks: string[], apply: (url: string, text: string) => void) {
  const body = html(`
    <div class="form" style="grid-template-columns:auto 1fr">
      <label>표시할 텍스트</label><input type="text" name="text" value="${text.replace(/"/g, '&quot;')}" ${text ? 'disabled' : ''}>
      <label>주소</label><input type="text" name="url" placeholder="https://">
      ${bookmarks.length ? `<label>문서 안 위치</label><select name="bm"><option value="">(웹 주소 사용)</option>${bookmarks.map((b) => `<option value="${b}">${b}</option>`).join('')}</select>` : ''}
    </div>`);
  const q = <T extends HTMLElement = HTMLInputElement>(n: string) => body.querySelector(`[name="${n}"]`) as T;
  dialog('하이퍼링크 삽입', body, [
    { label: '취소' },
    {
      label: '확인', primary: true, action: () => {
        const bm = q<HTMLSelectElement>('bm')?.value;
        let url = bm ? '#' + bm : q('url').value.trim();
        if (!url) return false;
        if (!bm && !/^[a-z][a-z0-9+.-]*:/i.test(url)) url = (url.includes('@') ? 'mailto:' : 'https://') + url;
        apply(url, text || q('text').value.trim());
      },
    },
  ]);
}

/** 날짜 및 시간 dialog. */
export function dateDialog(apply: (s: string) => void) {
  const d = new Date();
  const p2 = (n: number) => String(n).padStart(2, '0');
  const days = '일월화수목금토';
  const fmts = [
    `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`,
    `${d.getFullYear()}년 ${d.getMonth() + 1}월 ${d.getDate()}일`,
    `${d.getFullYear()}년 ${d.getMonth() + 1}월 ${d.getDate()}일 ${days[d.getDay()]}요일`,
    `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`,
    `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`,
    d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }),
    `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`,
    `${d.getHours() < 12 ? '오전' : '오후'} ${d.getHours() % 12 || 12}:${p2(d.getMinutes())}`,
  ];
  const body = html(`<select name="f" size="${fmts.length}" class="list-select">${fmts.map((f, i) => `<option value="${i}" ${i ? '' : 'selected'}>${f}</option>`).join('')}</select>`);
  const sel = body.querySelector('select')!;
  let close = () => {};
  sel.ondblclick = () => { close(); apply(fmts[Number(sel.value)]); };
  close = dialog('날짜 및 시간', body, [{ label: '취소' }, { label: '삽입', primary: true, action: () => apply(fmts[Number(sel.value)]) }]);
}

const SYMBOLS = '©®™§¶†‡•·…‰′″‹›«»“”‘’–—±×÷≠≈≤≥∞√∑∏∫∂∆∇∈∉∩∪⊂⊃∧∨¬∀∃°℃℉€£¥¢₩αβγδεθλμπσφωΩ←→↑↓↔⇒⇔①②③④⑤⑥⑦⑧⑨⑩★☆○●◎◇◆□■△▲▽▼※♠♣♥♦✓✗☐☑☎♪';

/** 기호 picker. */
export function symbolPopover(anchor: HTMLElement, pick: (s: string) => void) {
  const grid = document.createElement('div');
  grid.className = 'symbol-grid';
  for (const ch of Array.from(SYMBOLS)) {
    const b = document.createElement('button');
    b.textContent = ch;
    b.title = 'U+' + ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0');
    b.onclick = () => { closePopover(); pick(ch); };
    grid.appendChild(b);
  }
  popover(anchor, grid);
}

/** A small menu under a ribbon button (margins, orientation, size, delete…). */
export function menuPopover(anchor: HTMLElement, items: MenuItem[]) {
  const r = anchor.getBoundingClientRect();
  contextMenu(r.left, r.bottom + 2, items);
}

// ------------------------------------------------------------------ emoji

const EMOJI: [string, string][] = [
  ['표정', '😀 😃 😄 😁 😆 😅 🤣 😂 🙂 🙃 😉 😊 😇 🥰 😍 🤩 😘 😗 😚 😋 😛 😜 🤪 😝 🤑 🤗 🤭 🤫 🤔 🤐 🤨 😐 😑 😶 😏 😒 🙄 😬 😌 😔 😪 🤤 😴 😷 🤒 🤕 🤢 🤮 🥵 🥶 😵 🤯 🤠 🥳 😎 🤓 🧐 😕 😟 🙁 😮 😯 😲 😳 🥺 😦 😧 😨 😰 😥 😢 😭 😱 😖 😣 😞 😓 😩 😫 🥱 😤 😡 😠 🤬 😈 👿 💀 💩 🤡 👻 👽 🤖'],
  ['사람·손', '👋 🤚 🖐️ ✋ 🖖 👌 🤌 🤏 ✌️ 🤞 🤟 🤘 🤙 👈 👉 👆 👇 ☝️ 👍 👎 ✊ 👊 🤛 🤜 👏 🙌 👐 🤲 🤝 🙏 ✍️ 💪 🧠 👀 👁️ 👶 🧒 👦 👧 🧑 👨 👩 🧓 👴 👵 🙋 🙆 🙅 🤷 🤦 🙇 💁 👨‍💻 👩‍💻 👨‍🏫 👩‍🏫 👨‍⚕️ 👩‍⚕️ 👪'],
  ['동물·자연', '🐶 🐱 🐭 🐹 🐰 🦊 🐻 🐼 🐨 🐯 🦁 🐮 🐷 🐸 🐵 🐔 🐧 🐦 🐤 🦆 🦅 🦉 🐺 🐴 🦄 🐝 🦋 🐌 🐞 🐢 🐍 🐙 🐬 🐳 🐟 🌸 🌼 🌻 🌹 🌷 🌱 🌿 🍀 🍁 🍂 🌲 🌳 🌴 🌵 🌞 🌝 🌛 ⭐ 🌟 ✨ ⚡ 🔥 🌈 ☀️ ⛅ ☁️ 🌧️ ⛄ ❄️ 💧 🌊'],
  ['음식', '🍏 🍎 🍐 🍊 🍋 🍌 🍉 🍇 🍓 🫐 🍒 🍑 🥭 🍍 🥥 🥝 🍅 🥑 🥦 🥕 🌽 🌶️ 🥔 🍠 🥐 🍞 🧀 🥚 🍳 🥞 🥓 🍗 🍖 🌭 🍔 🍟 🍕 🥪 🌮 🍜 🍲 🍛 🍣 🍱 🥟 🍙 🍚 🍘 🍥 🍡 🍦 🍰 🎂 🍫 🍬 🍭 ☕ 🍵 🥤 🧋 🍺 🍷 🥂'],
  ['활동·여행', '⚽ 🏀 🏈 ⚾ 🎾 🏐 🏓 🏸 ⛳ 🏊 🚴 🏃 🎯 🎮 🎲 🧩 🎨 🎬 🎤 🎧 🎹 🎸 🎻 🏆 🥇 🥈 🥉 🏅 🎉 🎊 🎈 🎁 🚗 🚕 🚌 🚓 🚑 🚒 🚚 🚲 🛵 🚆 🚇 ✈️ 🚀 🛸 🚢 ⛵ 🏠 🏢 🏥 🏫 🏛️ ⛪ 🗼 🗽 🏰 ⛺ 🌋 🗻 🏖️ 🌃 🗺️ 🧭'],
  ['사물', '⌚ 📱 💻 ⌨️ 🖥️ 🖨️ 🖱️ 💾 💿 📷 📹 🎥 📞 ☎️ 📺 📻 ⏰ ⏳ 🔋 🔌 💡 🔦 🕯️ 💰 💳 💎 🔧 🔨 🛠️ ⚙️ 🔩 🧲 🔬 🔭 💊 🩺 🧪 📦 📫 ✉️ 📧 📝 ✏️ 🖊️ 📌 📍 📎 ✂️ 📁 📂 📅 📆 📊 📈 📉 📋 📚 📖 🔖 🏷️ 🔑 🔒 🔓 🛒 🎓'],
  ['기호', '❤️ 🧡 💛 💚 💙 💜 🖤 🤍 💔 💕 💖 💯 ✅ ☑️ ✔️ ❌ ❎ ⭕ ❗ ❓ ❕ ❔ ⚠️ 🚫 ⛔ 🔴 🟠 🟡 🟢 🔵 🟣 ⚫ ⚪ 🟥 🟧 🟨 🟩 🟦 🟪 ⬛ ⬜ ▶️ ⏸️ ⏹️ ⏺️ ⏭️ ⏮️ 🔼 🔽 ➡️ ⬅️ ⬆️ ⬇️ ↗️ ↘️ 🔄 🆗 🆕 🆒 🆓 🔝 ♻️ ™️ ©️ ®️ #️⃣ 1️⃣ 2️⃣ 3️⃣ 🔟'],
];

const RECENT_KEY = 'oword.recentEmoji';

function recentEmoji(): string[] {
  try { return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]').filter((x: unknown) => typeof x === 'string').slice(0, 24); } catch { return []; }
}

export function pushRecentEmoji(e: string) {
  try { localStorage.setItem(RECENT_KEY, JSON.stringify([e, ...recentEmoji().filter((x) => x !== e)].slice(0, 24))); } catch { /* storage unavailable */ }
}

/** 이모지 picker with categories and recently used emoji. */
export function emojiPopover(anchor: HTMLElement, pick: (s: string) => void) {
  const wrap = document.createElement('div');
  wrap.className = 'emoji-pop';
  const tabs = document.createElement('div');
  tabs.className = 'emoji-tabs';
  const grid = document.createElement('div');
  grid.className = 'emoji-grid';
  const recents = recentEmoji();
  const cats: [string, string[]][] = [...(recents.length ? [['최근', recents] as [string, string[]]] : []), ...EMOJI.map(([n, list]) => [n, list.split(' ')] as [string, string[]])];
  const show = (i: number) => {
    tabs.querySelectorAll('button').forEach((b, k) => b.classList.toggle('active', k === i));
    grid.innerHTML = '';
    for (const e of cats[i][1]) {
      const b = document.createElement('button');
      b.textContent = e;
      b.onclick = () => { closePopover(); pushRecentEmoji(e); pick(e); };
      grid.appendChild(b);
    }
  };
  cats.forEach(([name, list], i) => {
    const b = document.createElement('button');
    b.textContent = list[0];
    b.title = name;
    b.onclick = () => show(i);
    tabs.appendChild(b);
  });
  const hint = document.createElement('div');
  hint.className = 'emoji-hint';
  hint.textContent = /Mac/.test(navigator.platform) ? '시스템 이모지: ⌃⌘Space' : '시스템 이모지: Win + .';
  wrap.append(tabs, grid, hint);
  popover(anchor, wrap);
  show(0);
}

export type ZoomChoice = { kind: 'percent'; value: number } | { kind: 'width' | 'text' | 'page' };

/** 확대/축소 dialog (View ▸ Zoom…), like Word's: presets, fits and a custom percentage. */
export function zoomDialog(current: number, web: boolean, apply: (c: ZoomChoice) => void) {
  const pct = Math.round(current * 100);
  const opt = (v: string, label: string, disabled = false) => `<label class="zoom-opt${disabled ? ' disabled' : ''}"><input type="radio" name="z" value="${v}" ${disabled ? 'disabled' : ''}> ${label}</label>`;
  const body = html(`
    <div class="zoom-grid">
      <div class="zoom-col">
        <div class="dlg-sect" style="margin-top:0">배율</div>
        ${opt('200', '200%')}${opt('100', '100%')}${opt('75', '75%')}
      </div>
      <div class="zoom-col">
        <div class="dlg-sect" style="margin-top:0">맞춤</div>
        ${opt('width', '페이지 너비', web)}${opt('text', '텍스트 너비', web)}${opt('page', '전체 페이지', web)}
      </div>
    </div>
    <div class="zoom-custom"><label class="zoom-opt"><input type="radio" name="z" value="custom"> 사용자 지정:</label>
      <input type="number" name="pct" min="20" max="500" step="5" value="${pct}"> %</div>
    <div class="zoom-preview"><div class="zoom-sample">가나다 AaBbCc</div></div>`);
  const radios = Array.from(body.querySelectorAll<HTMLInputElement>('input[name="z"]'));
  const input = body.querySelector<HTMLInputElement>('[name="pct"]')!;
  const sample = body.querySelector<HTMLElement>('.zoom-sample')!;
  (radios.find((r) => r.value === String(pct)) ?? radios.find((r) => r.value === 'custom'))!.checked = true;
  const update = () => {
    const v = radios.find((r) => r.checked)?.value;
    const p = v && /^\d+$/.test(v) ? Number(v) : v === 'custom' ? Number(input.value) || 100 : pct;
    if (v && /^\d+$/.test(v)) input.value = v;
    sample.style.fontSize = Math.max(6, Math.min(40, (p / 100) * 14)) + 'px';
  };
  body.addEventListener('change', update);
  input.addEventListener('input', () => { radios.find((r) => r.value === 'custom')!.checked = true; update(); });
  update();
  dialog('확대/축소', body, [
    { label: '취소' },
    {
      label: '확인', primary: true, action: () => {
        const v = radios.find((r) => r.checked)?.value ?? 'custom';
        if (v === 'width' || v === 'text' || v === 'page') apply({ kind: v });
        else {
          const n = Number(v === 'custom' ? input.value : v);
          if (!(n >= 20 && n <= 500)) { input.focus(); return false; }
          apply({ kind: 'percent', value: n / 100 });
        }
      },
    },
  ]);
}

/** 창 전환 list (shared by the View menu, the ribbon and the status bar). */
export function windowMenuItems(list: { id: string; title: string; current: boolean }[], focus: (id: string) => void, newWindow: () => void): MenuItem[] {
  return [
    ...list.map((w, i) => ({ label: `${w.current ? '✓ ' : ' '}${i + 1}. ${w.title.replace(/ — Oword$/, '')}`, action: () => { if (!w.current) focus(w.id); } })),
    { sep: true },
    { label: '새 창', shortcut: 'Ctrl+Alt+N', action: newWindow },
  ];
}
