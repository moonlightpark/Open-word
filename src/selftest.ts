// In-browser end-to-end self test (dev server only): open /?selftest
// Drives the real editor with browser input events (execCommand → beforeinput/input), runs every
// command, checks layout, printing and a save/reload round trip, and reports to the page, the
// console and tests/out/selftest.json.
import { DocxDocument, Range } from './engine/document';
import { App } from './ui/app';

type Check = { name: string; ok: boolean; detail?: string; ms: number };

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const frame = () => new Promise((r) => requestAnimationFrame(() => r(null)));

export async function runSelfTest(app: App) {
  const results: Check[] = [];
  const panel = document.createElement('div');
  panel.id = 'selftest';
  panel.style.cssText = 'position:fixed;right:12px;bottom:36px;width:440px;max-height:70vh;overflow:auto;background:#fff;border:1px solid #ccd;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.25);font:12px/1.5 monospace;z-index:9000;padding:8px 10px';
  panel.innerHTML = '<b>Oword 셀프 테스트</b><div id="st-list"></div>';
  document.body.appendChild(panel);
  const list = panel.querySelector('#st-list')!;

  const ed = app.editor;
  const doc = () => app.doc!;
  const texts = () => doc().paragraphs().map((p) => doc().textOf(p));
  const find = (s: string, exact = false) => {
    const i = texts().findIndex((t) => (exact ? t === s : t.includes(s)));
    if (i < 0) throw new Error(`문단 없음: ${s} (문단 ${texts().length}개, 비슷한 문단: ${JSON.stringify(texts().filter((t) => t.includes(s.slice(0, 1))).slice(0, 3).map((t) => t.slice(0, 20)))})`);
    return i;
  };
  const caret = (pi: number, off: number) => ed.setRange({ start: { pi, off }, end: { pi, off } });
  const select = (r: Range) => ed.setRange(r);
  const type = (s: string) => { for (const ch of s) document.execCommand('insertText', false, ch); };
  // A real Enter key sends a cancelable beforeinput (execCommand does not).
  const enter = () => ed.root.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertParagraph', bubbles: true, cancelable: true }));
  const run = (id: string) => app.cmd.run(id);
  const expect = (cond: unknown, msg: string | null | undefined) => { if (!cond) throw new Error(msg ?? "실패"); };
  const eq = (a: unknown, b: unknown, what: string) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${what}: ${JSON.stringify(a)} ≠ ${JSON.stringify(b)}`); };

  async function check(name: string, fn: () => unknown) {
    const t0 = performance.now();
    let ok = true, detail: string | undefined;
    try { await fn(); } catch (e) { ok = false; detail = (e as Error).message; console.error(name, e); }
    const r = { name, ok, detail, ms: Math.round(performance.now() - t0) };
    results.push(r);
    const row = document.createElement('div');
    row.style.color = ok ? '#1b7f3b' : '#c62828';
    row.textContent = `${ok ? '✓' : '✗'} ${name}${detail ? ' — ' + detail : ''} (${r.ms}ms)`;
    list.appendChild(row);
    row.scrollIntoView({ block: 'nearest' });
    await frame();
  }

  const open = async (url: string) => {
    const res = await fetch(url);
    await app.open(new Uint8Array(await res.arrayBuffer()), url.split('/').pop()!);
    await sleep(50);
  };

  // Every block must sit inside a page's text area unless it is taller than a page.
  const layoutOk = () => {
    const g = ed.geometry();
    const stride = (g.H + ed.gapPt) * (4 / 3), content = (g.H - g.top - g.bottom) * (4 / 3);
    for (const b of ed.layout) {
      if (b.el instanceof HTMLTableElement) {
        for (const tr of Array.from(b.el.rows)) {
          if (tr.classList.contains('pg-spacer')) continue;
          const top = b.y + tr.offsetTop, bot = top + tr.offsetHeight;
          const pg = Math.floor((top + 1) / stride);
          if (tr.offsetHeight <= content && bot > pg * stride + content + 1) return `표 행이 쪽 아래 여백을 넘음: "${tr.textContent?.slice(0, 20)}" (쪽 ${pg + 1})`;
        }
        continue;
      }
      if (b.h > content) continue;
      const pg = Math.floor((b.y + 1) / stride);
      if (b.y + b.h > pg * stride + content + 1) return `블록이 쪽 아래 여백을 넘음: "${b.el.textContent?.slice(0, 20)}" (쪽 ${pg + 1})`;
    }
    return null;
  };

  // ------------------------------------------------------------------ open & render
  await check('기능 점검 문서 열기', async () => {
    await open('./features.docx');
    expect(ed.pages >= 10, `쪽 수 ${ed.pages}`);
  });
  await check('쪽 배치: 모든 블록과 표 행이 여백 안에 있음', () => { const e = layoutOk(); expect(!e, e); });
  await check('긴 표가 여러 쪽으로 나뉨 (쪽 나눔 행 삽입)', () => {
    const t = Array.from(ed.root.querySelectorAll<HTMLTableElement>(':scope > table')).find((x) => x.textContent?.includes('항목 40'))!;
    expect(t.querySelectorAll('tr.pg-spacer').length >= 1, '나눔 행 없음');
    const box = ed.layout.find((b) => b.el === t)!;
    expect(box.endPage > box.page, '한 쪽에만 있음');
  });
  await check('머리글/바닥글: 첫 쪽·홀수·짝수, 쪽 번호', () => {
    const pages = Array.from(ed.frames.children) as HTMLElement[];
    eq(pages[0].querySelector('.page-footer')?.textContent, '표지 바닥글 (첫 쪽 전용)', '1쪽 바닥글');
    expect(pages[1].querySelector('.page-header')?.textContent?.includes('짝수'), '2쪽 머리글');
    expect(pages[2].querySelector('.page-header')?.textContent?.includes('홀수'), '3쪽 머리글');
    eq(pages[4].querySelector('.page-footer')?.textContent, `쪽 5 / ${ed.pages}`, '5쪽 바닥글');
  });
  await check('본문의 PAGE / NUMPAGES 필드가 실제 쪽을 표시', () => {
    const pg = ed.root.querySelector<HTMLElement>('.w-fld[data-fld="PAGE"]')!;
    const box = ed.layout.find((b) => b.el.contains(pg))!;
    eq(pg.textContent, String(box.page + 1), 'PAGE');
    eq(ed.root.querySelector('.w-fld[data-fld="NUMPAGES"]')!.textContent, String(ed.pages), 'NUMPAGES');
  });
  await check('탐색 창에 제목 목록', () => {
    const items = Array.from(document.querySelectorAll('#outline-list .outline-item')).map((d) => d.textContent);
    expect(items.includes('4.4 여러 쪽에 걸친 긴 표') && items.includes('9. 긴 본문'), items.join(','));
  });
  await check('탐색 창 클릭 → 해당 문단으로 이동', () => {
    const it = Array.from(document.querySelectorAll<HTMLElement>('#outline-list .outline-item')).find((d) => d.textContent === '3. 목록')!;
    it.click();
    eq(ed.getRange()?.start.pi, Number(it.dataset.pi), '커서 문단');
  });

  // ------------------------------------------------------------------ typing
  await check('글자 입력 (한글·영문·공백 여러 칸)', () => {
    const pi = find('표지 끝.');
    caret(pi, 0);
    type('입력  테스트 ABC ');
    ed.commit();
    eq(texts()[pi], '입력  테스트 ABC 표지 끝.\n', '문단 (끝의 \\n은 쪽 나누기)');
  });
  await check('Enter: 문단 나누기, 제목 다음은 표준 스타일', () => {
    const pi = find('1.1 기본 모양', true);
    caret(pi, doc().lengthOf(doc().para(pi)!));
    const n = doc().paragraphs().length;
    enter();
    type('새 문단');
    ed.commit();
    eq(doc().paragraphs().length, n + 1, '문단 수');
    eq(texts()[pi + 1], '새 문단', '새 문단');
    eq(doc().pStyleOf(doc().para(pi + 1)!), null, '스타일');
  });
  await check('Enter: 목록 다음 항목, 빈 항목에서 목록 끝내기', () => {
    const pi = find('세부 ii');
    caret(pi, doc().lengthOf(doc().para(pi)!));
    enter();
    type('세부 iii');
    ed.commit();
    eq(doc().labels().get(doc().para(pi + 1)!)?.text, 'iii)', '번호');
    enter();
    enter();
    ed.commit();
    expect(!doc().labels().get(doc().para(pi + 2)!), '목록이 끝나지 않음');
  });
  await check('Backspace: 목록 머리 지우기', () => {
    const pi = find('다섯째');
    caret(pi, 0);
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true }));
    expect(!doc().labels().get(doc().para(pi)!), '번호가 남음');
    run('edit:undo');
    expect(doc().labels().get(doc().para(pi)!), '되돌리기 실패');
  });
  await check('Tab: 탭 문자, 목록 수준, 다음 셀', () => {
    const pi = find('외부 링크');
    caret(pi, 0);
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
    expect(texts()[pi].startsWith('\t'), '탭 없음');
    const li = find('첫째 절');
    caret(li, 0);
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
    eq(doc().paraFmt(doc().para(li)!).ilvl, 2, '수준');
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }));
    eq(doc().paraFmt(doc().para(li)!).ilvl, 1, '수준 복구');
    const c = find('1분기');
    caret(c, 0);
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
    eq(ed.getRange()?.start.pi, find('2분기'), '다음 셀');
  });

  // ------------------------------------------------------------------ character formatting
  const fmtOf = (pi: number, off: number) => doc().fmtAt({ pi, off })!;
  await check('글자 모양: 진하게/기울임/밑줄/취소선/첨자 (선택 영역)', () => {
    const pi = find('왼쪽 정렬 문단');
    select({ start: { pi, off: 0 }, end: { pi, off: 2 } });
    for (const id of ['format:bold', 'format:italic', 'format:underline', 'format:strike']) run(id);
    const f = fmtOf(pi, 1);
    expect(f.b && f.i && f.u === 'single' && f.strike, JSON.stringify(f));
    expect(!fmtOf(pi, 4).b, '범위 밖까지 적용됨');
    run('format:superscript');
    eq(fmtOf(pi, 1).vertAlign, 'superscript', '위 첨자');
    run('format:subscript');
    eq(fmtOf(pi, 1).vertAlign, 'subscript', '아래 첨자');
    run('format:clear');
    const c = fmtOf(pi, 1);
    expect(!c.b && !c.i && !c.vertAlign, '서식 지우기');
  });
  await check('글자 크기 크게/작게, 글꼴, 색, 형광펜', () => {
    const pi = find('가운데 정렬 문단');
    select({ start: { pi, off: 0 }, end: { pi, off: 3 } });
    run('format:grow');
    eq(fmtOf(pi, 1).sz, 21, '크게 (10 → 10.5pt)');
    run('format:shrink');
    run('format:shrink');
    eq(fmtOf(pi, 1).sz, 18, '작게 (10.5 → 10 → 9pt)');
    const fn = document.getElementById('font-name') as HTMLInputElement;
    fn.value = '바탕';
    fn.dispatchEvent(new Event('change'));
    eq(fmtOf(pi, 1).ea, '바탕', '글꼴');
    select({ start: { pi, off: 0 }, end: { pi, off: 3 } });
    document.getElementById('btn-text-color')!.dispatchEvent(new MouseEvent('click', { clientX: 0, detail: 1 }));
    eq(fmtOf(pi, 1).color, 'C00000', '글자 색');
    document.getElementById('btn-highlight')!.dispatchEvent(new MouseEvent('click', { clientX: 0, detail: 1 }));
    eq(fmtOf(pi, 1).highlight, 'FFFF00', '형광펜');
  });
  await check('커서만 있을 때: 단어 전체에 적용', () => {
    const pi = find('오른쪽 정렬 문단');
    caret(pi, 1); // inside "오른쪽"
    run('format:bold');
    expect(fmtOf(pi, 1).b && fmtOf(pi, 3).b && !fmtOf(pi, 5).b, '단어 범위');
  });
  await check('커서만 있을 때(단어 밖): 다음 입력에 적용', () => {
    const pi = find('오른쪽 정렬 문단');
    const len = doc().lengthOf(doc().para(pi)!);
    caret(pi, len);
    run('format:italic');
    type('기울임추가');
    ed.commit();
    expect(fmtOf(pi, len + 2).i, '다음 입력이 기울임 아님');
    expect(!fmtOf(pi, len - 1).i, '앞 글자까지 바뀜');
    eq(texts()[pi].endsWith('기울임추가'), true, '텍스트');
  });

  // ------------------------------------------------------------------ paragraph formatting
  await check('문단: 정렬 4가지, 줄 간격, 들여쓰기/내어쓰기', () => {
    const pi = find('양쪽 정렬:');
    caret(pi, 3);
    for (const [id, jc] of [['format:align-center', 'center'], ['format:align-right', 'right'], ['format:align-justify', 'both'], ['format:align-left', 'left']]) {
      run(id);
      eq(doc().paraFmt(doc().para(pi)!).jc, jc, id);
    }
    run('format:line-1.5');
    eq(doc().paraFmt(doc().para(pi)!).line, 360, '줄 간격');
    run('format:indent');
    run('format:indent');
    eq(doc().paraFmt(doc().para(pi)!).left, doc().defaultTabStop * 2, '들여쓰기');
    run('format:outdent');
    eq(doc().paraFmt(doc().para(pi)!).left, doc().defaultTabStop, '내어쓰기');
  });
  await check('문단: 스타일 바꾸기 (제목 1/2/3, 표준) → 탐색 창 갱신', async () => {
    const pi = find('문단 위 12pt');
    caret(pi, 0);
    run('style:Heading2');
    eq(doc().pStyleOf(doc().para(pi)!), 'Heading2', '제목 2');
    await sleep(500);
    expect(Array.from(document.querySelectorAll('#outline-list .outline-item')).some((d) => d.textContent?.startsWith('문단 위 12pt')), '탐색 창');
    run('style:normal');
    eq(doc().pStyleOf(doc().para(pi)!), null, '표준');
  });
  await check('글머리표/문단 번호 켜고 끄기 (여러 문단)', () => {
    const a = find('줄 간격 1줄'), b = find('줄 간격 2줄');
    select({ start: { pi: a, off: 0 }, end: { pi: b, off: 1 } });
    run('format:numbering');
    eq([a, a + 1, b].map((i) => doc().labels().get(doc().para(i)!)?.text), ['1.', '2.', '3.'], '번호');
    run('format:bullet');
    expect([a, a + 1, b].every((i) => doc().numbering.isBullet(doc().paraFmt(doc().para(i)!).numId!, 0)), '글머리표');
    run('format:bullet');
    expect([a, a + 1, b].every((i) => !doc().labels().get(doc().para(i)!)), '끄기');
  });

  // ------------------------------------------------------------------ insert
  await check('쪽 나누기 넣기 → 다음 문단이 새 쪽에서 시작', () => {
    const pi = find('2.4 테두리, 음영, 탭', true);
    const pages = ed.pages;
    caret(pi, 0);
    run('insert:page-break');
    const p = ed.paragraphElement(find('2.4 테두리, 음영, 탭', true))!;
    const box = ed.layout.find((b) => b.el === p)!;
    const g = ed.geometry();
    const stride = (g.H + ed.gapPt) * (4 / 3);
    expect(Math.abs(box.y - box.mt - box.page * stride) < 2, '쪽 맨 위가 아님');
    expect(ed.pages >= pages, '쪽 수');
  });
  await check('표 넣기 → 첫 셀에 커서, 입력', async () => {
    const pi = find('문서 끝.');
    caret(pi, 0);
    const n = ed.root.querySelectorAll('table').length;
    run('insert:table');
    (document.querySelector('.table-grid')!.children[2 * 10 + 3] as HTMLElement).click(); // 3×4
    eq(ed.root.querySelectorAll('table').length, n + 1, '표 수');
    type('새 표');
    ed.commit();
    const cell = doc().cellOf(ed.getRange()!.start.pi)!;
    eq(cell.tc.textContent, '새 표', '셀 내용');
    eq(cell.tbl.getElementsByTagNameNS(W, 'tr').length, 3, '줄');
    eq(cell.tbl.getElementsByTagNameNS(W, 'gridCol').length, 4, '칸');
  });
  await check('표 편집: 줄/칸 넣기·지우기', () => {
    const tbl = () => doc().cellOf(find('새 표'))!.tbl;
    const size = () => [tbl().getElementsByTagNameNS(W, 'tr').length, tbl().getElementsByTagNameNS(W, 'gridCol').length];
    caret(find('새 표'), 0);
    run('table:row-below'); // the caret moves into the new row
    eq(size(), [4, 4], '줄 넣기');
    expect(doc().cellOf(ed.getRange()!.start.pi)!.tr !== doc().cellOf(find('새 표'))!.tr, '커서가 새 줄에 없음');
    run('table:del-row');
    eq(size(), [3, 4], '줄 지우기');
    caret(find('새 표'), 0);
    run('table:col-right'); // the caret moves into the new column
    eq(size(), [3, 5], '칸 넣기');
    run('table:del-col');
    eq(size(), [3, 4], '칸 지우기');
    caret(find('새 표'), 0);
    run('table:row-above');
    eq(size(), [4, 4], '위에 줄 넣기');
    run('table:del-row');
    run('table:col-left');
    eq(size(), [3, 5], '왼쪽에 칸 넣기');
    run('table:del-col');
    eq(size(), [3, 4], '원래대로');
  });
  await check('표 셀 세로 정렬: 위/가운데/아래 (메뉴·서식 바)', () => {
    const pi = find('새 표');
    caret(pi, 0);
    expect(!document.getElementById('btn-valign-center')!.classList.contains('disabled'), '버튼 비활성');
    run('table:valign-center');
    eq(doc().cellVAlign(find('새 표')), 'center', '가운데');
    const td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    eq(td.style.verticalAlign, 'middle', 'CSS');
    document.getElementById('btn-valign-bottom')!.click();
    eq(doc().cellVAlign(find('새 표')), 'bottom', '아래');
    run('table:valign-top');
    eq(doc().cellVAlign(find('새 표')), 'top', '위');
    caret(find('문서 끝.'), 0);
    expect(!app.cmd.enabled('table:valign-top'), '표 밖에서 활성화됨');
  });
  await check('표 테두리 끌기: 칸 너비', async () => {
    const td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    const table = td.closest('table')!;
    const tbl = ed.reg.dom.get(table)!;
    const before = doc().columnWidths(tbl);
    const r = td.getBoundingClientRect();
    const y = (r.top + r.bottom) / 2;
    td.dispatchEvent(new MouseEvent('mousedown', { clientX: r.right, clientY: y, bubbles: true, button: 0, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: r.right + 30 * ed.zoom, clientY: y, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: r.right + 30 * ed.zoom, clientY: y }));
    const after = doc().columnWidths(doc().cellOf(find('새 표'))!.tbl);
    eq(Math.round(after[0] - before[0]), 450, '첫 칸 +30px(450twip)');
    eq(Math.round(before[1] - after[1]), 450, '둘째 칸 -450twip');
    eq(after.reduce((a, b) => a + b, 0), before.reduce((a, b) => a + b, 0), '표 너비 유지');
  });
  await check('표 테두리 끌기: 줄 높이', () => {
    const td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    const h0 = (td.parentElement as HTMLElement).offsetHeight;
    const r = td.getBoundingClientRect();
    const x = (r.left + r.right) / 2;
    td.dispatchEvent(new MouseEvent('mousedown', { clientX: x, clientY: r.bottom - 1, bubbles: true, button: 0, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: x, clientY: r.bottom + 40 * ed.zoom, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: x, clientY: r.bottom + 40 * ed.zoom }));
    const h1 = (ed.paragraphElement(find('새 표'))!.closest('tr') as HTMLElement).offsetHeight;
    expect(Math.abs(h1 - (h0 + 40)) <= 2, `${h0} → ${h1}`);
    eq(doc().undoLabel(), '행 높이', '실행 취소 항목');
  });
  await check('그림 넣기(붙여넣기) → 선택·크기 조절·지우기', async () => {
    const pi = find('문서 끝.');
    caret(pi, 0);
    const c = document.createElement('canvas');
    c.width = 200; c.height = 100;
    const g = c.getContext('2d')!;
    g.fillStyle = '#2b6fd6'; g.fillRect(0, 0, 200, 100);
    const blob: Blob = await new Promise((r) => c.toBlob((b) => r(b!), 'image/png'));
    const dt = new DataTransfer();
    dt.items.add(new File([blob], 'paste.png', { type: 'image/png' }));
    ed.root.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    await sleep(300);
    const img = ed.paragraphElement(find('문서 끝.'))!.querySelector('img')!;
    expect(img, '그림 없음');
    eq(Math.round(img.offsetWidth), 200, '그림 너비(px)');
    img.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    const h = document.querySelector<HTMLElement>('.img-handle')!;
    const hr = h.getBoundingClientRect();
    h.dispatchEvent(new MouseEvent('mousedown', { clientX: hr.left, clientY: hr.top, bubbles: true, button: 0 }));
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: hr.left - 100 * ed.zoom, clientY: hr.top }));
    window.dispatchEvent(new MouseEvent('mouseup', {}));
    const img2 = ed.paragraphElement(find('문서 끝.'))!.querySelector('img')!;
    eq(Math.round(img2.offsetWidth), 100, '줄인 너비');
    eq(Math.round(img2.offsetHeight), 50, '비율 유지');
    img2.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    ed.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }));
    expect(!ed.paragraphElement(find('문서 끝.'))!.querySelector('img'), '지워지지 않음');
  });

  // ------------------------------------------------------------------ clipboard, find, undo
  await check('복사/붙여넣기: 서식 유지 (문서 안)', () => {
    const pi = find('글자 음영');
    select({ start: { pi, off: 0 }, end: { pi, off: 6 } }); // "노랑 형광펜"
    document.dispatchEvent(new ClipboardEvent('copy', { bubbles: true }));
    ed.root.dispatchEvent(new ClipboardEvent('copy', { bubbles: true }));
    const target = find('문서 끝.');
    caret(target, 0);
    const dt = new DataTransfer();
    dt.setData('text/plain', '노랑 형광펜');
    ed.root.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    expect(texts()[target].startsWith('노랑 형광펜'), texts()[target]);
    eq(fmtOf(target, 2).highlight, 'FFFF00', '형광펜 유지');
  });
  await check('붙여넣기: 외부 텍스트 여러 줄 → 문단', () => {
    const target = find('문서 끝.');
    caret(target, 0);
    const n = doc().paragraphs().length;
    const dt = new DataTransfer();
    dt.setData('text/plain', '첫 줄\n둘째 줄\t탭\n');
    ed.root.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    eq(doc().paragraphs().length, n + 2, '문단 수');
    eq(texts().slice(target, target + 2), ['첫 줄', '둘째 줄\t탭'], '내용');
  });
  await check('찾기/바꾸기: 모두 바꾸기', () => {
    const n = doc().findAll('긴 표의 행입니다', true).length;
    expect(n === 40, `찾기 ${n}`);
    eq(doc().replaceAll('긴 표의 행입니다', '바꾼 행', true), 40, '바꾼 개수');
    ed.render();
    expect(ed.root.textContent!.includes('바꾼 행'), '화면 갱신');
    run('edit:undo');
    eq(doc().findAll('바꾼 행', true).length, 0, '되돌리기');
  });
  await check('되돌리기/다시 실행 연속 20번', () => {
    const before = doc().xmlText();
    const steps: string[] = [];
    for (let i = 0; i < 20 && doc().canUndo(); i++) { steps.push(doc().undoLabel()!); run('edit:undo'); }
    expect(steps.length === 20, `실행 취소 단계 ${steps.length}`);
    for (let i = 0; i < 20; i++) run('edit:redo');
    eq(doc().xmlText() === before, true, '다시 실행 후 같은 문서');
  });

  // ------------------------------------------------------------------ Word-style menus and ribbon
  await check('메뉴 막대: Word 구성, 모든 항목에 명령 연결', () => {
    eq(Array.from(document.querySelectorAll('#menu-bar > .menu-item > .menu-title')).map((t) => t.textContent), ['파일', '편집', '보기', '삽입', '서식', '도구', '표', '도움말'], '메뉴');
    const missing = Array.from(document.querySelectorAll<HTMLElement>('[data-cmd]')).map((e) => e.dataset.cmd!).filter((id) => !app.cmd.has(id));
    expect(!missing.length, '명령 없음: ' + missing.join(', '));
    const checks = Array.from(document.querySelectorAll<HTMLElement>('[data-check]')).map((e) => e.dataset.check!).filter((id) => !app.cmd.has(id));
    expect(!checks.length, '체크 명령 없음: ' + checks.join(', '));
    const tbl = document.querySelector<HTMLElement>('[data-menu="table"] .menu-title')!;
    tbl.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(document.querySelector('[data-menu="table"]')!.classList.contains('open'), '표 메뉴가 열리지 않음');
    expect(document.querySelector('[data-menu="table"] .md-item[data-cmd="table:del-row"]')!.classList.contains('disabled'), '표 밖에서 행 삭제가 활성');
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(!document.querySelector('.menu-item.open'), '메뉴가 닫히지 않음');
  });
  await check('리본: 탭 전환, 표 안에서만 표 레이아웃 탭', async () => {
    for (const t of ['insert', 'layout', 'review', 'view', 'home']) {
      document.querySelector<HTMLElement>(`.rb-tab[data-tab="${t}"]`)!.click();
      expect(document.querySelector<HTMLElement>(`.rb-panel[data-panel="${t}"]`)!.classList.contains('active'), t);
    }
    const tab = document.querySelector<HTMLElement>('.rb-tab[data-tab="table"]')!;
    caret(find('문서 끝.'), 0);
    await sleep(60);
    expect(tab.hidden, '표 밖인데 표 탭이 보임');
    caret(find('구분', true), 0);
    await sleep(60);
    expect(!tab.hidden, '표 안인데 표 탭이 없음');
    tab.click();
    run('table:align-center-center');
    eq([doc().cellVAlign(find('구분', true)), doc().paraFmt(doc().para(find('구분', true))!).jc], ['center', 'center'], '3×3 맞춤');
    await sleep(60);
    expect(document.getElementById('btn-cell-center-center')!.classList.contains('active'), '맞춤 버튼 표시');
    caret(find('문서 끝.'), 0);
    await sleep(60);
    expect(tab.hidden && document.querySelector('.rb-tab[data-tab="home"]')!.classList.contains('active'), '표를 벗어나면 홈 탭으로');
    run('view:toggle-ribbon');
    expect(document.getElementById('ribbon')!.classList.contains('collapsed'), '리본 접기');
    run('view:toggle-ribbon');
  });
  await check('홈 탭: 스타일 갤러리', async () => {
    const pi = find('문단 위 12pt');
    caret(pi, 0);
    await sleep(60);
    const card = document.querySelector<HTMLElement>('#style-gallery .style-card[data-style="Heading1"]')!;
    expect(card, '제목 1 카드 없음');
    card.click();
    eq(doc().pStyleOf(doc().para(pi)!), 'Heading1', '스타일');
    await sleep(60);
    expect(document.querySelector('#style-gallery .style-card.active')?.getAttribute('data-style') === 'Heading1', '활성 카드');
    run('edit:undo');
  });
  await check('서식 ▸ 글꼴… 대화상자', () => {
    const pi = find('가운데 정렬 문단');
    select({ start: { pi, off: 0 }, end: { pi, off: 2 } });
    run('format:font');
    const dlg = document.querySelector('.dlg')!;
    const q = (n: string) => dlg.querySelector<HTMLInputElement>(`[name="${n}"]`)!;
    q('ea').value = '굴림'; q('ascii').value = 'Georgia';
    q('sz').value = '15';
    (dlg.querySelector('[name="style"]') as HTMLSelectElement).value = 'bi';
    (dlg.querySelector('[name="u"]') as HTMLSelectElement).value = 'double';
    q('dstrike').checked = true; q('smallCaps').checked = true;
    Array.from(dlg.querySelectorAll('button')).find((b) => b.textContent === '확인')!.click();
    const f = fmtOf(pi, 1);
    eq([f.ea, f.ascii, f.sz, f.b, f.i, f.u, f.dstrike, f.smallCaps], ['굴림', 'Georgia', 30, true, true, 'double', true, true], '글꼴');
  });
  await check('대/소문자 바꾸기, 날짜, 기호, 줄 바꿈', () => {
    const pi = find('외부 링크');
    const t = texts()[pi];
    const s0 = t.indexOf('Oword');
    select({ start: { pi, off: s0 }, end: { pi, off: s0 + 5 } });
    run('format:case-upper');
    expect(texts()[pi].includes('OWORD'), '대문자');
    caret(find('문서 끝.'), 0);
    run('insert:date');
    Array.from(document.querySelectorAll('.dlg button')).find((b) => b.textContent === '삽입')!.dispatchEvent(new MouseEvent('click'));
    expect(/^\d{4}-\d{2}-\d{2}\D/.test(texts()[find('문서 끝.')]), '날짜: ' + texts()[find('문서 끝.')]);
    run('insert:symbol');
    (document.querySelector<HTMLElement>('.symbol-grid button[title="U+2605"]'))!.click(); // ★
    expect(texts()[find('문서 끝.')].includes('★'), '기호');
    const n = doc().paragraphs().length;
    run('insert:line-break');
    eq(doc().paragraphs().length, n, '줄 바꿈은 문단을 나누지 않음');
    expect(texts()[find('문서 끝.')].includes('\n'), '줄 바꿈');
  });
  await check('삽입 ▸ 하이퍼링크… (Ctrl+K)', () => {
    const pi = find('오른쪽 정렬 문단');
    select({ start: { pi, off: 0 }, end: { pi, off: 3 } });
    run('insert:hyperlink');
    const dlg = document.querySelector('.dlg')!;
    dlg.querySelector<HTMLInputElement>('[name="url"]')!.value = 'example.org';
    Array.from(dlg.querySelectorAll('button')).find((b) => b.textContent === '확인')!.click();
    const a = ed.paragraphElement(find('오른쪽 정렬 문단'))!.querySelector<HTMLElement>('a.w-hyperlink')!;
    eq([a.textContent, a.dataset.href], ['오른쪽', 'https://example.org'], '링크');
  });
  await check('검토 ▸ 변경 내용 모두 적용 / 거부', () => {
    eq(doc().revisionCount(), 2, '변경 내용 수');
    run('review:reject-all');
    expect(texts().some((t) => t.includes('삭제된 글자 끝.')), '거부 → 삭제 취소');
    run('edit:undo');
    run('review:accept-all');
    eq(doc().revisionCount(), 0, '적용');
    expect(!ed.root.querySelector('.w-del, .w-ins'), '화면에 표시가 남음');
    run('edit:undo');
    eq(doc().revisionCount(), 2, '되돌리기');
  });
  await check('레이아웃 탭: 여백·방향·크기 메뉴, 단락 들여쓰기 칸', () => {
    document.querySelector<HTMLElement>('.rb-tab[data-tab="layout"]')!.click();
    run('layout:margins');
    Array.from(document.querySelectorAll<HTMLElement>('.ctx-menu .md-item')).find((m) => m.textContent!.startsWith('좁게'))!.click();
    eq([doc().pageSetup().left, doc().pageSetup().top], [720, 720], '좁게');
    run('layout:orientation');
    Array.from(document.querySelectorAll<HTMLElement>('.ctx-menu .md-item')).find((m) => m.textContent === '가로')!.click();
    eq(doc().pageSetup().orient, 'landscape', '가로');
    run('layout:size');
    Array.from(document.querySelectorAll<HTMLElement>('.ctx-menu .md-item')).find((m) => m.textContent!.startsWith('Letter'))!.click();
    eq([doc().pageSetup().w, doc().pageSetup().h], [15840, 12240], 'Letter 가로');
    run('edit:undo'); run('edit:undo'); run('edit:undo');
    eq(doc().pageSetup().w, 11906, '되돌리기');
    const pi = find('왼쪽 정렬 문단');
    caret(pi, 0);
    const inp = document.getElementById('lay-left') as HTMLInputElement;
    inp.value = '1.5';
    inp.dispatchEvent(new Event('change'));
    eq(Math.round(doc().paraFmt(doc().para(pi)!).left!), 850, '왼쪽 1.5cm');
    document.querySelector<HTMLElement>('.rb-tab[data-tab="home"]')!.click();
  });
  await check('표 눈금선 보기 끄기/켜기', () => {
    const t = Array.from(ed.root.querySelectorAll<HTMLTableElement>('table.nogrid'))[0];
    expect(getComputedStyle(t.querySelector('td')!).outlineStyle === 'dashed', '눈금선 표시');
    run('view:gridlines');
    eq(getComputedStyle(t.querySelector('td')!).outlineStyle, 'none', '눈금선 숨김');
    run('view:gridlines');
  });

  // ------------------------------------------------------------------ view, page setup, print
  await check('보기: 확대/축소, 쪽 맞춤, 문단 부호, 눈금자, 탐색 창', () => {
    run('view:zoom-100');
    run('view:zoom-in');
    expect(ed.zoom > 1, '확대');
    run('view:zoom-page');
    expect(ed.zoom < 1, '쪽 맞춤');
    run('view:zoom-width');
    run('view:toggle-marks');
    expect(ed.root.classList.contains('show-marks'), '문단 부호');
    run('view:toggle-marks');
    run('view:toggle-ruler');
    expect(document.getElementById('editor-area')!.classList.contains('no-ruler'), '눈금자');
    run('view:toggle-ruler');
    run('view:toggle-outline');
    run('view:toggle-outline');
    run('view:zoom-100');
  });
  await check('쪽 설정: 가로 방향·여백 → 다시 배치', () => {
    const ps = doc().pageSetup();
    const pages = ed.pages;
    doc().setPageSetup({ ...ps, w: ps.h, h: ps.w, orient: 'landscape', left: 720, right: 720 });
    ed.render();
    eq(ed.geometry().W > ed.geometry().H, true, '가로 방향');
    const e = layoutOk();
    expect(!e, e);
    run('edit:undo');
    eq(ed.pages, pages, '되돌린 쪽 수');
  });
  await check('인쇄 배치: 쪽마다 한 장, 용지 크기', () => {
    const real = window.print;
    window.print = () => {};
    try {
      run('file:print');
    } finally { setTimeout(() => { window.print = real; }, 200); }
    const root = document.getElementById('print-root')!;
    eq(root.querySelectorAll('.print-page').length, ed.pages, '장 수');
    expect(root.querySelector('style')!.textContent!.includes('595.3pt 841.9pt'), 'A4');
    expect(root.querySelectorAll('.print-page')[ed.pages - 1].textContent!.includes('문서 끝'), '마지막 장');
    root.innerHTML = '';
    ed.paginate();
  });

  // ------------------------------------------------------------------ save & reload
  await check('저장 → 다시 열기: 내용·서식·개체 보존', async () => {
    ed.commit();
    const bytes = await doc().saveBytes();
    await fetch('/__debug/save', { method: 'POST', body: bytes as BlobPart });
    const re = await DocxDocument.load(bytes);
    const count = (d: DocxDocument, l: string) => d.body.getElementsByTagNameNS(W, l).length;
    for (const l of ['tbl', 'drawing', 'hyperlink', 'ins', 'del', 'commentReference', 'sdt', 'footnoteReference', 'sym', 'sectPr', 'bookmarkStart']) {
      eq(count(re, l), count(doc(), l), l);
    }
    eq(re.paragraphs().map((p) => re.textOf(p)), texts(), '모든 문단 텍스트');
    eq(re.paragraphs().map((p) => re.labels().get(p)?.text ?? ''), doc().paragraphs().map((p) => doc().labels().get(p)?.text ?? ''), '목록 번호');
    for (const part of ['word/footnotes.xml', 'word/comments.xml']) expect(re.pkg.getXml(part), part + ' 없음');
  });
  await check('새 문서 → 입력 → 저장 → 다시 열기', async () => {
    await open('./blank.docx');
    caret(0, 0);
    type('새 문서 첫 줄');
    enter();
    run('format:bullet');
    type('목록');
    ed.commit();
    const re = await DocxDocument.load(await doc().saveBytes());
    eq(re.paragraphs().map((p) => re.textOf(p)), ['새 문서 첫 줄', '목록'], '텍스트');
    expect(re.labels().get(re.para(1)!), '목록 번호');
  });

  const failed = results.filter((r) => !r.ok);
  const summary = `${results.length - failed.length} / ${results.length} 통과`;
  const head = document.createElement('div');
  head.style.cssText = `margin-top:6px;font-weight:bold;color:${failed.length ? '#c62828' : '#1b7f3b'}`;
  head.textContent = summary;
  list.appendChild(head);
  console.log('[selftest]', summary, failed);
  await fetch('/__debug/selftest', { method: 'POST', body: JSON.stringify({ summary, results }, null, 1) }).catch(() => {});
  (window as unknown as { selftest: unknown }).selftest = { summary, results };
  return results;
}
