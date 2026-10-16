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
  await check('표 셀 세로 정렬: 위/가운데/아래 (메뉴·리본·홈 탭)', async () => {
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
    // The same buttons on the 홈 tab (next to paragraph alignment).
    document.getElementById('home-valign-center')!.click();
    eq(doc().cellVAlign(find('새 표')), 'center', '홈 탭 가운데');
    await frame();
    expect(document.getElementById('home-valign-center')!.classList.contains('active'), '홈 탭 단추 상태');
    run('table:valign-top');
    caret(find('문서 끝.'), 0);
    await frame();
    expect(document.getElementById('home-valign-top')!.classList.contains('disabled'), '표 밖에서 홈 탭 단추가 활성');
    expect(!app.cmd.enabled('table:valign-top'), '표 밖에서 활성화됨');
  });
  await check('오른쪽 클릭 메뉴: 선택한 표 텍스트의 가로 정렬 · 세로 정렬', async () => {
    const ctxItem = (path: string[]) => {
      let scope: Element = document.querySelector('.ctx-menu')!;
      let row: HTMLElement | undefined;
      for (const label of path) {
        row = Array.from(scope.querySelectorAll<HTMLElement>(':scope > .md-item')).find((r) => r.querySelector('.md-label')!.textContent === label);
        if (!row) throw new Error(`메뉴 항목 없음: ${path.join(' › ')}`);
        scope = row.querySelector(':scope > .menu-dropdown') ?? scope;
      }
      return row!;
    };
    const openCtx = (el: HTMLElement) => {
      const r = el.getBoundingClientRect();
      el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 2, clientY: r.top + 2 }));
      const menu = document.querySelector<HTMLElement>('.ctx-menu');
      expect(menu, '오른쪽 클릭 메뉴가 열리지 않음');
      expect(menu!.offsetWidth > 100 && menu!.offsetHeight > 100, `메뉴가 화면에 보이지 않음 (${menu!.offsetWidth}×${menu!.offsetHeight})`);
    };
    // Select the text of a table cell, then right-click it.
    const pi = find('새 표');
    select({ start: { pi, off: 0 }, end: { pi, off: doc().lengthOf(doc().para(pi)!) } });
    openCtx(ed.paragraphElement(pi)!);
    expect(ctxItem(['세로 정렬', '위쪽']).classList.contains('checked'), '현재 세로 정렬(위쪽) 표시');
    ctxItem(['세로 정렬', '가운데']).click();
    expect(!document.querySelector('.ctx-menu'), '메뉴가 닫히지 않음');
    eq(doc().cellVAlign(find('새 표')), 'center', '세로 가운데');
    eq(ed.paragraphElement(find('새 표'))!.closest('td')!.style.verticalAlign, 'middle', 'CSS');
    expect(!ed.collapsed(), '선택이 유지되지 않음');
    openCtx(ed.paragraphElement(find('새 표'))!);
    ctxItem(['가로 정렬', '오른쪽 맞춤']).click();
    eq(doc().paraFmt(doc().para(find('새 표'))!).jc, 'right', '가로 오른쪽');
    openCtx(ed.paragraphElement(find('새 표'))!);
    expect(ctxItem(['가로 정렬', '오른쪽 맞춤']).classList.contains('checked'), '현재 가로 정렬 표시');
    ctxItem(['세로 정렬', '아래쪽']).click();
    eq(doc().cellVAlign(find('새 표')), 'bottom', '세로 아래');
    run('edit:undo'); run('edit:undo'); run('edit:undo');
    eq(doc().cellVAlign(find('새 표')), 'top', '되돌리기');
    // Outside a table: horizontal alignment works, vertical alignment is disabled.
    const q = find('문서 끝.');
    caret(q, 0);
    openCtx(ed.paragraphElement(q)!);
    expect(ctxItem(['세로 정렬 (표 안에서)']).classList.contains('disabled'), '표 밖에서 세로 정렬이 활성');
    ctxItem(['가로 정렬', '가운데 맞춤']).click();
    eq(doc().paraFmt(doc().para(q)!).jc, 'center', '표 밖 가로 가운데');
    run('edit:undo');
  });
  await check('표 텍스트 방향: 가로 → 세로 → 아래에서 위로 (메뉴·리본)', async () => {
    const pi = find('새 표');
    caret(pi, 1);
    run('table:textdir-tbRl');
    let td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    eq(getComputedStyle(td).writingMode, 'vertical-rl', '세로 쓰기');
    const pr = ed.paragraphElement(find('새 표'))!.getBoundingClientRect();
    expect(pr.height > pr.width, `글줄이 세로가 아님 ${Math.round(pr.width)}×${Math.round(pr.height)}`);
    eq(doc().cellTextDirection(find('새 표')), 'tbRl', '모델');
    expect(app.cmd.checked('table:textdir-tbRl'), '메뉴 체크 표시');
    await frame();
    expect(document.getElementById('btn-textdir')!.classList.contains('active'), '리본 단추 상태');
    // Typing still goes into the vertical cell.
    caret(find('새 표'), 0);
    type('세로');
    ed.commit();
    expect(texts()[find('새 표')].startsWith('세로'), '세로 셀에 입력');
    run('edit:undo');
    document.getElementById('btn-textdir')!.click(); // the ribbon button cycles to bottom-to-top
    td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    expect(td.classList.contains('w-vert-up'), '아래에서 위로');
    eq(doc().undoLabel(), '텍스트 방향', '실행 취소 항목');
    run('table:textdir'); // back to horizontal
    td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    eq(getComputedStyle(td).writingMode, 'horizontal-tb', '가로로 복귀');
    expect(!doc().cellPrXml(find('새 표')).includes('textDirection'), 'textDirection 제거');
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
  await check('표 바깥 오른쪽 테두리 끌기: 표 넓히기', () => {
    const td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    const tbl = doc().cellOf(find('새 표'))!.tbl;
    const before = doc().columnWidths(tbl);
    const last = td.parentElement!.lastElementChild as HTMLElement;
    const r = last.getBoundingClientRect();
    const y = (r.top + r.bottom) / 2;
    panel.style.pointerEvents = 'none'; // the result panel may cover the table
    // Just outside the table: the pointer is not over a cell.
    const outside = document.elementFromPoint(r.right + 2, y) as HTMLElement;
    outside.dispatchEvent(new MouseEvent('mousedown', { clientX: r.right + 2, clientY: y, bubbles: true, button: 0, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: r.right + 2 + 20 * ed.zoom, clientY: y, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: r.right + 2 + 20 * ed.zoom, clientY: y }));
    const after = doc().columnWidths(doc().cellOf(find('새 표'))!.tbl);
    panel.style.pointerEvents = '';
    eq(Math.round(after[after.length - 1] - before[before.length - 1]), 300, '마지막 칸 +20px(300twip)');
    eq(Math.round(after.reduce((a, b) => a + b, 0) - before.reduce((a, b) => a + b, 0)), 300, '표 너비 +300twip');
  });
  await check('표 모서리 핸들 끌기: 표 전체 크기', () => {
    const td = ed.paragraphElement(find('새 표'))!.closest('td')!;
    const table = td.closest('table')!;
    const tbl = ed.reg.dom.get(table)!;
    const before = doc().columnWidths(tbl);
    const rows0 = Array.from(table.rows).map((r) => r.offsetHeight);
    td.dispatchEvent(new MouseEvent('mousemove', { clientX: td.getBoundingClientRect().left + 3, clientY: td.getBoundingClientRect().top + 3, bubbles: true }));
    const h = ed.host.querySelector<HTMLElement>('.tbl-handle');
    expect(h, '모서리 핸들이 나타나지 않음');
    const hr = h!.getBoundingClientRect(), tr = table.getBoundingClientRect();
    expect(Math.abs(hr.left - tr.right) <= 3 * ed.zoom && Math.abs(hr.top - tr.bottom) <= 3 * ed.zoom, '핸들 위치가 표 오른쪽 아래가 아님');
    const x0 = hr.left + 3, y0 = hr.top + 3;
    const dx = tr.width * 0.5, dy = (rows0.reduce((a, b) => a + b, 0) * ed.zoom) * 0.5;
    h!.dispatchEvent(new MouseEvent('mousedown', { clientX: x0, clientY: y0, bubbles: true, button: 0, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: x0 - dx, clientY: y0 + dy, buttons: 1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: x0 - dx, clientY: y0 + dy }));
    const nt = doc().cellOf(find('새 표'))!.tbl;
    const after = doc().columnWidths(nt);
    const sum = (a: number[]) => a.reduce((x, y) => x + y, 0);
    expect(Math.abs(sum(after) / sum(before) - 0.5) < 0.02, `너비 비율 ${(sum(after) / sum(before)).toFixed(3)}`);
    expect(Math.abs(after[0] / after[1] - before[0] / before[1]) < 0.02, '칸 비율이 유지되지 않음');
    const t2 = ed.paragraphElement(find('새 표'))!.closest('table')!;
    const rows1 = Array.from(t2.rows).map((r) => r.offsetHeight);
    expect(Math.abs(sum(rows1) / sum(rows0) - 1.5) < 0.05, `높이 비율 ${(sum(rows1) / sum(rows0)).toFixed(3)}`);
    eq(doc().undoLabel(), '표 크기', '실행 취소 항목');
    run('edit:undo');
    expect(Math.abs(sum(doc().columnWidths(doc().cellOf(find('새 표'))!.tbl)) - sum(before)) < 2, '실행 취소 후 너비');
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

  // ------------------------------------------------------------------ headers, footers, emoji
  const dblMargin = (page: number, where: 'top' | 'bottom') => {
    const f = ed.frames.children[page - 1] as HTMLElement;
    f.scrollIntoView({ block: 'center' });
    const r = f.getBoundingClientRect();
    const y = where === 'top' ? r.top + 20 * ed.zoom : r.bottom - 20 * ed.zoom;
    f.dispatchEvent(new MouseEvent('dblclick', { clientX: r.left + r.width / 2, clientY: y, bubbles: true }));
  };
  await check('머리글: 여백 두 번 클릭 → 편집 → 입력·Enter·굵게 → 닫기', async () => {
    dblMargin(3, 'top');
    const hf = ed.hf!;
    expect(hf && hf.kind === 'header' && hf.page === 3, '머리글 편집 모드가 아님');
    expect(!document.querySelector<HTMLElement>('.rb-tab[data-tab="hf"]')!.hidden, '머리글 및 바닥글 탭 없음');
    expect(!app.cmd.enabled('format:bullet') && app.cmd.enabled('format:bold'), '명령 제한');
    type(' (수정)');
    document.execCommand('insertParagraph');
    type('둘째 줄');
    const p2 = hf.el.querySelectorAll('p.w-p')[1] as HTMLElement;
    window.getSelection()!.selectAllChildren(p2);
    run('format:bold');
    run('format:align-center');
    run('hf:close');
    expect(!ed.hf, '닫히지 않음');
    const part = doc().headerFooter(3).header!;
    const ps = Array.from(doc().pkg.getXml(part)!.documentElement.getElementsByTagNameNS(W, 'p'));
    eq(ps.map((x) => doc().textOf(x)), ['Oword 기능 점검 — 홀수 쪽 머리글 (수정)', '둘째 줄'], '머리글 내용');
    eq([doc().paraFmt(ps[1]).jc, doc().runFmt(ps[1], doc().items(ps[1])[0].rPr).b], ['center', true], '서식');
    const shown = (ed.frames.children[4] as HTMLElement).querySelector('.page-header')!.textContent; // page 5 is odd too
    expect(shown!.includes('(수정)') && shown!.includes('둘째 줄'), '다른 홀수 페이지에 반영');
    eq([doc().undoLabel()], ['맞춤'], '실행 취소 단계');
    for (let i = 0; i < 3; i++) run('edit:undo'); // 맞춤 → 글자 모양 → 머리글 편집
    expect(!(ed.frames.children[4] as HTMLElement).querySelector('.page-header')!.textContent!.includes('(수정)'), '실행 취소');
  });
  await check('바닥글: 페이지 번호 필드 넣기, Esc로 닫기', () => {
    dblMargin(4, 'bottom');
    expect(ed.hf?.kind === 'footer', '바닥글 모드 아님');
    const p = ed.hf!.el.querySelector('p.w-p')!;
    window.getSelection()!.setBaseAndExtent(p, 0, p, 0);
    type('[');
    run('hf:insert-page');
    type(']');
    ed.hf!.el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    expect(!ed.hf, 'Esc로 닫히지 않음');
    for (const n of [4, 6]) eq((ed.frames.children[n - 1] as HTMLElement).querySelector('.page-footer')!.textContent!.slice(0, 3), `[${n}]`, `${n}쪽 바닥글`);
  });
  await check('머리글/바닥글 옵션: 첫 페이지·홀짝 다르게, 삭제', () => {
    eq(doc().hfOptions(), { titlePg: true, evenOdd: true }, '처음');
    run('hf:odd-even');
    expect(!doc().hfOptions().evenOdd, '홀짝 끄기');
    expect((ed.frames.children[1] as HTMLElement).querySelector('.page-header')!.textContent!.includes('홀수'), '2쪽도 기본 머리글');
    run('hf:first-different');
    expect((ed.frames.children[0] as HTMLElement).querySelector('.page-header'), '1쪽에 기본 머리글');
    run('hf:remove-footer');
    expect(!(ed.frames.children[2] as HTMLElement).querySelector('.page-footer'), '바닥글 삭제');
    run('edit:undo'); run('edit:undo'); run('edit:undo');
    eq(doc().hfOptions(), { titlePg: true, evenOdd: true }, '실행 취소');
  });
  await check('새 문서에 머리글·페이지 번호 만들기 → 저장 → 다시 열기', async () => {
    await open('./blank.docx');
    caret(0, 0);
    type('본문');
    ed.commit();
    run('hf:page-number-footer-right');
    eq((ed.frames.children[0] as HTMLElement).querySelector('.page-footer')?.textContent, '1', '바닥글 페이지 번호');
    dblMargin(1, 'top');
    type('새 머리글');
    run('hf:close');
    eq((ed.frames.children[0] as HTMLElement).querySelector('.page-header')?.textContent, '새 머리글', '머리글');
    const re = await DocxDocument.load(await doc().saveBytes());
    const hf = re.headerFooter(1);
    expect(hf.header && hf.footer, '저장된 머리글/바닥글');
    eq(re.paraFmt(re.pkg.getXml(hf.footer!)!.documentElement.getElementsByTagNameNS(W, 'p')[0]).jc, 'right', '오른쪽 정렬');
    await open('./features.docx');
  });
  await check('이모지 넣기 (삽입 ▸ 이모지, 최근 사용)', () => {
    const pi = find('문서 끝.');
    caret(pi, 0);
    run('insert:emoji');
    const tabs = document.querySelectorAll<HTMLElement>('.emoji-tabs button');
    expect(tabs.length >= 7, '분류 탭');
    Array.from(tabs).find((b) => b.title === '음식')!.click();
    document.querySelector<HTMLElement>('.emoji-grid button')!.click(); // 🍏
    expect(texts()[pi].startsWith('🍏'), texts()[pi]);
    run('insert:emoji');
    eq(document.querySelector<HTMLElement>('.emoji-tabs button')!.title, '최근', '최근 탭');
    document.querySelector<HTMLElement>('.emoji-grid button')!.click();
    expect(texts()[pi].startsWith('🍏🍏'), texts()[pi]);
    const p = ed.paragraphElement(pi)!;
    eq(ed.getRange()?.start.off, 4, '커서 위치 (서로게이트 쌍 2개)');
    expect(getComputedStyle(p.querySelector('span[data-r]')!).fontFamily.includes('Emoji'), '이모지 글꼴');
  });
  await check('머리글 편집 중 이모지·날짜', () => {
    dblMargin(3, 'top');
    run('insert:emoji');
    document.querySelector<HTMLElement>('.emoji-grid button')!.click();
    run('hf:close');
    expect((ed.frames.children[2] as HTMLElement).querySelector('.page-header')!.textContent!.includes('🍏'), '머리글에 이모지');
    run('edit:undo');
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
  await check('보기: 웹 모양 ↔ 인쇄 모양 (선택 유지, 인쇄는 페이지로)', async () => {
    const pi = find('2.1 정렬', true);
    caret(pi, 2);
    run('view:web-layout');
    eq([ed.mode, ed.pages, ed.frames.children.length], ['web', 1, 0], '웹 모양');
    eq(document.getElementById('sb-page')!.textContent, '웹 모양', '상태 표시줄');
    expect(document.getElementById('editor-area')!.classList.contains('no-ruler'), '눈금자 숨김');
    await sleep(120); // a scrollbar may appear and re-wrap the text once
    const sc = document.getElementById('scroll-container')!;
    const want = sc.clientWidth - 4 - 48 * ed.zoom, got = ed.root.getBoundingClientRect().width;
    expect(Math.abs(got - want) < 6, `창 너비에 맞춤 (${Math.round(got)} ≠ ${Math.round(want)})`);
    eq(ed.getRange()?.start, { pi, off: 2 }, '선택 유지');
    expect(!app.cmd.enabled('view:zoom-page'), '웹 모양에서 페이지 맞춤 비활성');
    type('웹');
    ed.commit();
    expect(texts()[pi].startsWith('2.웹1'), '웹 모양에서 입력');
    const real = window.print;
    window.print = () => {};
    run('file:print');
    await sleep(20);
    eq(document.getElementById('print-root')!.querySelectorAll('.print-page').length > 5, true, '웹 모양에서도 페이지로 인쇄');
    window.print = real;
    document.getElementById('print-root')!.innerHTML = '';
    eq(ed.mode, 'web', '인쇄 뒤 웹 모양 유지');
    run('hf:edit-header');
    eq([ed.mode, !!ed.hf], ['print', true], '머리글 편집은 인쇄 모양으로');
    run('hf:close');
    run('view:web-layout');
    run('view:print-layout');
    expect(ed.mode === 'print' && ed.pages > 5 && ed.frames.children.length === ed.pages, '인쇄 모양 복귀');
    run('edit:undo');
  });
  await check('보기 ▸ 확대/축소… 대화상자', () => {
    for (const [v, check] of [['200', (z: number) => z === 2], ['75', (z: number) => z === 0.75], ['page', () => app.fit === 'page'], ['text', () => app.fit === 'text']] as [string, (z: number) => boolean][]) {
      run('view:zoom-dialog');
      const d = document.querySelector('.dlg')!;
      d.querySelector<HTMLInputElement>(`input[value="${v}"]`)!.click();
      Array.from(d.querySelectorAll('button')).find((b) => b.textContent === '확인')!.click();
      expect(check(ed.zoom), `${v}: ${ed.zoom}`);
    }
    run('view:zoom-dialog');
    const d = document.querySelector('.dlg')!;
    const inp = d.querySelector<HTMLInputElement>('[name="pct"]')!;
    inp.value = '133';
    inp.dispatchEvent(new Event('input', { bubbles: true }));
    Array.from(d.querySelectorAll('button')).find((b) => b.textContent === '확인')!.click();
    eq(Math.round(ed.zoom * 100), 133, '사용자 지정');
    run('view:zoom-100');
  });
  await check('보기 ▸ 창 전환 (창 목록, 새 창)', async () => {
    const m = document.querySelector<HTMLElement>('[data-menu="view"] .menu-title')!;
    m.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await sleep(300);
    const items = Array.from(document.querySelectorAll('#window-list .md-item')).map((x) => x.textContent ?? '');
    expect(items.some((t) => t.startsWith('✓ 1.') && t.includes(document.title.replace(/^● /, '').replace(/ — Oword$/, ''))), '현재 창: ' + items.join(' | '));
    expect(items.some((t) => t.startsWith('새 창')), '새 창 항목');
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(app.cmd.has('window:new') && app.cmd.has('window:switch'), '명령');
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
