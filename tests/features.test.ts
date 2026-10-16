// @vitest-environment jsdom
// Exercises the multi-page feature document (scripts/make-features.py): rendering of every
// construct, preservation through edits and saves, and every editing command.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DocxDocument } from '../src/engine/document';
import { commitDom, measure, posToDom } from '../src/engine/edit';
import { ImageCache, newRegistry, renderBlocks, renderHeaderFooter, symbolChar, tableGrid } from '../src/engine/render';
import { descendants } from '../src/engine/xml';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const file = () => new Uint8Array(readFileSync(join(process.cwd(), 'tests/fixtures/features.docx')));
const load = () => DocxDocument.load(file());
const texts = (doc: DocxDocument) => doc.paragraphs().map((p) => doc.textOf(p));
const find = (doc: DocxDocument, s: string) => {
  const i = texts(doc).findIndex((t) => t.includes(s));
  if (i < 0) throw new Error('paragraph not found: ' + s);
  return i;
};
/** Exact paragraph text (headings also appear inside the table of contents). */
const exact = (doc: DocxDocument, s: string) => {
  const i = texts(doc).findIndex((t) => t === s);
  if (i < 0) throw new Error('paragraph not found: ' + s);
  return i;
};
const reload = async (doc: DocxDocument) => DocxDocument.load(await doc.saveBytes());
const count = (doc: DocxDocument, local: string) => doc.body.getElementsByTagNameNS(W, local).length;

function render(doc: DocxDocument) {
  const root = document.createElement('div');
  const reg = newRegistry();
  renderBlocks(doc.body, { doc, images: new ImageCache(doc), part: doc.docPath, reg, editable: true }, root);
  document.body.innerHTML = '';
  document.body.appendChild(root);
  return { root, reg };
}
const paraEl = (root: HTMLElement, pi: number) => root.querySelector<HTMLElement>(`p[data-pi="${pi}"]`)!;
const spanOf = (root: HTMLElement, text: string) => Array.from(root.querySelectorAll<HTMLElement>('span')).find((s) => s.textContent === text)!;

/** Everything the document must keep, counted. */
function inventory(doc: DocxDocument) {
  return {
    tables: count(doc, 'tbl'), drawings: count(doc, 'drawing'), fldChar: count(doc, 'fldChar'), fldSimple: count(doc, 'fldSimple'),
    hyperlinks: count(doc, 'hyperlink'), bookmarks: count(doc, 'bookmarkStart'), ins: count(doc, 'ins'), del: count(doc, 'del'),
    comments: count(doc, 'commentReference'), sdt: count(doc, 'sdt'), footnotes: count(doc, 'footnoteReference'), sym: count(doc, 'sym'),
    sectPr: count(doc, 'sectPr'), math: doc.body.getElementsByTagNameNS('http://schemas.openxmlformats.org/officeDocument/2006/math', 'oMath').length,
  };
}

describe('feature document: reading', () => {
  it('loads every construct', async () => {
    const doc = await load();
    expect(inventory(doc)).toEqual({ tables: 6, drawings: 3, fldChar: 9, fldSimple: 1, hyperlinks: 6, bookmarks: 4, ins: 1, del: 1, comments: 1, sdt: 1, footnotes: 1, sym: 2, sectPr: 2, math: 1 });
    expect(doc.outline().map((o) => o.text)).toEqual(expect.arrayContaining(['1. 글자 모양', '1.4 색과 형광펜', '4.4 여러 쪽에 걸친 긴 표', '9. 긴 본문']));
    expect(doc.outline().find((o) => o.text === '2.1 정렬')!.level).toBe(1);
  });

  it('resolves character formatting', async () => {
    const doc = await load();
    const p = doc.para(find(doc, '굵게, 기울임'))!;
    const fmt = (t: string) => { const it = doc.items(p).find((x) => x.text === t)!; return doc.runFmt(p, it.rPr); };
    expect(fmt('굵은 기울임')).toMatchObject({ b: true, i: true });
    expect(fmt('이중 취소선').dstrike).toBe(true);
    expect(fmt('All Caps').caps).toBe(true);
    expect(fmt('Small Caps').smallCaps).toBe(true);
    expect(fmt('자간 넓게').spacing).toBe(60);
    expect(fmt('보이면 안 됨').vanish).toBe(true);
    const u = doc.para(find(doc, '실선'))!;
    expect(doc.items(u).filter((x) => x.text.trim()).map((x) => doc.runFmt(u, x.rPr).u)).toEqual(['single', 'double', 'dotted', 'dash', 'wave', 'thick', 'words']);
    const c = doc.para(find(doc, '테마 강조 1'))!;
    const col = (t: string) => doc.runFmt(c, doc.items(c).find((x) => x.text.startsWith(t))!.rPr).color;
    expect(col('빨강')).toBe('C00000');
    expect(col('테마 강조 1')).toBe('4F81BD'); // accent1 of the python-docx (Office 2007) theme
    expect(col('테마 어둡게')).toBe('3B618E'); // accent1 with shade BF
    const h = doc.para(find(doc, '노랑 형광펜'))!;
    expect(doc.runFmt(h, doc.items(h)[0].rPr).highlight).toBe('FFFF00');
    expect(doc.runFmt(h, doc.items(h).find((x) => x.text === '글자 음영')!.rPr).shd).toBe('D9E2F3');
  });

  it('resolves paragraph formatting', async () => {
    const doc = await load();
    const f = (s: string) => doc.paraFmt(doc.para(find(doc, s))!);
    expect(f('가운데 정렬 문단').jc).toBe('center');
    expect(f('배분 정렬').jc).toBe('distribute');
    expect(f('왼쪽 2cm, 오른쪽 2cm')).toMatchObject({ left: 1134, right: 1134 });
    expect(f('첫 줄 1cm 들여쓰기')).toMatchObject({ firstLine: 567 });
    expect(f('첫 줄 1cm 내어쓰기')).toMatchObject({ left: 567, hanging: 567 });
    expect(f('문단 위 12pt')).toMatchObject({ before: 240, after: 240 });
    expect(f('줄 간격 2줄')).toMatchObject({ line: 480, lineRule: 'auto' });
    expect(f('줄 간격 고정 20pt')).toMatchObject({ line: 400, lineRule: 'exact' });
    expect(f('줄 간격 최소 30pt')).toMatchObject({ line: 600, lineRule: 'atLeast' });
    expect(f('테두리와 음영이')).toMatchObject({ shd: 'EEF3FC' });
    expect(f('테두리와 음영이').bdr?.bottom?.style).toBe('double');
    expect(doc.paraFmt(doc.para(exact(doc, '2. 문단 모양'))!).pageBreakBefore).toBe(true);
    expect(doc.paraFmt(doc.para(exact(doc, '1.1 기본 모양'))!).keepNext).toBe(true);
  });

  it('numbers multi-level and restarted lists', async () => {
    const doc = await load();
    const lab = (s: string) => doc.labels().get(doc.para(find(doc, s))!)?.text;
    expect(['첫째 장', '첫째 절', '첫째 항', '세부 i', '세부 ii', '둘째 항', '둘째 절', '둘째 장'].map(lab)).toEqual(['1.', '가.', '①', 'i)', 'ii)', '②', '나.', '2.']);
    expect(lab('초성 수준')).toBe('ㄱ.');
    expect(lab('한글 숫자 수준')).toBe('일.');
    expect([lab('다섯째'), lab('여섯째')]).toEqual(['5.', '6.']);
    expect(lab('첫 수준 항목')).toBe('•');
    expect(lab('둘째 수준 항목')).toBeTruthy();
  });

  it('resolves headers and footers per page (first, odd, even)', async () => {
    const doc = await load();
    const text = (part: string | null, n: number) => part ? renderHeaderFooter(doc, new ImageCache(doc), part, n, 9).textContent : null;
    const p1 = doc.headerFooter(1), p2 = doc.headerFooter(2), p3 = doc.headerFooter(3);
    expect(text(p1.header, 1) ?? '').toBe('');
    expect(text(p1.footer, 1)).toBe('표지 바닥글 (첫 쪽 전용)');
    expect(text(p2.header, 2)).toContain('짝수 쪽');
    expect(text(p3.header, 3)).toContain('홀수 쪽');
    expect(text(p2.footer, 2)).toBe('쪽 2 / 9');
    expect(text(p3.footer, 3)).toBe('쪽 3 / 9');
  });
});

describe('feature document: rendering', () => {
  it('renders runs, paragraphs and lists to CSS', async () => {
    const doc = await load();
    const { root } = render(doc);
    expect(spanOf(root, '굵은 기울임').style.fontWeight).toBe('bold');
    expect(spanOf(root, '이중선').style.textDecorationStyle).toBe('double');
    expect(spanOf(root, '물결선').style.textDecorationStyle).toBe('wavy');
    expect(spanOf(root, '보이면 안 됨').style.display).toBe('none');
    expect(spanOf(root, 'All Caps').style.textTransform).toBe('uppercase');
    expect(spanOf(root, '36pt ').style.fontSize).toBe('36pt');
    expect(spanOf(root, '2').style.verticalAlign).toMatch(/super|sub/);
    const dist = paraEl(root, find(doc, '배분 정렬'));
    expect([dist.style.textAlign, dist.style.textAlignLast]).toEqual(['justify', 'justify']);
    const hang = paraEl(root, find(doc, '첫 줄 1cm 내어쓰기'));
    expect([hang.style.marginLeft, hang.style.textIndent]).toEqual(['28.35pt', '-28.35pt']);
    expect(paraEl(root, find(doc, '줄 간격 고정 20pt')).style.lineHeight).toBe('20pt');
    expect(paraEl(root, find(doc, '테두리와 음영이')).style.borderBottom).toMatch(/double/);
    expect(paraEl(root, find(doc, '첫째 항')).dataset.label).toBe('①');
    expect(paraEl(root, exact(doc, '2. 문단 모양')).dataset.breakBefore).toBe('1');
    expect(paraEl(root, find(doc, '구역 1의 마지막')).dataset.breakAfter).toBe('1');
  });

  it('renders merged, nested and borderless tables', async () => {
    const doc = await load();
    const { root } = render(doc);
    const tables = Array.from(root.querySelectorAll<HTMLTableElement>(':scope > table'));
    expect(tables).toHaveLength(5);
    const merged = tables[1];
    const cell = (t: string) => Array.from(merged.querySelectorAll('td')).find((td) => td.textContent === t)!;
    expect(cell('가로 병합 (2칸)').colSpan).toBe(2);
    expect(cell('세로 병합 (3줄)').rowSpan).toBe(3);
    expect([cell('가로·세로 병합 (2×2)').colSpan, cell('가로·세로 병합 (2×2)').rowSpan]).toEqual([2, 2]);
    expect(cell('세로 병합 (3줄)').style.verticalAlign).toBe('middle');
    expect(cell('아래 정렬').style.verticalAlign).toBe('bottom');
    expect(merged.querySelectorAll('td')).toHaveLength(10); // 16 grid cells minus spans and vertical continuations
    expect(tableGrid(doc.body.getElementsByTagNameNS(W, 'tbl')[1]).flat().filter((c) => c.vmerge === 'continue')).toHaveLength(3);
    const nested = tables[2].querySelector('td table') as HTMLTableElement;
    expect(nested.style.width).toBe('100%');
    expect(tables[3].classList.contains('nogrid')).toBe(true);
    expect(tables[3].style.alignSelf).toBe('center');
    expect(tables[0].querySelector('td')!.style.backgroundColor).toBe('rgb(217, 226, 243)');
    expect(tables[4].rows).toHaveLength(41);
  });

  it('renders objects, fields, links, revisions and comments', async () => {
    const doc = await load();
    const { root } = render(doc);
    // jsdom has no blob URLs, so pictures render as placeholders sized to their extent.
    expect(root.querySelectorAll('.w-obj .w-objph').length).toBe(2);
    const anchor = root.querySelector<HTMLElement>('.w-anchor')!;
    expect(anchor.style.float).toBe('right');
    expect(root.querySelector('.w-txbx')!.textContent).toContain('글상자 제목');
    expect(symbolChar('Wingdings', 0xf0fc)).toBe('✓');
    expect(symbolChar('Symbol', 0xf061)).toBe('α');
    expect(root.textContent).toContain('✓ 확인');
    expect(root.querySelector('.w-math sup')!.textContent).toBe('2');
    expect(root.querySelector('.w-noteref')!.textContent).toBe('1');
    const link = root.querySelector<HTMLElement>('a.w-hyperlink[data-href]')!;
    expect(link.dataset.href).toBe('https://github.com/moonlightpark/Open-word');
    expect(root.querySelectorAll('.w-fld[data-fld="PAGE"]').length).toBe(1);
    expect(root.querySelector('.w-del')!.textContent).toBe('삭제된 글자');
    expect(root.querySelector('.w-ins')!.textContent).toBe('삽입된 글자');
    expect(root.querySelector<HTMLElement>('.w-cmt')!.title).toBe('검토자: 메모 내용');
    expect(spanOf(root, '메모가 달린 글자').classList.contains('w-cmt-range')).toBe(true);
    expect(root.querySelector('.w-sdt-inline')!.textContent).toBe('홍길동');
    // The table of contents is one field per entry paragraph run; its first line shows its text.
    expect(paraEl(root, exact(doc, '목차') + 1).textContent).toContain('1. 글자 모양');
  });

  it('maps every offset of a complex paragraph in both directions', async () => {
    const doc = await load();
    const { root } = render(doc);
    for (const s of ['기호:', '외부 링크', '변경 내용 추적', '메모가 달린', '콘텐츠 컨트롤', '이름\t부서', '수식', '각주가 달린', '이 쪽 번호는']) {
      const pi = find(doc, s);
      const p = paraEl(root, pi);
      const len = doc.lengthOf(doc.para(pi)!);
      for (let off = 0; off <= len; off++) {
        const [node, o] = posToDom(p, off);
        expect(measure(p, node, o), `${s} @${off}`).toBe(off);
      }
    }
  });
});

describe('feature document: editing keeps everything else', () => {
  it('typing in every paragraph that has special content preserves it', async () => {
    const doc = await load();
    const before = inventory(doc);
    const { root, reg } = render(doc);
    for (const s of ['기호:', '외부 링크', '변경 내용 추적', '메모가 달린', '콘텐츠 컨트롤', '수식', '각주가 달린', '이 쪽 번호는', '이 문단에는 오른쪽에', '글상자:']) {
      const p = paraEl(root, find(doc, s));
      const first = p.querySelector('span[data-r]')!;
      first.textContent = '★' + first.textContent;
    }
    doc.transact('입력', [], () => commitDom(root, doc, reg));
    expect(inventory(doc)).toEqual(before);
    expect(texts(doc).filter((t) => t.startsWith('★')).length).toBeGreaterThanOrEqual(9);
    const re = await reload(doc);
    expect(inventory(re)).toEqual(before);
    expect(re.labels().get(re.para(find(re, '세부 ii'))!)?.text).toBe('ii)');
  });

  it('typing inside a table cell and a hyperlink keeps structure', async () => {
    const doc = await load();
    const { root, reg } = render(doc);
    const td = Array.from(root.querySelectorAll('td')).find((x) => x.textContent === '가로 병합 (2칸)')!;
    td.querySelector('span')!.textContent = '수정한 병합 셀';
    spanOf(root, 'Oword 저장소').textContent = 'Oword 저장소 (GitHub)';
    doc.transact('입력', [], () => commitDom(root, doc, reg));
    const tbl = doc.body.getElementsByTagNameNS(W, 'tbl')[1];
    expect(tableGrid(tbl).flat().filter((c) => c.vmerge === 'continue')).toHaveLength(3);
    expect(tbl.textContent).toContain('수정한 병합 셀');
    const link = Array.from(doc.body.getElementsByTagNameNS(W, 'hyperlink')).find((h) => h.textContent?.includes('GitHub'))!;
    expect(doc.pkg.rel(doc.docPath, link.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'))?.target).toContain('github.com');
  });

  it('deleting a paragraph range in the DOM drops exactly those paragraphs', async () => {
    const doc = await load();
    const n = doc.paragraphs().length;
    const { root, reg } = render(doc);
    paraEl(root, find(doc, '왼쪽 정렬 문단')).remove();
    paraEl(root, find(doc, '오른쪽 정렬 문단')).remove();
    doc.transact('입력', [], () => commitDom(root, doc, reg));
    expect(doc.paragraphs().length).toBe(n - 2);
    expect(texts(doc).some((t) => t.includes('왼쪽 정렬 문단'))).toBe(false);
    doc.undo();
    expect(doc.paragraphs().length).toBe(n);
  });

  it('applies formatting across paragraphs with runs, links and fields', async () => {
    const doc = await load();
    const before = inventory(doc);
    const a = find(doc, '외부 링크'), b = find(doc, '이 쪽 번호는');
    doc.applyRunProps({ start: { pi: a, off: 2 }, end: { pi: b, off: 4 } }, { b: true, color: '2B6FD6', sz: 13 });
    const p = doc.para(a)!;
    const linkText = doc.items(p).find((x) => x.text === 'Oword 저장소')!;
    expect(doc.runFmt(p, linkText.rPr)).toMatchObject({ b: true, color: '2B6FD6', sz: 26 });
    expect(doc.runFmt(p, doc.items(p)[0].rPr).b).toBeFalsy(); // "외부" before the range
    expect(inventory(doc)).toEqual(before);
    doc.clearRunProps({ start: { pi: a, off: 0 }, end: { pi: a, off: doc.lengthOf(p) } });
    expect(doc.runFmt(doc.para(a)!, doc.items(doc.para(a)!).find((x) => x.text.includes('Oword'))!.rPr).color).toBeUndefined();
    expect(inventory(doc)).toEqual(before);
  });

  it('runs every paragraph command and survives a save', async () => {
    const doc = await load();
    const pis = [find(doc, '왼쪽 정렬 문단'), find(doc, '가운데 정렬 문단')];
    doc.applyParaProps(pis, { jc: 'right' });
    doc.applyParaProps(pis, { line: 360, lineRule: 'auto', before: 120, after: 0 });
    doc.indent(pis, 1);
    doc.indent(pis, 1);
    doc.indent(pis, -1);
    doc.setStyle(pis, 'Heading3');
    doc.toggleList(pis, 'number');
    for (const pi of pis) expect(doc.paraFmt(doc.para(pi)!)).toMatchObject({ jc: 'right', line: 360, before: 120, after: 0 });
    expect(doc.pStyleOf(doc.para(pis[0])!)).toBe('Heading3');
    expect(doc.labels().get(doc.para(pis[1])!)).toBeTruthy();
    const re = await reload(doc);
    expect(re.pStyleOf(re.para(pis[0])!)).toBe('Heading3');
    expect(re.outline().some((o) => o.text === '왼쪽 정렬 문단입니다.')).toBe(true);
  });

  it('table commands on merged cells and column/row sizing', async () => {
    const doc = await load();
    const pi = find(doc, '매출');
    const tbl = doc.cellOf(pi)!.tbl;
    doc.tableOp(pi, 'row-above');
    doc.tableOp(find(doc, '매출'), 'col-right');
    expect(descendants(tbl, 'w:tr')).toHaveLength(5);
    expect(descendants(tbl, 'w:gridCol')).toHaveLength(5);
    doc.setColumnWidths(tbl, [3000, 1000, 1500, 1500, 2000]);
    expect(doc.columnWidths(tbl)).toEqual([3000, 1000, 1500, 1500, 2000]);
    expect(descendants(tbl, 'w:tcW').slice(0, 5).map((w) => w.getAttributeNS(W, 'w'))).toEqual(['3000', '1000', '1500', '1500', '2000']);
    expect(tbl.getElementsByTagNameNS(W, 'tblLayout')[0].getAttributeNS(W, 'type')).toBe('fixed');
    doc.setRowHeight(tbl, 2, 800);
    const trH = descendants(tbl, 'w:tr')[2].getElementsByTagNameNS(W, 'trHeight')[0];
    expect([trH.getAttributeNS(W, 'val'), trH.getAttributeNS(W, 'hRule')]).toEqual(['800', 'atLeast']);
    // Whole-table resize (corner handle): widths and row heights change in one undo step.
    const undoDepth = doc.undoLabel();
    doc.resizeTable(tbl, [1500, 500, 750, 750, 1000], descendants(tbl, 'w:tr').map((_, i) => (i === 0 ? 600 : null)));
    expect(doc.columnWidths(tbl)).toEqual([1500, 500, 750, 750, 1000]);
    expect(tbl.getElementsByTagNameNS(W, 'tblW')[0].getAttributeNS(W, 'w')).toBe('4500');
    expect(descendants(tbl, 'w:tr')[0].getElementsByTagNameNS(W, 'trHeight')[0].getAttributeNS(W, 'val')).toBe('600');
    expect(doc.undoLabel()).toBe('표 크기');
    doc.undo();
    expect(doc.undoLabel()).toBe(undoDepth);
    expect(doc.columnWidths(doc.cellOf(find(doc, '매출'))!.tbl)).toEqual([3000, 1000, 1500, 1500, 2000]);
    // Merged table: deleting a column that a gridSpan crosses shrinks the span.
    const m = find(doc, '가로 병합');
    const mt = doc.cellOf(m)!.tbl;
    doc.tableOp(m, 'del-col');
    expect(descendants(mt, 'w:gridCol')).toHaveLength(3);
    const re = await reload(doc);
    expect(re.body.getElementsByTagNameNS(W, 'tbl')[0].getElementsByTagNameNS(W, 'gridCol')).toHaveLength(5);
    // pPr/tcPr/trPr children stay in schema order after all of the above.
    const xml = re.xmlText();
    expect(xml).not.toMatch(/<w:tcPr>(?:(?!<\/w:tcPr>).)*<w:vMerge[^>]*\/>(?:(?!<\/w:tcPr>).)*<w:tcW /);
    expect(xml).not.toMatch(/<w:tblPr>(?:(?!<\/w:tblPr>).)*<w:tblLayout[^>]*\/>(?:(?!<\/w:tblPr>).)*<w:tblW /);
  });

  it('find/replace skips objects and replaces across runs', async () => {
    const doc = await load();
    const n = doc.findAll('편집기', false).length;
    expect(n).toBeGreaterThan(20);
    expect(doc.replaceAll('편집기', 'EDITOR', true)).toBe(n);
    expect(doc.findAll('편집기', false)).toHaveLength(0);
    expect(doc.findAll('editor', false)).toHaveLength(n);
    // "굵게, 기울" spans two runs.
    expect(doc.replaceAll('굵게, 기울', '[교체]', true)).toBe(1);
    expect(doc.textOf(doc.para(find(doc, '[교체]'))!).startsWith('[교체]임')).toBe(true);
  });

  it('undo/redo walks back through a long session', async () => {
    const doc = await load();
    const snap = doc.xmlText();
    const pi = find(doc, '첫 수준 항목');
    doc.insertText({ pi, off: 0 }, '앞 ');
    doc.splitAt({ pi, off: 2 });
    doc.applyRunProps({ start: { pi, off: 0 }, end: { pi, off: 1 } }, { i: true });
    doc.toggleList([pi], 'bullet');
    doc.insertTable({ pi, off: 0 }, 2, 2);
    doc.insertPageBreak({ pi: pi + 3, off: 0 });
    doc.setPageSetup({ ...doc.pageSetup(), orient: 'landscape', w: 16838, h: 11906 });
    let k = 0;
    while (doc.canUndo()) { doc.undo(); k++; }
    expect(k).toBe(7);
    expect(doc.xmlText()).toBe(snap);
    while (doc.canRedo()) doc.redo();
    expect(doc.pageSetup().orient).toBe('landscape');
  });

  it('keeps footnotes, comments and settings parts byte-identical through a save', async () => {
    const doc = await load();
    doc.insertText({ pi: 0, off: 0 }, '수정 ');
    const re = await reload(doc);
    for (const part of ['word/footnotes.xml', 'word/comments.xml', 'word/settings.xml', 'word/numbering.xml', 'word/styles.xml']) {
      const a = new XMLSerializer().serializeToString(doc.pkg.getXml(part)!);
      const b = new XMLSerializer().serializeToString(re.pkg.getXml(part)!);
      expect(b, part).toBe(a);
    }
    expect(re.pkg.partNames().sort()).toEqual(doc.pkg.partNames().sort());
  });
});

describe('table cell alignment', () => {
  it('sets top / center / bottom on the selected cells and renders it', async () => {
    const doc = await load();
    const a = find(doc, '매출'), b = find(doc, '비용');
    expect(doc.cellVAlign(a)).toBe('top');
    expect(doc.cellVAlign(exact(doc, '1. 글자 모양'))).toBeNull();
    doc.setCellVAlign([a, a + 1, b], 'center');
    expect([doc.cellVAlign(a), doc.cellVAlign(a + 1), doc.cellVAlign(b)]).toEqual(['center', 'center', 'center']);
    doc.setCellVAlign([a], 'bottom');
    expect(doc.cellVAlign(a)).toBe('bottom');
    const { root } = render(doc);
    const td = Array.from(root.querySelectorAll('td')).find((x) => x.textContent === '매출')!;
    expect(td.style.verticalAlign).toBe('bottom');
    doc.setCellVAlign([a], 'top');
    expect(doc.cellPrXml(a)).not.toContain('vAlign');
    doc.undo();
    expect(doc.cellVAlign(a)).toBe('bottom');
    const re = await reload(doc);
    expect(re.cellVAlign(find(re, '비용'))).toBe('center');
  });
});

describe('Word-style menu commands', () => {
  it('changes case without touching formatting', async () => {
    const doc = await load();
    const pi = find(doc, 'All Caps');
    const p = doc.para(pi)!;
    const t = doc.textOf(p);
    const s = t.indexOf('All Caps'), e = t.indexOf('Small Caps') + 'Small Caps'.length;
    doc.changeCase({ start: { pi, off: s }, end: { pi, off: e } }, 'upper');
    expect(doc.textOf(doc.para(pi)!).slice(s, e)).toBe('ALL CAPS, SMALL CAPS');
    doc.changeCase({ start: { pi, off: s }, end: { pi, off: e } }, 'lower');
    expect(doc.textOf(doc.para(pi)!).slice(s, e)).toBe('all caps, small caps');
    doc.changeCase({ start: { pi, off: s }, end: { pi, off: e } }, 'title');
    expect(doc.textOf(doc.para(pi)!).slice(s, e)).toBe('All Caps, Small Caps');
    const q = doc.para(pi)!;
    expect(doc.runFmt(q, doc.items(q).find((x) => x.text === 'Small Caps')!.rPr).smallCaps).toBe(true);
  });

  it('inserts hyperlinks (selection, empty caret, bookmark) that survive a save', async () => {
    const doc = await load();
    const pi = find(doc, '왼쪽 정렬 문단');
    const before = doc.body.getElementsByTagNameNS(W, 'hyperlink').length;
    doc.insertHyperlink({ start: { pi, off: 0 }, end: { pi, off: 2 } }, 'https://example.com');
    const end = doc.insertHyperlink({ start: { pi, off: 5 }, end: { pi, off: 5 } }, 'https://oword.dev', '링크 글');
    doc.insertHyperlink({ start: { pi, off: end.off }, end: { pi, off: end.off } }, '#_Toc3', '표로');
    expect(doc.textOf(doc.para(pi)!).startsWith('왼쪽 정렬링크 글표로')).toBe(true);
    const re = await reload(doc);
    const links = Array.from(re.body.getElementsByTagNameNS(W, 'hyperlink'));
    expect(links.length).toBe(before + 3);
    const rid = (h: Element) => h.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const targets = links.map((h) => rid(h) ? re.pkg.rel(re.docPath, rid(h))!.target : '#' + h.getAttributeNS(W, 'anchor'));
    expect(targets).toEqual(expect.arrayContaining(['https://example.com', 'https://oword.dev', '#_Toc3']));
    const relsXml = new XMLSerializer().serializeToString(re.pkg.getXml('word/_rels/document.xml.rels')!);
    expect(relsXml).toMatch(/Target="https:\/\/example.com" TargetMode="External"/);
    const { root } = render(re);
    expect(Array.from(root.querySelectorAll<HTMLElement>('a.w-hyperlink')).some((a) => a.dataset.href === 'https://example.com' && a.textContent === '왼쪽')).toBe(true);
  });

  it('inserts a line break', async () => {
    const doc = await load();
    const pi = find(doc, '왼쪽 정렬 문단');
    doc.insertLineBreak({ pi, off: 2 });
    expect(doc.textOf(doc.para(pi)!).startsWith('왼쪽\n 정렬')).toBe(true);
    expect(doc.paragraphs().length).toBe((await load()).paragraphs().length);
  });

  it('accepts or rejects all tracked changes', async () => {
    const a = await load();
    expect(a.revisionCount()).toBe(2);
    a.resolveRevisions(true);
    expect(a.revisionCount()).toBe(0);
    expect(a.textOf(a.para(find(a, '변경 내용 추적'))!)).toBe('변경 내용 추적: 삽입된 글자 /  끝.');
    const r = await load();
    r.resolveRevisions(false);
    expect(r.textOf(r.para(find(r, '변경 내용 추적'))!)).toBe('변경 내용 추적:  / 삭제된 글자 끝.');
    r.undo();
    expect(r.revisionCount()).toBe(2);
  });

  it('aligns cells on both axes (3×3 grid) and sets separate Korean/Latin fonts', async () => {
    const doc = await load();
    const a = find(doc, '매출');
    doc.alignCells([a], 'bottom', 'right');
    expect(doc.cellVAlign(a)).toBe('bottom');
    expect(doc.paraFmt(doc.para(a)!).jc).toBe('right');
    expect(doc.cellParagraphs([a])).toEqual([a]);
    const pi = find(doc, '왼쪽 정렬 문단');
    doc.applyRunProps({ start: { pi, off: 0 }, end: { pi, off: 4 } }, { ascii: 'Georgia', ea: '바탕', dstrike: true, smallCaps: true, vanish: false });
    const f = doc.fmtAt({ pi, off: 1 })!;
    expect([f.ascii, f.ea, f.dstrike, f.smallCaps, f.strike]).toEqual(['Georgia', '바탕', true, true, undefined]);
  });
});

describe('headers and footers', () => {
  it('creates a header and footer in a document that has none, and undo removes them', async () => {
    const doc = await DocxDocument.load(new Uint8Array(readFileSync(join(process.cwd(), 'public/blank.docx'))));
    expect(doc.headerFooter(1)).toEqual({ header: null, footer: null });
    const parts = doc.pkg.partNames().length;
    const h = doc.openHeaderFooter('header', 1);
    expect(doc.headerFooter(1).header).toBe(h);
    expect(doc.pkg.getXml(h)!.documentElement.localName).toBe('hdr');
    doc.insertPageNumber('footer', 'center', true);
    const f = doc.headerFooter(1).footer!;
    expect(renderHeaderFooter(doc, new ImageCache(doc), f, 3, 7).textContent).toBe('3 / 7');
    const re = await reload(doc);
    expect(re.headerFooter(2)).toEqual({ header: h, footer: f });
    const ct = new XMLSerializer().serializeToString(re.pkg.getXml('[Content_Types].xml')!);
    expect(ct).toContain(`PartName="/${h}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"`);
    const sect = new XMLSerializer().serializeToString(re.body.getElementsByTagNameNS(W, 'sectPr')[0]);
    expect(sect.indexOf('headerReference')).toBeLessThan(sect.indexOf('footerReference'));
    expect(sect.indexOf('footerReference')).toBeLessThan(sect.indexOf('pgSz'));
    doc.undo();
    doc.undo();
    expect(doc.headerFooter(1)).toEqual({ header: null, footer: null });
    expect(doc.pkg.partNames().length).toBe(parts);
  });

  it('edits header text through the DOM and keeps fields', async () => {
    const doc = await load();
    const part = doc.headerFooter(3).header!;
    const host = document.createElement('div');
    const reg = newRegistry();
    renderBlocks(doc.pkg.getXml(part)!.documentElement, { doc, images: new ImageCache(doc), part, reg, editable: true, page: 3, pages: 9 }, host);
    const span = host.querySelector('span[data-r]')!;
    span.textContent = '바뀐 머리글 ' + span.textContent;
    host.insertAdjacentHTML('beforeend', '<p class="w-p">둘째 줄</p>');
    const { commitContainer } = await import('../src/engine/edit');
    doc.transact('머리글 편집', [part], () => commitContainer(host, doc.pkg.getXml(part)!.documentElement, doc, reg));
    const text = renderHeaderFooter(doc, new ImageCache(doc), part, 3, 9).textContent;
    expect(text).toContain('바뀐 머리글 Oword 기능 점검');
    expect(text).toContain('둘째 줄');
    // The new paragraph inherits the header paragraph's alignment.
    const ps = doc.pkg.getXml(part)!.documentElement.getElementsByTagNameNS(W, 'p');
    expect(doc.paraFmt(ps[1]).jc).toBe('right');
    doc.formatSegments(part, [{ p: ps[1], s: 0, e: -1 }], { run: { b: true }, jc: 'center' });
    expect(doc.paraFmt(ps[1]).jc).toBe('center');
    expect(doc.runFmt(ps[1], doc.items(ps[1])[0].rPr).b).toBe(true);
    const re = await reload(doc);
    expect(renderHeaderFooter(re, new ImageCache(re), part, 3, 9).textContent).toContain('둘째 줄');
    // Footer fields still work after an edit elsewhere.
    expect(renderHeaderFooter(re, new ImageCache(re), re.headerFooter(4).footer!, 4, 9).textContent).toBe('쪽 4 / 9');
  });

  it('toggles first-page / odd-even options and removes a header', async () => {
    const doc = await load();
    expect(doc.hfOptions()).toEqual({ titlePg: true, evenOdd: true });
    doc.setHFOptions({ titlePg: false, evenOdd: false });
    expect(doc.hfOptions()).toEqual({ titlePg: false, evenOdd: false });
    expect([doc.hfType(1), doc.hfType(2)]).toEqual(['default', 'default']);
    doc.setHFOptions({ evenOdd: true });
    expect(doc.hfType(2)).toBe('even');
    const settings = new XMLSerializer().serializeToString(doc.settings!);
    expect(settings).toContain('evenAndOddHeaders');
    doc.removeHeaderFooter('header', 'even');
    expect(doc.headerFooter(2).header).toBeNull();
    expect(doc.headerFooter(3).header).not.toBeNull();
    const re = await reload(doc);
    expect(re.headerFooter(2).header).toBeNull();
    doc.undo();
    expect(doc.headerFooter(2).header).not.toBeNull();
  });

  it('inserts emoji (surrogate pairs) and keeps offsets consistent', async () => {
    const doc = await load();
    const pi = find(doc, '왼쪽 정렬 문단');
    const end = doc.insertText({ pi, off: 2 }, '😀👍🏽');
    expect(end.off).toBe(2 + '😀👍🏽'.length);
    const { root } = render(doc);
    const p = paraEl(root, pi);
    for (let off = 0; off <= doc.lengthOf(doc.para(pi)!); off++) {
      const [n, o] = posToDom(p, off);
      expect(measure(p, n, o)).toBe(off);
    }
    expect((await reload(doc)).textOf((await reload(doc)).para(pi)!).startsWith('왼쪽😀👍🏽')).toBe(true);
  });
});
