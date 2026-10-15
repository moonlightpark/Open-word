// @vitest-environment jsdom
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DocxDocument } from '../src/engine/document';
import { commitDom, domToPos, measure, posToDom } from '../src/engine/edit';
import { formatNumber } from '../src/engine/numbering';
import { ImageCache, newRegistry, renderBlocks } from '../src/engine/render';
import { descendants } from '../src/engine/xml';

const dir = join(process.cwd(), 'tests');
const fixture = () => new Uint8Array(readFileSync(join(dir, 'fixtures/sample.docx')));
const blank = () => new Uint8Array(readFileSync(join(process.cwd(), 'public/blank.docx')));
const texts = (doc: DocxDocument) => doc.paragraphs().map((p) => doc.textOf(p));
const find = (doc: DocxDocument, s: string) => texts(doc).findIndex((t) => t.includes(s));
const reload = async (doc: DocxDocument) => DocxDocument.load(await doc.saveBytes());

function render(doc: DocxDocument) {
  const root = document.createElement('div');
  const reg = newRegistry();
  renderBlocks(doc.body, { doc, images: new ImageCache(doc), part: doc.docPath, reg, editable: true }, root);
  document.body.appendChild(root);
  return { root, reg };
}

describe('DocxDocument', () => {
  it('loads paragraphs, styles, page setup and headers', async () => {
    const doc = await DocxDocument.load(fixture());
    expect(texts(doc)[0]).toBe('Oword 샘플 문서');
    const ps = doc.pageSetup();
    expect(Math.round(ps.w / 566.929)).toBe(21);
    expect(Math.round(ps.h / 566.929 * 10) / 10).toBe(29.7);
    expect(doc.headerFooter(1).header).toMatch(/^word\/header\d*\.xml$/);
    expect(doc.paraFmt(doc.para(0)!).outlineLvl).toBeUndefined();
    const h1 = find(doc, '1. 글자 모양');
    expect(doc.paraFmt(doc.para(h1)!).outlineLvl).toBe(0);
    expect(doc.outline().map((o) => o.text)).toContain('3. 표');
  });

  it('resolves run formatting through styles', async () => {
    const doc = await DocxDocument.load(fixture());
    const pi = find(doc, '이 문서는');
    const p = doc.para(pi)!;
    const bold = doc.items(p).find((it) => it.text === 'Oword')!;
    expect(doc.runFmt(p, bold.rPr).b).toBe(true);
    const red = doc.items(p).find((it) => it.text === '렌더링과 편집')!;
    expect(doc.runFmt(p, red.rPr)).toMatchObject({ i: true, color: 'C00000', sz: 20, ea: '맑은 고딕' });
  });

  it('computes list labels', async () => {
    const doc = await DocxDocument.load(fixture());
    const labels = doc.labels();
    const b = doc.para(find(doc, '첫째 글머리표'))!;
    const n2 = doc.para(find(doc, '번호 매긴 둘째'))!;
    expect(labels.get(b)?.text).toBe('•');
    expect(labels.get(n2)?.text).toBe('2.');
    expect(formatNumber(3, 'ganada')).toBe('다');
    expect(formatNumber(4, 'upperRoman')).toBe('IV');
    expect(formatNumber(2, 'decimalEnclosedCircle')).toBe('②');
  });

  it('applies character formatting to part of a run, with undo', async () => {
    const doc = await DocxDocument.load(fixture());
    const pi = find(doc, '이 문서는');
    doc.applyRunProps({ start: { pi, off: 0 }, end: { pi, off: 2 } }, { u: true, sz: 14, font: '바탕' });
    const p = doc.para(pi)!;
    const first = doc.items(p)[0];
    expect(first.text).toBe('이 ');
    expect(doc.runFmt(p, first.rPr)).toMatchObject({ u: 'single', sz: 28, ascii: '바탕', ea: '바탕' });
    expect(doc.textOf(p).startsWith('이 문서는 Oword')).toBe(true);
    // Run properties keep the schema order (rFonts … sz … u).
    const xml = new XMLSerializer().serializeToString(first.rPr!);
    expect(xml.indexOf('rFonts')).toBeLessThan(xml.indexOf('w:sz'));
    expect(xml.indexOf('w:sz')).toBeLessThan(xml.indexOf('w:u '));
    doc.undo();
    expect(doc.items(doc.para(pi)!)[0].text).toBe('이 문서는 ');
  });

  it('splits paragraphs, ends lists on empty items and uses the next style after headings', async () => {
    const doc = await DocxDocument.load(fixture());
    const n = doc.paragraphs().length;
    const h = find(doc, '1. 글자 모양');
    const pos = doc.splitAt({ pi: h, off: doc.lengthOf(doc.para(h)!) });
    expect(pos).toEqual({ pi: h + 1, off: 0 });
    expect(doc.paragraphs().length).toBe(n + 1);
    expect(doc.pStyleOf(doc.para(h + 1)!)).toBeNull(); // Heading 1 → Normal
    const li = find(doc, '셋째 글머리표');
    const after = doc.splitAt({ pi: li, off: doc.lengthOf(doc.para(li)!) });
    expect(doc.labels().get(doc.para(after.pi)!)?.text).toBe('•');
    doc.splitAt(after); // Enter on an empty item leaves the list
    expect(doc.labels().get(doc.para(after.pi)!)).toBeUndefined();
    const mid = find(doc, '이 문서는');
    doc.splitAt({ pi: mid, off: 3 });
    expect(doc.textOf(doc.para(mid)!)).toBe('이 문');
    expect(doc.textOf(doc.para(mid + 1)!).startsWith('서는 Oword')).toBe(true);
  });

  it('toggles bullets and numbering, creating numbering definitions when needed', async () => {
    const doc = await DocxDocument.load(blank());
    doc.insertText({ pi: 0, off: 0 }, '하나\n둘\n셋');
    expect(texts(doc)).toEqual(['하나', '둘', '셋']);
    doc.toggleList([0, 1, 2], 'number');
    expect([0, 1, 2].map((i) => doc.labels().get(doc.para(i)!)?.text)).toEqual(['1.', '2.', '3.']);
    doc.indent([1], 1);
    expect(doc.labels().get(doc.para(1)!)?.text).toBe('가.');
    doc.toggleList([0, 1, 2], 'number');
    expect(doc.labels().size).toBe(0);
    doc.toggleList([0], 'bullet');
    expect(doc.labels().get(doc.para(0)!)?.text).toBe('●');
    const re = await reload(doc);
    expect(re.labels().get(re.para(0)!)?.text).toBe('●');
  });

  it('inserts and edits tables', async () => {
    const doc = await DocxDocument.load(blank());
    doc.insertText({ pi: 0, off: 0 }, '앞');
    const pos = doc.insertTable({ pi: 0, off: 1 }, 2, 3);
    const cell = doc.cellOf(pos.pi)!;
    expect(cell).not.toBeNull();
    expect(descendants(cell.tbl, 'w:tc')).toHaveLength(6);
    doc.tableOp(pos.pi, 'row-below');
    doc.tableOp(pos.pi, 'col-right');
    expect(descendants(cell.tbl, 'w:tr')).toHaveLength(3);
    expect(descendants(cell.tbl, 'w:gridCol')).toHaveLength(4);
    expect(descendants(cell.tbl, 'w:tc')).toHaveLength(12);
    doc.tableOp(doc.indexOf(descendants(cell.tbl, 'w:p')[0]), 'del-col');
    expect(descendants(cell.tbl, 'w:tc')).toHaveLength(9);
    const after = doc.tableOp(doc.indexOf(descendants(cell.tbl, 'w:p')[0]), 'del-table')!;
    expect(descendants(doc.body, 'w:tbl')).toHaveLength(0);
    expect(after.pi).toBeGreaterThanOrEqual(0);
  });

  it('inserts images, page breaks, tabs; finds and replaces', async () => {
    const doc = await DocxDocument.load(blank());
    doc.insertText({ pi: 0, off: 0 }, '사과와 사과');
    doc.insertImage({ pi: 0, off: 2 }, new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), 'image/png', 100, 50);
    expect(doc.textOf(doc.para(0)!)).toBe('사과￼와 사과');
    doc.insertTab({ pi: 0, off: 0 });
    doc.insertPageBreak({ pi: 0, off: doc.lengthOf(doc.para(0)!) });
    expect(doc.paragraphs()).toHaveLength(2);
    expect(doc.findAll('사과', false)).toHaveLength(2);
    expect(doc.replaceAll('사과', '배', false)).toBe(2);
    expect(doc.textOf(doc.para(0)!)).toBe('\t배￼와 배\n');
    const re = await reload(doc);
    expect(re.pkg.rels(re.docPath).some((r) => r.type.endsWith('/image') && r.target.startsWith('word/media/'))).toBe(true);
    expect(re.xmlText()).toContain('w:type="page"');
  });

  it('changes page setup and paragraph properties', async () => {
    const doc = await DocxDocument.load(fixture());
    const ps = doc.pageSetup();
    doc.setPageSetup({ ...ps, w: ps.h, h: ps.w, orient: 'landscape', left: 567 });
    expect(doc.pageSetup()).toMatchObject({ w: ps.h, orient: 'landscape', left: 567 });
    const pi = find(doc, '이 문서는');
    doc.applyParaProps([pi], { jc: 'center', left: 720, firstLine: 360, before: 240, line: 360, lineRule: 'auto' });
    expect(doc.paraFmt(doc.para(pi)!)).toMatchObject({ jc: 'center', left: 720, firstLine: 360, before: 240, line: 360 });
    doc.setStyle([pi], 'Heading2');
    expect(doc.pStyleOf(doc.para(pi)!)).toBe('Heading2');
    const xml = new XMLSerializer().serializeToString(doc.para(pi)!.firstElementChild!);
    // pPr children in schema order: pStyle, spacing, ind, jc.
    expect(xml.indexOf('pStyle')).toBeLessThan(xml.indexOf('spacing'));
    expect(xml.indexOf('spacing')).toBeLessThan(xml.indexOf('w:ind'));
    expect(xml.indexOf('w:ind')).toBeLessThan(xml.indexOf('w:jc'));
  });

  it('copies a range and pastes it as paragraphs', async () => {
    const doc = await DocxDocument.load(fixture());
    const a = find(doc, '첫째 글머리표'), b = find(doc, '둘째 글머리표');
    const clip = doc.extractRange({ start: { pi: a, off: 3 }, end: { pi: b, off: 2 } });
    expect(clip).toHaveLength(2);
    const end = doc.paragraphs().length - 1;
    const len = doc.lengthOf(doc.para(end)!);
    const at = doc.insertFragment({ pi: end, off: len }, clip);
    expect(doc.textOf(doc.para(end)!).endsWith('글머리표 항목')).toBe(true);
    expect(doc.textOf(doc.para(at.pi)!)).toBe('둘째');
  });
});

describe('renderer and DOM bridge', () => {
  it('renders paragraphs, lists, tables, images and fields', async () => {
    const doc = await DocxDocument.load(fixture());
    const { root } = render(doc);
    expect(root.querySelectorAll('p.w-p[data-pi]').length).toBe(doc.paragraphs().length);
    expect(root.querySelector('table.w-tbl')).not.toBeNull();
    expect(root.querySelectorAll('td')).toHaveLength(9);
    const red = Array.from(root.querySelectorAll('span')).find((s) => s.textContent === '렌더링과 편집') as HTMLElement;
    expect(red.style.color).toBe('rgb(192, 0, 0)');
    expect(red.style.fontStyle).toBe('italic');
    expect(root.querySelector('p[data-label="•"]')).not.toBeNull();
    expect(root.querySelector('.w-obj')).not.toBeNull();
    expect(root.querySelector('.w-pgbr')).not.toBeNull();
    expect(root.querySelector('.w-tab')).not.toBeNull();
  });

  it('maps DOM positions to model offsets and back', async () => {
    const doc = await DocxDocument.load(fixture());
    const { root } = render(doc);
    const pi = find(doc, '이 문서는');
    const p = root.querySelector<HTMLElement>(`p[data-pi="${pi}"]`)!;
    const t = doc.textOf(doc.para(pi)!);
    for (const off of [0, 3, 6, 10, 15, t.length]) {
      const [node, o] = posToDom(p, off);
      expect(measure(p, node, o)).toBe(off);
      expect(domToPos(root, node, o)).toEqual({ pi, off });
    }
    const tabP = root.querySelector<HTMLElement>(`p[data-pi="${find(doc, '탭으로')}"]`)!;
    const [n2, o2] = posToDom(tabP, 4);
    expect(measure(tabP, n2, o2)).toBe(4);
  });

  it('commits typed text back to XML while keeping formatting and untouched paragraphs', async () => {
    const doc = await DocxDocument.load(fixture());
    const { root, reg } = render(doc);
    const pi = find(doc, '이 문서는');
    const untouched = doc.para(find(doc, '1. 글자 모양'))!;
    const p = root.querySelector<HTMLElement>(`p[data-pi="${pi}"]`)!;
    const bold = Array.from(p.querySelectorAll('span')).find((s) => s.textContent === 'Oword')!;
    bold.textContent = 'Oword 편집기';
    p.insertAdjacentHTML('afterend', '<p class="w-p">새 문단 &amp; 끝</p>');
    // Delete the picture paragraph's image.
    root.querySelector('.w-obj')!.remove();
    doc.transact('입력', [], () => commitDom(root, doc, reg));
    const q = doc.para(pi)!;
    const it = doc.items(q).find((x) => x.text === 'Oword 편집기')!;
    expect(doc.runFmt(q, it.rPr).b).toBe(true);
    expect(doc.textOf(doc.para(pi + 1)!)).toBe('새 문단 & 끝');
    expect(doc.paragraphs()).toContain(untouched);
    expect(descendants(doc.body, 'w:drawing')).toHaveLength(0);
    // The footer field and section properties survive a round trip.
    const re = await reload(doc);
    expect(re.xmlText()).toContain('<w:sectPr');
    expect(texts(re)).toContain('새 문단 & 끝');
    doc.undo();
    expect(descendants(doc.body, 'w:drawing')).toHaveLength(1);
  });

  it('treats a <br> followed only by empty elements as the line placeholder', async () => {
    const doc = await DocxDocument.load(fixture());
    const { root, reg } = render(doc);
    const pi = find(doc, '이 문서는');
    const p = root.querySelector<HTMLElement>(`p[data-pi="${pi}"]`)!;
    p.innerHTML = '<br><span data-r="0"></span>';
    expect(measure(p, p, 2)).toBe(0);
    doc.transact('입력', [], () => commitDom(root, doc, reg));
    expect(doc.textOf(doc.para(pi)!)).toBe('');
  });

  it('writes edited documents for external validation', async () => {
    const doc = await DocxDocument.load(fixture());
    doc.applyRunProps({ start: { pi: 0, off: 0 }, end: { pi: 0, off: 5 } }, { color: '2B6FD6', highlight: 'yellow' });
    doc.insertTable({ pi: 1, off: 0 }, 2, 2);
    doc.toggleList([find(doc, '가운데 정렬된')], 'number');
    doc.insertImage({ pi: 1, off: 0 }, readFileSync(join(dir, 'fixtures/sample.docx')).subarray(0, 0).length ? new Uint8Array() : pngBytes(), 'image/png', 40, 20);
    const bytes = await doc.saveBytes();
    mkdirSync(join(dir, 'out'), { recursive: true });
    writeFileSync(join(dir, 'out/edited.docx'), bytes);
    const blankDoc = await DocxDocument.load(blank());
    blankDoc.insertText({ pi: 0, off: 0 }, '첫 줄\n둘째 줄');
    blankDoc.toggleList([0, 1], 'bullet');
    writeFileSync(join(dir, 'out/new.docx'), await blankDoc.saveBytes());
  });
});

function pngBytes() {
  // 1×1 transparent PNG.
  return Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='), (c) => c.charCodeAt(0));
}
