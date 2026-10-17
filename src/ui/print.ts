// Printing / PDF export: lay the document out with no page gap, then copy each page's blocks
// into its own print sheet using the positions the paginator computed.
import { DocEditor } from './editor';

export function printDocument(editor: DocEditor) {
  const root = document.getElementById('print-root')!;
  const savedGap = editor.gapPt;
  const savedMode = editor.mode;
  editor.mode = 'print'; // printing always uses pages, also from web layout
  editor.gapPt = 0;
  editor.paginate();
  const geos = editor.pageGeo;
  const g0 = geos[0];
  // The root sits at page 1's text column; other sections' blocks carry their own offsets.
  const rootX = g0.x + g0.left;
  root.innerHTML = '';
  // Pages of another size (a landscape section) print on their own named @page.
  const sizes = [...new Set(geos.map((g) => `${g.W}pt ${g.H}pt`))];
  const style = document.createElement('style');
  style.textContent = `@page { size: ${g0.W}pt ${g0.H}pt; margin: 0; }\n` + sizes.map((sz, k) => `@page sz${k} { size: ${sz}; margin: 0; }`).join('\n');
  root.appendChild(style);
  const frames = Array.from(editor.frames.children) as HTMLElement[];
  for (let i = 0; i < editor.pages; i++) {
    const g = geos[i];
    const sheet = document.createElement('div');
    sheet.className = 'print-page';
    sheet.style.width = g.W + 'pt';
    sheet.style.height = g.H + 'pt';
    if (sizes.length > 1) sheet.style.setProperty('page', 'sz' + sizes.indexOf(`${g.W}pt ${g.H}pt`));
    const frame = frames[i]?.cloneNode(true) as HTMLElement | undefined;
    if (frame) { frame.style.top = '0'; frame.style.left = '0'; sheet.appendChild(frame); }
    const body = document.createElement('div');
    body.className = 'print-body';
    Object.assign(body.style, { left: rootX - g.x + 'pt', top: g.top + 'pt', width: g0.CW + 'pt' });
    const top = editor.contentTop(i);
    for (const b of editor.layout) {
      if (b.page > i || b.endPage < i) continue;
      const c = b.el.cloneNode(true) as HTMLElement;
      c.style.position = 'absolute';
      c.style.left = '0';
      c.style.right = '0';
      c.style.margin = '0';
      if (c.tagName === 'P') { c.style.marginLeft = b.el.style.marginLeft; c.style.marginRight = b.el.style.marginRight; }
      c.style.top = b.y - top + 'px';
      if (c.tagName === 'TABLE') c.style.left = b.el.offsetLeft + 'px';
      body.appendChild(c);
    }
    sheet.appendChild(body);
    root.appendChild(sheet);
  }
  editor.gapPt = savedGap;
  editor.mode = savedMode;
  const done = () => { root.innerHTML = ''; editor.paginate(); window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  setTimeout(() => { window.print(); setTimeout(() => { if (root.innerHTML) done(); }, 500); }, 50);
}
