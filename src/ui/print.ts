// Printing / PDF export: lay the document out with no page gap, then copy each page's blocks
// into its own print sheet using the positions the paginator computed.
import { DocEditor } from './editor';

export function printDocument(editor: DocEditor) {
  const root = document.getElementById('print-root')!;
  const savedGap = editor.gapPt;
  editor.gapPt = 0;
  editor.paginate();
  const g = editor.geometry();
  const PX = 4 / 3;
  const stride = g.H * PX;
  root.innerHTML = '';
  const style = document.createElement('style');
  style.textContent = `@page { size: ${g.W}pt ${g.H}pt; margin: 0; }`;
  root.appendChild(style);
  const frames = Array.from(editor.frames.children) as HTMLElement[];
  for (let i = 0; i < editor.pages; i++) {
    const sheet = document.createElement('div');
    sheet.className = 'print-page';
    sheet.style.width = g.W + 'pt';
    sheet.style.height = g.H + 'pt';
    const frame = frames[i]?.cloneNode(true) as HTMLElement | undefined;
    if (frame) { frame.style.top = '0'; sheet.appendChild(frame); }
    const body = document.createElement('div');
    body.className = 'print-body';
    Object.assign(body.style, { left: g.left + 'pt', top: g.top + 'pt', width: g.CW + 'pt' });
    for (const b of editor.layout) {
      if (b.page > i || b.endPage < i) continue;
      const c = b.el.cloneNode(true) as HTMLElement;
      c.style.position = 'absolute';
      c.style.left = '0';
      c.style.right = '0';
      c.style.margin = '0';
      if (c.tagName === 'P') { c.style.marginLeft = b.el.style.marginLeft; c.style.marginRight = b.el.style.marginRight; }
      c.style.top = b.y - i * stride + 'px';
      if (c.tagName === 'TABLE') c.style.left = b.el.offsetLeft + 'px';
      body.appendChild(c);
    }
    sheet.appendChild(body);
    root.appendChild(sheet);
  }
  editor.gapPt = savedGap;
  const done = () => { root.innerHTML = ''; editor.paginate(); window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  setTimeout(() => { window.print(); setTimeout(() => { if (root.innerHTML) done(); }, 500); }, 50);
}
