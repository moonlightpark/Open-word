// Horizontal/vertical rulers in centimetres (HOP: #h-ruler / #v-ruler canvases).
// The horizontal ruler shows the page margins and the current paragraph's indents.
const PT_PER_CM = 28.3465;

export interface RulerInfo {
  pageW: number; pageH: number; // pt
  left: number; right: number; top: number; bottom: number; // margins, pt
  indLeft: number; indRight: number; first: number; // current paragraph, pt (first: first-line offset)
  pageTop: number; // CSS px offset of the caret's page top inside the host (unscaled)
}

export function drawRulers(h: HTMLCanvasElement, v: HTMLCanvasElement, host: HTMLElement, zoom: number, info: RulerInfo | null) {
  const dpr = window.devicePixelRatio || 1;
  const hr = host.getBoundingClientRect();
  const draw = (c: HTMLCanvasElement, horizontal: boolean) => {
    const r = c.getBoundingClientRect();
    if (!r.width || !r.height) return;
    c.width = r.width * dpr;
    c.height = r.height * dpr;
    const g = c.getContext('2d')!;
    g.scale(dpr, dpr);
    g.clearRect(0, 0, r.width, r.height);
    if (!info) return;
    const k = (4 / 3) * zoom; // screen px per pt
    const origin = horizontal ? hr.left - r.left : hr.top - r.top + info.pageTop * zoom;
    const len = (horizontal ? info.pageW : info.pageH) * k;
    const m0 = (horizontal ? info.left : info.top) * k;
    const m1 = (horizontal ? info.right : info.bottom) * k;
    const thick = horizontal ? r.height : r.width;
    const band = (a: number, b: number, color: string) => {
      g.fillStyle = color;
      if (horizontal) g.fillRect(origin + a, 3, b - a, thick - 6); else g.fillRect(3, origin + a, thick - 6, b - a);
    };
    band(0, len, '#dfe2e7');
    band(m0, len - m1, '#fff');
    // Ticks every 0.25 cm, labels every cm measured from the left/top margin (like Word).
    const cm = PT_PER_CM * k;
    g.strokeStyle = '#8b929c';
    g.fillStyle = '#5c6370';
    g.font = '9px -apple-system, sans-serif';
    g.textAlign = 'center';
    g.textBaseline = 'middle';
    const step = cm < 14 ? 2 : 1;
    g.beginPath();
    const from = -Math.floor(m0 / cm) * 4, to = Math.ceil((len - m0) / cm) * 4;
    for (let i = from; i <= to; i++) {
      const pos = origin + m0 + (i / 4) * cm;
      if (pos < origin || pos > origin + len) continue;
      const major = i % 4 === 0;
      const tl = major ? 0 : i % 2 === 0 ? 5 : 3;
      if (major) {
        const n = Math.abs(i / 4);
        if (n && n % step === 0) {
          if (horizontal) g.fillText(String(n), pos, thick / 2);
          else { g.save(); g.translate(thick / 2, pos); g.rotate(-Math.PI / 2); g.fillText(String(n), 0, 0); g.restore(); }
        }
        continue;
      }
      if (horizontal) { g.moveTo(pos + 0.5, thick / 2 - tl / 2); g.lineTo(pos + 0.5, thick / 2 + tl / 2); }
      else { g.moveTo(thick / 2 - tl / 2, pos + 0.5); g.lineTo(thick / 2 + tl / 2, pos + 0.5); }
    }
    g.stroke();
    if (horizontal) {
      // Indent markers: first line (top triangle), left (bottom triangle), right (bottom triangle).
      const tri = (x: number, down: boolean) => {
        g.beginPath();
        const y = down ? 2 : thick - 2;
        g.moveTo(x - 4, y);
        g.lineTo(x + 4, y);
        g.lineTo(x, down ? y + 5 : y - 5);
        g.closePath();
        g.fillStyle = '#eef3fc';
        g.strokeStyle = '#4a6fa5';
        g.fill();
        g.stroke();
      };
      tri(origin + m0 + (info.indLeft + info.first) * k, true);
      tri(origin + m0 + info.indLeft * k, false);
      tri(origin + len - m1 - info.indRight * k, false);
    }
  };
  draw(h, true);
  draw(v, false);
}
