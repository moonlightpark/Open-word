// Numbering (lists): w:num → w:abstractNum levels, label formatting and per-document counters.
import { RunFmt, StyleSheet } from './styles';
import { child, children, wa, wnum } from './xml';

export interface Level {
  ilvl: number;
  start: number;
  fmt: string; // numFmt
  text: string; // lvlText, e.g. "%1."
  suff: string; // tab | space | nothing
  jc: string;
  pPr: Element | null;
  rPr: Element | null;
  restart?: number;
}

export interface Label { text: string; level: Level; numId: string }

// Symbol/Wingdings private-use bullets → ordinary characters that every font has.
const PUA: Record<number, string> = {
  0xf0b7: '•', 0xf0a7: '▪', 0xf0a8: '○', 0xf0d8: '➢', 0xf0fc: '✓', 0xf076: '❖', 0xf06e: '■', 0xf0e0: '➔',
  0xf0de: '⇒', 0xf0f0: '⇨', 0xf0b2: '◆', 0xf075: '◆', 0xf06c: '●', 0xf06f: '□', 0xf071: '❑', 0xf0a1: '○',
};

export function bulletChar(text: string) {
  return Array.from(text).map((c) => {
    const cp = c.codePointAt(0)!;
    if (PUA[cp]) return PUA[cp];
    if (cp >= 0xf000 && cp <= 0xf0ff) return '•';
    return c;
  }).join('') || '•';
}

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const GANADA = '가나다라마바사아자차카타파하';
const CHOSUNG = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ';
const KO_DIGIT = '영일이삼사오육칠팔구';

function roman(n: number) {
  let s = '';
  for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; }
  return s;
}

function letters(n: number) {
  const k = (n - 1) % 26;
  return String.fromCharCode(97 + k).repeat(Math.floor((n - 1) / 26) + 1);
}

function koreanDigital(n: number) {
  return String(n).split('').map((d) => KO_DIGIT[Number(d)]).join('');
}

function koreanCounting(n: number) {
  if (n <= 0) return String(n);
  const units = ['', '십', '백', '천'];
  const ds = String(n).split('').reverse();
  let s = '';
  ds.forEach((d, i) => {
    const v = Number(d);
    if (!v) return;
    s = (v === 1 && i > 0 ? '' : KO_DIGIT[v]) + (units[i] ?? '') + s;
  });
  return s;
}

export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'decimal': return String(n);
    case 'decimalZero': return n < 10 ? '0' + n : String(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'lowerRoman': return roman(n);
    case 'upperLetter': return letters(n).toUpperCase();
    case 'lowerLetter': return letters(n);
    case 'ordinal': return n + (n % 10 === 1 && n % 100 !== 11 ? 'st' : n % 10 === 2 && n % 100 !== 12 ? 'nd' : n % 10 === 3 && n % 100 !== 13 ? 'rd' : 'th');
    case 'ganada': return GANADA[(n - 1) % GANADA.length];
    case 'chosung': return CHOSUNG[(n - 1) % CHOSUNG.length];
    case 'decimalEnclosedCircle': case 'decimalEnclosedCircleChinese': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2474 + n - 1) : `(${n})`;
    case 'decimalFullWidth': case 'decimalFullWidth2': return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)));
    case 'koreanDigital': case 'koreanDigital2': return koreanDigital(n);
    case 'koreanCounting': case 'koreanLegal': return koreanCounting(n);
    case 'none': return '';
    default: return String(n);
  }
}

export class Numbering {
  private nums = new Map<string, { abs: string; overrides: Map<number, { start?: number; lvl?: Element }> }>();
  private abstracts = new Map<string, Element>();

  constructor(doc: XMLDocument | null, private styles: StyleSheet) {
    const root = doc?.documentElement;
    for (const a of children(root, 'w:abstractNum')) this.abstracts.set(wa(a, 'abstractNumId') ?? '', a);
    for (const n of children(root, 'w:num')) {
      const overrides = new Map<number, { start?: number; lvl?: Element }>();
      for (const o of children(n, 'w:lvlOverride')) {
        overrides.set(wnum(o, 'ilvl') ?? 0, { start: wnum(child(o, 'w:startOverride'), 'val'), lvl: child(o, 'w:lvl') ?? undefined });
      }
      this.nums.set(wa(n, 'numId') ?? '', { abs: wa(child(n, 'w:abstractNumId'), 'val') ?? '', overrides });
    }
  }

  has(numId: string) { return this.nums.has(numId); }

  abstractOf(numId: string): Element | null {
    let abs = this.abstracts.get(this.nums.get(numId)?.abs ?? '') ?? null;
    // numStyleLink: the definition lives in the abstractNum that the linked numbering style points at.
    const link = wa(child(abs, 'w:numStyleLink'), 'val');
    if (abs && link) {
      const st = this.styles.style(link);
      const id = wa(child(child(child(st?.el, 'w:pPr'), 'w:numPr'), 'w:numId'), 'val');
      if (id && id !== numId) abs = this.abstractOf(id) ?? abs;
    }
    return abs;
  }

  level(numId: string, ilvl: number): Level | null {
    const num = this.nums.get(numId);
    if (!num) return null;
    const ov = num.overrides.get(ilvl);
    const abs = this.abstractOf(numId);
    const lvl = ov?.lvl ?? children(abs, 'w:lvl').find((l) => (wnum(l, 'ilvl') ?? 0) === ilvl) ?? null;
    if (!lvl) return null;
    const restart = wnum(child(lvl, 'w:lvlRestart'), 'val');
    return {
      ilvl,
      start: ov?.start ?? wnum(child(lvl, 'w:start'), 'val') ?? 1,
      fmt: wa(child(lvl, 'w:numFmt'), 'val') ?? 'decimal',
      text: wa(child(lvl, 'w:lvlText'), 'val') ?? '',
      suff: wa(child(lvl, 'w:suff'), 'val') ?? 'tab',
      jc: wa(child(lvl, 'w:lvlJc'), 'val') ?? 'left',
      pPr: child(lvl, 'w:pPr'),
      rPr: child(lvl, 'w:rPr'),
      restart,
    };
  }

  isBullet(numId: string, ilvl = 0) { return this.level(numId, ilvl)?.fmt === 'bullet'; }

  /** Label run formatting (level w:rPr). */
  labelFmt(level: Level): RunFmt { return this.styles.parseRPr(level.rPr); }

  /** Compute list labels for paragraphs in document order. */
  labels(paras: { el: Element; numId?: string; ilvl?: number }[]): Map<Element, Label> {
    const out = new Map<Element, Label>();
    const counters = new Map<string, number[]>();
    for (const p of paras) {
      if (!p.numId || p.numId === '0' || !this.nums.has(p.numId)) continue;
      const ilvl = Math.max(0, Math.min(8, p.ilvl ?? 0));
      const level = this.level(p.numId, ilvl);
      if (!level) continue;
      // Counters are kept per abstract definition so that several w:num sharing one continue.
      const key = this.nums.get(p.numId)!.abs + (this.nums.get(p.numId)!.overrides.size ? '#' + p.numId : '');
      let c = counters.get(key);
      if (!c) { c = new Array(9).fill(0); counters.set(key, c); }
      const levels = Array.from({ length: 9 }, (_, i) => this.level(p.numId!, i));
      if (c[ilvl] === 0) c[ilvl] = level.start; else c[ilvl]++;
      for (let k = ilvl + 1; k < 9; k++) c[k] = 0;
      let text: string;
      if (level.fmt === 'bullet') text = bulletChar(level.text);
      else {
        text = level.text.replace(/%(\d)/g, (_, d) => {
          const i = Number(d) - 1;
          const lv = levels[i];
          const v = c![i] || lv?.start || 1;
          return formatNumber(v, i === ilvl ? level.fmt : lv?.fmt ?? 'decimal');
        });
      }
      out.set(p.el, { text, level, numId: p.numId });
    }
    return out;
  }
}
