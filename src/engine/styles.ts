// Style resolution: docDefaults → table style → paragraph style chain → numbering level → direct
// formatting, plus theme fonts and colors. Results are plain objects the renderer turns into CSS.
import { NS, child, children, onOff, path, wa, wnum } from './xml';

export interface Border { style: string; sz: number; color: string | null; space: number }

export interface RunFmt {
  ascii?: string; ea?: string;
  sz?: number; // half-points
  b?: boolean; i?: boolean; u?: string; strike?: boolean; dstrike?: boolean;
  color?: string; // RRGGBB or 'auto'
  highlight?: string; shd?: string;
  vertAlign?: string; caps?: boolean; smallCaps?: boolean; vanish?: boolean;
  spacing?: number; // twips
  rStyle?: string;
}

export interface ParaFmt {
  jc?: string;
  before?: number; after?: number; // twips
  line?: number; lineRule?: string;
  left?: number; right?: number; firstLine?: number; hanging?: number; // twips
  numId?: string; ilvl?: number;
  outlineLvl?: number;
  keepNext?: boolean; pageBreakBefore?: boolean; contextual?: boolean;
  shd?: string;
  bdr?: { top?: Border; bottom?: Border; left?: Border; right?: Border };
}

export interface ThemeInfo {
  major: { latin: string; ea: string };
  minor: { latin: string; ea: string };
  colors: Record<string, string>;
}

export const DEFAULT_THEME: ThemeInfo = {
  major: { latin: 'Calibri Light', ea: '맑은 고딕' },
  minor: { latin: 'Calibri', ea: '맑은 고딕' },
  colors: { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72' },
};

export function parseTheme(doc: XMLDocument | null): ThemeInfo {
  if (!doc) return DEFAULT_THEME;
  const A = NS.a;
  const q = (el: Element | null | undefined, local: string) => el ? Array.from(el.children).find((c) => c.localName === local && c.namespaceURI === A) ?? null : null;
  const scheme = doc.getElementsByTagNameNS(A, 'fontScheme')[0];
  const font = (kind: string) => {
    const f = q(scheme, kind);
    const latin = q(f, 'latin')?.getAttribute('typeface') || '';
    let ea = q(f, 'ea')?.getAttribute('typeface') || '';
    if (!ea) ea = Array.from(f?.children ?? []).find((c) => c.localName === 'font' && c.getAttribute('script') === 'Hang')?.getAttribute('typeface') || '';
    return { latin, ea };
  };
  const colors: Record<string, string> = { ...DEFAULT_THEME.colors };
  const cs = doc.getElementsByTagNameNS(A, 'clrScheme')[0];
  for (const c of Array.from(cs?.children ?? [])) {
    const v = c.firstElementChild;
    const hex = v?.getAttribute('val') && v.localName === 'srgbClr' ? v.getAttribute('val')! : v?.getAttribute('lastClr');
    if (hex) colors[c.localName] = hex.toUpperCase();
  }
  const major = font('majorFont'), minor = font('minorFont');
  return {
    major: { latin: major.latin || DEFAULT_THEME.major.latin, ea: major.ea || DEFAULT_THEME.major.ea },
    minor: { latin: minor.latin || DEFAULT_THEME.minor.latin, ea: minor.ea || DEFAULT_THEME.minor.ea },
    colors,
  };
}

const THEME_COLOR_ALIAS: Record<string, string> = { text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2', dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' };

function shadeTint(hex: string, shade?: string | null, tint?: string | null) {
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  if (shade) { const k = parseInt(shade, 16) / 255; [r, g, b] = [r, g, b].map((v) => v * k); }
  if (tint) { const k = parseInt(tint, 16) / 255; [r, g, b] = [r, g, b].map((v) => v * k + 255 * (1 - k)); }
  return [r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function parseBorder(el: Element | null): Border | undefined {
  if (!el) return undefined;
  const style = wa(el, 'val') || 'single';
  if (style === 'nil' || style === 'none') return { style: 'none', sz: 0, color: null, space: 0 };
  const c = wa(el, 'color');
  return { style, sz: wnum(el, 'sz') ?? 4, color: c && c !== 'auto' ? c : null, space: wnum(el, 'space') ?? 0 };
}

export function parseBorders(el: Element | null): Record<string, Border> {
  const out: Record<string, Border> = {};
  for (const c of children(el)) { const b = parseBorder(c); if (b) out[c.localName] = b; }
  return out;
}

export const HIGHLIGHT: Record<string, string> = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
  darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
  darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
};

export interface StyleDef { id: string; type: string; name: string; basedOn?: string; next?: string; link?: string; el: Element; qFormat: boolean; hidden: boolean; uiPriority: number }

export class StyleSheet {
  readonly styles = new Map<string, StyleDef>();
  defaultPara: string | null = null;
  defaultTable: string | null = null;
  readonly defRun: RunFmt;
  readonly defPara: ParaFmt;

  constructor(doc: XMLDocument | null, readonly theme: ThemeInfo) {
    const root = doc?.documentElement;
    for (const s of children(root, 'w:style')) {
      const id = wa(s, 'styleId') || '';
      const type = wa(s, 'type') || 'paragraph';
      const def: StyleDef = {
        id, type, el: s,
        name: wa(child(s, 'w:name'), 'val') || id,
        basedOn: wa(child(s, 'w:basedOn'), 'val') ?? undefined,
        next: wa(child(s, 'w:next'), 'val') ?? undefined,
        link: wa(child(s, 'w:link'), 'val') ?? undefined,
        qFormat: !!child(s, 'w:qFormat'),
        hidden: !!child(s, 'w:semiHidden') || !!child(s, 'w:hidden'),
        uiPriority: wnum(child(s, 'w:uiPriority'), 'val') ?? 99,
      };
      this.styles.set(id, def);
      if (['1', 'true', 'on'].includes(wa(s, 'default') ?? '')) {
        if (type === 'paragraph') this.defaultPara = id;
        if (type === 'table') this.defaultTable = id;
      }
    }
    const dd = child(root, 'w:docDefaults');
    this.defRun = { sz: 20, ...this.parseRPr(path(dd, 'w:rPrDefault', 'w:rPr')) };
    this.defPara = this.parsePPr(path(dd, 'w:pPrDefault', 'w:pPr'));
  }

  style(id: string | null | undefined) { return id ? this.styles.get(id) : undefined; }

  /** Style chain from the root ancestor down to `id`. */
  chain(id: string | null | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let s = this.style(id);
    while (s && !seen.has(s.id)) { out.unshift(s); seen.add(s.id); s = this.style(s.basedOn); }
    return out;
  }

  // ---------------------------------------------------------------- parsing

  color(el: Element | null): string | undefined {
    if (!el) return undefined;
    const tc = wa(el, 'themeColor');
    if (tc) {
      const base = this.theme.colors[THEME_COLOR_ALIAS[tc] ?? tc];
      if (base) return shadeTint(base, wa(el, 'themeShade'), wa(el, 'themeTint'));
    }
    const v = wa(el, 'val');
    return v ? (v === 'auto' ? 'auto' : v.toUpperCase()) : undefined;
  }

  parseRPr(rPr: Element | null): RunFmt {
    const f: RunFmt = {};
    if (!rPr) return f;
    for (const c of children(rPr)) {
      if (c.namespaceURI !== NS.w) continue;
      switch (c.localName) {
        case 'rStyle': f.rStyle = wa(c, 'val') ?? undefined; break;
        case 'rFonts': {
          const at = wa(c, 'asciiTheme') ?? wa(c, 'hAnsiTheme');
          const et = wa(c, 'eastAsiaTheme');
          const themed = (t: string, ea: boolean) => {
            const set = t.startsWith('major') ? this.theme.major : this.theme.minor;
            return ea || t.endsWith('EastAsia') ? set.ea : set.latin;
          };
          const ascii = at ? themed(at, false) : wa(c, 'ascii') ?? wa(c, 'hAnsi');
          const ea = et ? themed(et, true) : wa(c, 'eastAsia');
          if (ascii) f.ascii = ascii;
          if (ea) f.ea = ea;
          break;
        }
        case 'b': f.b = onOff(c); break;
        case 'i': f.i = onOff(c); break;
        case 'strike': f.strike = onOff(c); break;
        case 'dstrike': f.dstrike = onOff(c); break;
        case 'caps': f.caps = onOff(c); break;
        case 'smallCaps': f.smallCaps = onOff(c); break;
        case 'vanish': f.vanish = onOff(c); break;
        case 'u': f.u = wa(c, 'val') || 'single'; break;
        case 'sz': f.sz = wnum(c, 'val'); break;
        case 'color': f.color = this.color(c); break;
        case 'highlight': { const h = wa(c, 'val'); f.highlight = h && h !== 'none' ? HIGHLIGHT[h] ?? undefined : 'none'; break; }
        case 'shd': { const fill = wa(c, 'fill'); if (fill) f.shd = fill === 'auto' ? 'auto' : fill.toUpperCase(); break; }
        case 'vertAlign': f.vertAlign = wa(c, 'val') ?? undefined; break;
        case 'spacing': f.spacing = wnum(c, 'val'); break;
      }
    }
    for (const k of Object.keys(f) as (keyof RunFmt)[]) if (f[k] === undefined) delete f[k];
    return f;
  }

  parsePPr(pPr: Element | null): ParaFmt {
    const f: ParaFmt = {};
    if (!pPr) return f;
    for (const c of children(pPr)) {
      if (c.namespaceURI !== NS.w) continue;
      switch (c.localName) {
        case 'jc': f.jc = wa(c, 'val') ?? undefined; break;
        case 'spacing':
          f.before = wnum(c, 'before') ?? f.before;
          f.after = wnum(c, 'after') ?? f.after;
          if (wa(c, 'beforeAutospacing') === '1') f.before = 280;
          if (wa(c, 'afterAutospacing') === '1') f.after = 280;
          f.line = wnum(c, 'line') ?? f.line;
          f.lineRule = wa(c, 'lineRule') ?? (wnum(c, 'line') != null ? 'auto' : f.lineRule);
          break;
        case 'ind': {
          const left = wnum(c, 'left') ?? wnum(c, 'start');
          const right = wnum(c, 'right') ?? wnum(c, 'end');
          if (left != null) f.left = left;
          if (right != null) f.right = right;
          const fl = wnum(c, 'firstLine'), hg = wnum(c, 'hanging');
          if (hg != null) { f.hanging = hg; f.firstLine = 0; } else if (fl != null) { f.firstLine = fl; f.hanging = 0; }
          break;
        }
        case 'numPr': {
          const id = wa(child(c, 'w:numId'), 'val');
          if (id != null) f.numId = id;
          const lvl = wnum(child(c, 'w:ilvl'), 'val');
          if (lvl != null) f.ilvl = lvl;
          break;
        }
        case 'outlineLvl': f.outlineLvl = wnum(c, 'val'); break;
        case 'keepNext': f.keepNext = onOff(c); break;
        case 'pageBreakBefore': f.pageBreakBefore = onOff(c); break;
        case 'contextualSpacing': f.contextual = onOff(c); break;
        case 'shd': { const fill = wa(c, 'fill'); if (fill && fill !== 'auto') f.shd = fill.toUpperCase(); break; }
        case 'pBdr': { const b = parseBorders(c); f.bdr = { top: b.top, bottom: b.bottom, left: b.left, right: b.right }; break; }
      }
    }
    for (const k of Object.keys(f) as (keyof ParaFmt)[]) if (f[k] === undefined) delete f[k];
    return f;
  }

  // ---------------------------------------------------------------- resolution

  /** Paragraph-level formatting before numbering and direct properties. */
  styleParaFmt(pStyle: string | null, tblStyle: string | null): ParaFmt {
    const f: ParaFmt = { ...this.defPara };
    for (const s of this.chain(tblStyle)) Object.assign(f, this.parsePPr(child(s.el, 'w:pPr')));
    for (const s of this.chain(pStyle ?? this.defaultPara)) Object.assign(f, this.parsePPr(child(s.el, 'w:pPr')));
    return f;
  }

  styleRunFmt(pStyle: string | null, tblStyle: string | null, rStyle?: string | null): RunFmt {
    const f: RunFmt = { ...this.defRun };
    for (const s of this.chain(tblStyle)) Object.assign(f, this.parseRPr(child(s.el, 'w:rPr')));
    for (const s of this.chain(pStyle ?? this.defaultPara)) Object.assign(f, this.parseRPr(child(s.el, 'w:rPr')));
    for (const s of this.chain(rStyle)) Object.assign(f, this.parseRPr(child(s.el, 'w:rPr')));
    delete f.rStyle;
    return f;
  }

  /** Effective run formatting for a run's direct w:rPr inside a paragraph of `pStyle`. */
  runFmt(pStyle: string | null, tblStyle: string | null, rPr: Element | null): RunFmt {
    const direct = this.parseRPr(rPr);
    const f = this.styleRunFmt(pStyle, tblStyle, direct.rStyle);
    Object.assign(f, direct);
    delete f.rStyle;
    return f;
  }

  /** Paragraph styles to offer in the style list: quick-format styles first, then others in use. */
  paragraphStyles(used: Set<string>): StyleDef[] {
    return [...this.styles.values()]
      .filter((s) => s.type === 'paragraph' && (s.qFormat || used.has(s.id)) && (!s.hidden || used.has(s.id)))
      .sort((a, b) => a.uiPriority - b.uiPriority || a.name.localeCompare(b.name));
  }
}

const KO_STYLE_NAMES: Record<string, string> = {
  normal: '표준', title: '제목', subtitle: '부제', 'no spacing': '간격 없음', quote: '인용', 'intense quote': '강한 인용',
  'list paragraph': '목록 단락', 'list bullet': '글머리 기호 목록', 'list number': '번호 목록', caption: '캡션',
  'toc heading': '목차 제목', header: '머리글', footer: '바닥글', 'body text': '본문',
};

export function styleLabel(s: StyleDef): string {
  const n = s.name.toLowerCase();
  const h = /^heading (\d)$/.exec(n);
  if (h) return '제목 ' + h[1];
  const t = /^toc (\d)$/.exec(n);
  if (t) return '목차 ' + t[1];
  return KO_STYLE_NAMES[n] ?? s.name;
}
