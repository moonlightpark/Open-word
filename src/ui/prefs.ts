// User options (도구 ▸ 옵션…), kept in this window's local storage.

export type LayoutMode = 'print' | 'web';

export interface Prefs {
  /** Paper for new documents, by name in PAPERS (A4 unless changed). */
  paper: string;
  /** Layout a document opens in: pages (인쇄 모양) or one web-page surface (웹 모양). */
  view: LayoutMode;
}

export const PAPERS: [string, number, number][] = [['A4', 210, 297], ['A3', 297, 420], ['A5', 148, 210], ['B5 (JIS)', 182, 257], ['B4 (JIS)', 257, 364], ['Letter', 215.9, 279.4], ['Legal', 215.9, 355.6]];

const KEY = 'oword.prefs';
export const DEFAULT_PREFS: Prefs = { paper: 'A4', view: 'print' };

export function loadPrefs(): Prefs {
  try {
    const v = JSON.parse(localStorage.getItem(KEY) ?? '{}') as Partial<Prefs>;
    return {
      paper: PAPERS.some(([n]) => n === v.paper) ? v.paper! : DEFAULT_PREFS.paper,
      view: v.view === 'web' ? 'web' : 'print',
    };
  } catch { return { ...DEFAULT_PREFS }; }
}

export function savePrefs(p: Prefs) {
  try { localStorage.setItem(KEY, JSON.stringify(p)); } catch { /* storage unavailable */ }
}

/** Paper size in twips (portrait). */
export function paperTwips(name: string): [number, number] {
  const p = PAPERS.find(([n]) => n === name) ?? PAPERS[0];
  return [Math.round((p[1] / 25.4) * 1440), Math.round((p[2] / 25.4) * 1440)];
}
