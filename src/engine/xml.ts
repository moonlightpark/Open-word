// XML helpers for OOXML (WordprocessingML / DrawingML) documents.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

type Prefix = keyof typeof NS;

function split(qname: string): [string, string] {
  const i = qname.indexOf(':');
  return [NS[qname.slice(0, i) as Prefix], qname.slice(i + 1)];
}

export function parseXml(text: string): XMLDocument {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('XML parse error: ' + err.textContent);
  return doc;
}

export function serializeXml(doc: XMLDocument): string {
  let s = new XMLSerializer().serializeToString(doc);
  // Imported fragments repeat the root's namespace declarations; drop the redundant copies.
  const rootEnd = s.indexOf('>', s.indexOf('<' + doc.documentElement.tagName)) + 1;
  if (rootEnd > 0) {
    let body = s.slice(rootEnd);
    for (const a of Array.from(doc.documentElement.attributes)) {
      if (a.name.startsWith('xmlns')) body = body.split(` ${a.name}="${a.value}"`).join('');
    }
    s = s.slice(0, rootEnd) + body;
  }
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}

export function is(el: Node | null | undefined, qname: string): boolean {
  if (!el || el.nodeType !== 1) return false;
  const [ns, local] = split(qname);
  return (el as Element).localName === local && (el as Element).namespaceURI === ns;
}

/** First direct child matching a prefixed name such as "w:pPr". */
export function child(el: Element | null | undefined, qname: string): Element | null {
  if (!el) return null;
  const [ns, local] = split(qname);
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === local && c.namespaceURI === ns) return c;
  }
  return null;
}

export function children(el: Element | null | undefined, qname?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  const [ns, local] = qname ? split(qname) : ['', ''];
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (!qname || (c.localName === local && c.namespaceURI === ns)) out.push(c);
  }
  return out;
}

/** Walk a path of direct children: path(el, 'w:pPr', 'w:numPr', 'w:numId'). */
export function path(el: Element | null | undefined, ...qnames: string[]): Element | null {
  let cur: Element | null = el ?? null;
  for (const q of qnames) cur = child(cur, q);
  return cur;
}

export function descendants(el: Element | Document, qname: string): Element[] {
  const [ns, local] = split(qname);
  return Array.from(el.getElementsByTagNameNS(ns, local));
}

export function create(doc: Document, qname: string, attrs: Record<string, string | number> = {}): Element {
  const [ns] = split(qname);
  const el = doc.createElementNS(ns, qname);
  for (const [k, v] of Object.entries(attrs)) setAttr(el, k, v);
  return el;
}

/** Set an attribute given as "w:val" (namespaced) or plain "Id". */
export function setAttr(el: Element, name: string, v: string | number) {
  const i = name.indexOf(':');
  if (i > 0 && name.slice(0, i) in NS) el.setAttributeNS(NS[name.slice(0, i) as Prefix], name, String(v));
  else el.setAttribute(name, String(v));
}

/** Read a WordprocessingML attribute (w:val etc.), tolerating unprefixed variants. */
export function wa(el: Element | null | undefined, local: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS.w, local) ?? el.getAttribute('w:' + local) ?? null;
}

export function wnum(el: Element | null | undefined, local: string): number | undefined {
  const v = wa(el, local);
  if (v == null || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** On/off property: <w:b/>, <w:b w:val="0"/>, <w:b w:val="false"/>. Undefined when absent. */
export function onOff(el: Element | null | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = wa(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

export function num(el: Element | null | undefined, attr: string, dflt = 0): number {
  const v = el?.getAttribute(attr);
  if (v == null || v === '') return dflt;
  const n = Number(v);
  return Number.isFinite(n) ? n : dflt;
}

export function getRId(el: Element | null | undefined, name: string): string | null {
  return el?.getAttributeNS(NS.r, name) || null;
}

/**
 * Return child `qname`, creating it at the schema position given by `order` (the sequence of
 * sibling names of the parent's content model). Word rejects files whose property elements are
 * out of order, so every property write goes through here.
 */
export function ensureOrdered(parent: Element, qname: string, order: readonly string[]): Element {
  const found = child(parent, qname);
  if (found) return found;
  const made = create(parent.ownerDocument, qname);
  insertOrdered(parent, made, order);
  return made;
}

export function insertOrdered(parent: Element, el: Element, order: readonly string[]) {
  const idx = order.indexOf(el.tagName);
  let ref: Element | null = null;
  if (idx >= 0) {
    for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
      const ci = order.indexOf(c.namespaceURI === NS.w ? 'w:' + c.localName : c.tagName);
      if (ci > idx) { ref = c; break; }
    }
  }
  parent.insertBefore(el, ref);
}

export function removeChild(el: Element | null | undefined, qname: string) {
  for (const c of children(el, qname)) c.remove();
}

export function escapeXml(s: string) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Schema order of w:rPr children (CT_RPr). */
export const RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath', 'w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:rPrChange'] as const;

/** Schema order of w:pPr children (CT_PPr). */
export const PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'] as const;

export const SECTPR_ORDER = ['w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid', 'w:printerSettings', 'w:sectPrChange'] as const;

export const TBLPR_ORDER = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription', 'w:tblPrChange'] as const;
export const TRPR_ORDER = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter', 'w:cantSplit', 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange'] as const;
export const TCPR_ORDER = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns', 'w:cellDel', 'w:cellMerge', 'w:tcPrChange'] as const;
