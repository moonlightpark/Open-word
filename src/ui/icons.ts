// Inline SVG toolbar icons (stroke-based, 24x24).
const S = (body: string) => `<svg viewBox="0 0 24 24" fill="none" stroke="#3b4452" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;

export const ICONS: Record<string, string> = {
  new: S('<path d="M6 3h8l4 4v14H6z" fill="#fff"/><path d="M14 3v4h4"/><circle cx="18" cy="18" r="4" fill="#2b6fd6" stroke="none"/><path d="M18 16v4M16 18h4" stroke="#fff"/>'),
  open: S('<path d="M3 7h6l2 2h10v10H3z" fill="#f6d68a" stroke="#b68b21"/><path d="M3 9h18" stroke="#b68b21"/>'),
  save: S('<path d="M4 4h13l3 3v13H4z" fill="#dbe7fb" stroke="#2b6fd6"/><rect x="7" y="4" width="9" height="5" stroke="#2b6fd6"/><rect x="7" y="13" width="10" height="7" fill="#fff" stroke="#2b6fd6"/>'),
  cut: S('<circle cx="6" cy="17" r="3"/><circle cx="18" cy="17" r="3"/><path d="M8 15 18 3M16 15 6 3"/>'),
  copy: S('<rect x="8" y="8" width="12" height="13" rx="1" fill="#fff"/><path d="M16 8V4H4v13h4"/>'),
  paste: S('<rect x="5" y="4" width="14" height="17" rx="1" fill="#f3e3c3" stroke="#a07a3a"/><rect x="9" y="2.5" width="6" height="3" rx="1" fill="#fff"/><rect x="9" y="9" width="11" height="12" fill="#fff"/>'),
  undo: S('<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>'),
  redo: S('<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>'),
  table: S('<rect x="3" y="4" width="18" height="16" rx="1" fill="#fff"/><path d="M3 9h18M3 14.5h18M9 4v16M15 4v16"/><path d="M3 4h18v5H3z" fill="#cfe0fb"/>'),
  image: S('<rect x="3" y="5" width="18" height="14" rx="1" fill="#fff"/><circle cx="8.5" cy="9.5" r="1.8" fill="#f2b632" stroke="none"/><path d="m3 17 5-5 4 4 3-3 6 6" stroke="#2f9e5b"/>'),
  pagebreak: S('<path d="M6 3h12v6H6zM6 15h12v6H6z" fill="#fff"/><path d="M3 12h3M9 12h2M13 12h2M18 12h3" stroke="#2b6fd6"/>'),
  para: S('<path d="M10 4h9M10 4a4 4 0 0 0 0 8h2V4M12 4v16M16 4v16"/>'),
  page: S('<rect x="5" y="3" width="14" height="18" rx="1" fill="#fff"/><rect x="8" y="6" width="8" height="12" stroke-dasharray="2 1.5" stroke="#2b6fd6"/>'),
  find: S('<circle cx="10" cy="10" r="6" fill="#fff"/><path d="m15 15 6 6" stroke-width="2.2"/>'),
  print: S('<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="1" fill="#dbe7fb" stroke="#2b6fd6"/><rect x="7" y="14" width="10" height="7" fill="#fff"/>'),
  bullet: S('<circle cx="5" cy="7" r="1.6" fill="#3b4452"/><circle cx="5" cy="12" r="1.6" fill="#3b4452"/><circle cx="5" cy="17" r="1.6" fill="#3b4452"/><path d="M10 7h10M10 12h10M10 17h10"/>'),
  numbering: S('<path d="M4 5h1.5v4M3.5 9h3M3.5 13.5c0-1 2.8-1 2.8 0 0 1.2-2.8 1.8-2.8 3h3M10 7h10M10 12h10M10 17h10" stroke-width="1.3"/>'),
  outdent: S('<path d="M4 5h16M11 10h9M11 14h9M4 19h16"/><path d="M8 9.5 4.5 12 8 14.5" stroke="#2b6fd6"/>'),
  indent: S('<path d="M4 5h16M11 10h9M11 14h9M4 19h16"/><path d="M4.5 9.5 8 12l-3.5 2.5" stroke="#2b6fd6"/>'),
  marks: S('<path d="M10 4h9M10 4a4 4 0 0 0 0 8h2V4M12 4v16M16 4v16" stroke="#2b6fd6"/>'),
  al: S('<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>'),
  ac: S('<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>'),
  ar: S('<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>'),
  aj: S('<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>'),
  vtop: S('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M7 7h10M7 10.5h10" stroke="#2b6fd6" stroke-width="2"/>'),
  vmid: S('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M7 10.25h10M7 13.75h10" stroke="#2b6fd6" stroke-width="2"/>'),
  vbot: S('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M7 13.5h10M7 17h10" stroke="#2b6fd6" stroke-width="2"/>'),
  eraser: S('<path d="m4 15 8-8 6 6-6 6H8z" fill="#fff"/><path d="m8 11 6 6M12 19h8" stroke="#2b6fd6"/>'),
  replace: S('<path d="M4 7h11l-3-3M20 17H9l3 3" stroke="#2b6fd6"/><path d="M4 7v3M20 17v-3"/>'),
  select: S('<rect x="4" y="5" width="16" height="14" rx="1" stroke-dasharray="2 2"/><path d="M8 10h8M8 14h5"/>'),
  link: S('<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1" stroke="#2b6fd6"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>'),
  date: S('<rect x="3" y="5" width="18" height="16" rx="1" fill="#fff"/><path d="M3 9h18M8 3v4M16 3v4"/><path d="M7 13h3M12 13h3M7 17h3" stroke="#2b6fd6"/>'),
  symbol: S('<path d="M16 6H8l5 6-5 6h8" stroke="#2b6fd6" stroke-width="2"/>'),
  margins: S('<rect x="4" y="3" width="16" height="18" rx="1" fill="#fff"/><path d="M8 3v18M16 3v18M4 7h16M4 17h16" stroke="#2b6fd6" stroke-dasharray="2 1.5"/>'),
  orient: S('<rect x="3" y="7" width="13" height="10" rx="1" fill="#fff"/><rect x="11" y="3" width="10" height="13" rx="1" fill="#dbe7fb" stroke="#2b6fd6"/>'),
  count: S('<rect x="4" y="3" width="16" height="18" rx="1" fill="#fff"/><path d="M8 8h8M8 12h8M8 16h4"/><path d="M15 16h3" stroke="#2b6fd6" stroke-width="2.2"/>'),
  accept: S('<rect x="4" y="3" width="16" height="18" rx="1" fill="#fff"/><path d="m8 12 3 3 5-6" stroke="#2f9e5b" stroke-width="2.2"/>'),
  reject: S('<rect x="4" y="3" width="16" height="18" rx="1" fill="#fff"/><path d="m9 9 6 6M15 9l-6 6" stroke="#c62828" stroke-width="2.2"/>'),
  zoom: S('<circle cx="10" cy="10" r="6" fill="#fff"/><path d="m15 15 6 6" stroke-width="2.2"/><path d="M7.5 10h5M10 7.5v5" stroke="#2b6fd6"/>'),
  width: S('<rect x="5" y="3" width="14" height="18" rx="1" fill="#fff"/><path d="M2 12h20M5 9l-3 3 3 3M19 9l3 3-3 3" stroke="#2b6fd6"/>'),
  tdelete: S('<rect x="3" y="4" width="14" height="13" rx="1" fill="#fff"/><path d="M3 8.5h14M3 13h14M8 4v13M12.5 4v13"/><circle cx="18" cy="18" r="4" fill="#c62828" stroke="none"/><path d="m16.3 16.3 3.4 3.4M19.7 16.3l-3.4 3.4" stroke="#fff"/>'),
  rowabove: S('<rect x="3" y="10" width="18" height="10" rx="1" fill="#fff"/><path d="M3 15h18M9 10v10M15 10v10"/><path d="M12 2v6M9 5h6" stroke="#2b6fd6" stroke-width="2"/>'),
  rowbelow: S('<rect x="3" y="4" width="18" height="10" rx="1" fill="#fff"/><path d="M3 9h18M9 4v10M15 4v10"/><path d="M12 16v6M9 19h6" stroke="#2b6fd6" stroke-width="2"/>'),
  colleft: S('<rect x="10" y="3" width="10" height="18" rx="1" fill="#fff"/><path d="M15 3v18M10 9h10M10 15h10"/><path d="M2 12h6M5 9v6" stroke="#2b6fd6" stroke-width="2"/>'),
  colright: S('<rect x="4" y="3" width="10" height="18" rx="1" fill="#fff"/><path d="M9 3v18M4 9h10M4 15h10"/><path d="M16 12h6M19 9v6" stroke="#2b6fd6" stroke-width="2"/>'),
  highlight: S('<path d="m14 4 6 6-8 8H6v-6z" fill="#fff"/><path d="m10 8 6 6"/>'),
};

export function installIcons(root: ParentNode = document) {
  root.querySelectorAll<HTMLElement>('[data-icon]').forEach((el) => {
    const svg = ICONS[el.dataset.icon!];
    if (svg) el.innerHTML = svg;
  });
}
