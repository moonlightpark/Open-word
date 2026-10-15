// Command registry (HOP-style "category:action" ids bound to menus, toolbar and shortcuts).

export interface Command {
  run: () => unknown;
  enabled?: () => boolean;
  checked?: () => boolean;
}

export class Commands {
  private map = new Map<string, Command>();
  private keys = new Map<string, string>();
  onError: (e: unknown) => void = (e) => console.error(e);

  register(id: string, cmd: Command | Command['run'], ...shortcuts: string[]) {
    this.map.set(id, typeof cmd === 'function' ? { run: cmd } : cmd);
    for (const s of shortcuts) this.keys.set(normalize(s), id);
  }

  has(id: string) { return this.map.has(id); }
  enabled(id: string) { const c = this.map.get(id); return !!c && (c.enabled?.() ?? true); }
  checked(id: string) { return this.map.get(id)?.checked?.() ?? false; }

  async run(id: string) {
    const c = this.map.get(id);
    if (!c || !(c.enabled?.() ?? true)) return;
    try { await c.run(); } catch (e) { this.onError(e); }
  }

  /** Returns the command id bound to a keyboard event, if any. */
  lookup(e: KeyboardEvent): string | undefined {
    const parts: string[] = [];
    if (e.ctrlKey || e.metaKey) parts.push('ctrl');
    if (e.altKey) parts.push('alt');
    if (e.shiftKey) parts.push('shift');
    let k = e.key.length === 1 ? e.key.toLowerCase() : e.key.toLowerCase();
    if (e.code.startsWith('Key')) k = e.code.slice(3).toLowerCase();
    else if (e.code.startsWith('Digit')) k = e.code.slice(5);
    else if (e.code === 'BracketLeft') k = '[';
    else if (e.code === 'BracketRight') k = ']';
    else if (e.code === 'Equal') k = '=';
    else if (e.code === 'Minus') k = '-';
    parts.push(k);
    return this.keys.get(parts.join('+'));
  }

  shortcutList(): [string, string][] {
    return [...this.keys.entries()];
  }
}

function normalize(s: string) {
  const parts = s.toLowerCase().split('+');
  const key = parts.pop()!;
  const mods = ['ctrl', 'alt', 'shift'].filter((m) => parts.includes(m));
  return [...mods, key].join('+');
}
