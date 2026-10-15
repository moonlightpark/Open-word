import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { defineConfig, type Plugin } from 'vite';

// Dev-only: POST /__debug/save writes the request body to tests/out/browser.docx
// so documents edited in the browser can be validated with external tools.
const debugSave: Plugin = {
  name: 'oword-debug-save',
  apply: 'serve',
  configureServer(server) {
    server.middlewares.use('/__debug/save', (req, res) => {
      const chunks: Buffer[] = [];
      req.on('data', (c: Buffer) => chunks.push(c));
      req.on('end', () => {
        const dir = join(process.cwd(), 'tests/out');
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, 'browser.docx'), Buffer.concat(chunks));
        res.end('ok');
      });
    });
    // Results of the in-browser self test (/?selftest).
    server.middlewares.use('/__debug/selftest', (req, res) => {
      const chunks: Buffer[] = [];
      req.on('data', (c: Buffer) => chunks.push(c));
      req.on('end', () => {
        const dir = join(process.cwd(), 'tests/out');
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, 'selftest.json'), Buffer.concat(chunks));
        res.end('ok');
      });
    });
  },
};

export default defineConfig({
  base: './',
  server: { port: 5175 },
  plugins: [debugSave],
  test: { environment: 'jsdom' },
} as any);
