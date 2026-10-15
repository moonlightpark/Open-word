import { App } from './ui/app';
import { launchFiles } from './ui/shell';

const app = new App();

// Open a document passed as ?file=URL (handy for demos), otherwise start with a blank document.
const file = new URLSearchParams(location.search).get('file');
(async () => {
  const launched = await launchFiles().catch(() => []);
  if (launched[0]) {
    await app.open(launched[0].bytes, launched[0].name, undefined, launched[0].path);
    return;
  }
  if (file) {
    try {
      const res = await fetch(file);
      await app.open(new Uint8Array(await res.arrayBuffer()), decodeURIComponent(file.split('/').pop() || 'document.docx'));
      return;
    } catch (e) {
      console.error(e);
    }
  }
  await app.newDocument();
})().then(async () => {
  // Dev only: end-to-end self test (stripped from production builds).
  if (import.meta.env.DEV && new URLSearchParams(location.search).has('selftest')) {
    const { runSelfTest } = await import('./selftest');
    await runSelfTest(app);
  }
});

(window as unknown as { oword: App }).oword = app;
