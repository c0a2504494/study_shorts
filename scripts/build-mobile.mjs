import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const dist = resolve(root, 'dist');

async function read(name) {
  return readFile(resolve(root, name), 'utf8');
}

function extract(source, expression, label) {
  const match = source.match(expression);
  if (!match) throw new Error(`${label} could not be extracted.`);
  return match[1];
}

const [indexSource, appSource, localStore, editor, localUi, startGuard] = await Promise.all([
  read('Index.html'),
  read('App.html'),
  read('LocalStore.html'),
  read('Editor.html'),
  read('LocalUi.html'),
  read('StartGuard.html'),
]);

const appStyles = extract(appSource, /<style>([\s\S]*?)<\/style>/i, 'App styles');
const appBody = extract(appSource, /<body[^>]*>([\s\S]*?)<\/body>/i, 'App body');
const mobileContent = [
  `<style>${appStyles}</style>`,
  '<script>window.STUDY_SHORTS_RUNTIME = "capacitor";</script>',
  localStore,
  appBody,
  editor,
  localUi,
  startGuard,
].join('\n');

let output = indexSource
  .replace(/<link rel="preconnect" href="https:\/\/drive\.google\.com">\s*/g, '')
  .replace(/<link rel="preconnect" href="https:\/\/lh3\.googleusercontent\.com">\s*/g, '')
  .replace('user-scalable=no', 'maximum-scale=5')
  .replace('<?!= includeStudyApp(); ?>', mobileContent)
  .replace(/Version 0\.2\.0/g, 'Version 0.5.0');

if (/<\?[!=]?/.test(output)) {
  throw new Error('The mobile build still contains a Google Apps Script template tag.');
}
if (!output.includes('window.STUDY_SHORTS_RUNTIME = "capacitor"')) {
  throw new Error('The Capacitor runtime marker is missing.');
}
if (!output.includes("const DB_NAME = 'study-shorts'")) {
  throw new Error('The IndexedDB storage layer is missing.');
}
if (!output.includes('id="launchStart"')) {
  throw new Error('The launch screen is missing.');
}

await rm(dist, { recursive: true, force: true });
await mkdir(dist, { recursive: true });
await writeFile(resolve(dist, 'index.html'), output, 'utf8');
await writeFile(resolve(dist, 'version.json'), JSON.stringify({ version: '0.5.0', runtime: 'capacitor' }, null, 2), 'utf8');

console.log(`Built ${resolve(dist, 'index.html')}`);
