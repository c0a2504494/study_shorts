import { copyFile, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const root = resolve(scriptDirectory, '..');
const dist = resolve(root, 'dist');
const APP_VERSION = '0.6.1';
const APP_RUNTIME = process.env.STUDY_SHORTS_RUNTIME || 'capacitor';

async function read(name) {
  return readFile(resolve(root, name), 'utf8');
}

function extract(source, expression, label) {
  const match = source.match(expression);
  if (!match) throw new Error(`${label} could not be extracted.`);
  return match[1];
}

const [indexSource, appSource, localStore, editor, localUi, startGuard, appEnhancements, uiFixes] = await Promise.all([
  read('Index.html'),
  read('App.html'),
  read('LocalStore.html'),
  read('Editor.html'),
  read('LocalUi.html'),
  read('StartGuard.html'),
  read('AppEnhancements.html'),
  read('UiFixes.html'),
]);

const appStyles = extract(appSource, /<style>([\s\S]*?)<\/style>/i, 'App styles');
const appBody = extract(appSource, /<body[^>]*>([\s\S]*?)<\/body>/i, 'App body');
const mobileContent = [
  `<style>${appStyles}</style>`,
  `<script>window.STUDY_SHORTS_RUNTIME = ${JSON.stringify(APP_RUNTIME)};</script>`,
  localStore,
  appBody,
  editor,
  localUi,
  startGuard,
  appEnhancements,
  uiFixes,
].join('\n');

const output = indexSource
  .replace(/<link rel="preconnect" href="https:\/\/drive\.google\.com">\s*/g, '')
  .replace(/<link rel="preconnect" href="https:\/\/lh3\.googleusercontent\.com">\s*/g, '')
  .replace('user-scalable=no', 'maximum-scale=5')
  .replace('<?!= includeStudyApp(); ?>', mobileContent)
  .replace(/Version 0\.2\.0/g, `Version ${APP_VERSION}`);

if (/<\?[!=]?/.test(output)) {
  throw new Error('The standalone build still contains a Google Apps Script template tag.');
}
if (!output.includes(`window.STUDY_SHORTS_RUNTIME = ${JSON.stringify(APP_RUNTIME)}`)) {
  throw new Error(`The ${APP_RUNTIME} runtime marker is missing.`);
}
if (!output.includes("const DB_NAME = 'study-shorts'")) {
  throw new Error('The IndexedDB storage layer is missing.');
}
if (!output.includes('id="launchStart"')) {
  throw new Error('The launch screen is missing.');
}
if (!output.includes('id="appGuide"')) {
  throw new Error('The onboarding and safety layer is missing.');
}
if (!output.includes('id="drawerCloseButton"')) {
  throw new Error('The settings close control is missing.');
}

await rm(dist, { recursive: true, force: true });
await mkdir(dist, { recursive: true });
await writeFile(resolve(dist, 'index.html'), output, 'utf8');
await writeFile(
  resolve(dist, 'version.json'),
  `${JSON.stringify({ version: APP_VERSION, runtime: APP_RUNTIME }, null, 2)}\n`,
  'utf8',
);

for (const asset of ['manifest.webmanifest', 'service-worker.js', 'app-icon-192.svg', 'app-icon-512.svg']) {
  await copyFile(resolve(root, asset), resolve(dist, asset));
}

console.log(`Built ${resolve(dist, 'index.html')} for ${APP_RUNTIME}`);
