#!/usr/bin/env node
// Public, self-contained benchmark. Requires Chrome, LibreOffice and pdftoppm.
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawnSync } = require('child_process');
const puppeteer = require('puppeteer-core');
const { PNG } = require('pngjs');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const samples = ['gradient-slide', 'light-slide'];
const output = path.join(root, 'benchmarks', 'output');
const rivalBundle = path.join(path.dirname(require.resolve('dom-to-pptx')), 'dom-to-pptx.bundle.js');
fs.mkdirSync(output, { recursive: true });

function run(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 120000 });
  if (result.status !== 0) throw new Error(`${command} failed: ${result.stderr || result.stdout || result.error}`);
}
function renderPptx(file, name) {
  const pdf = path.join(output, `${name}.pdf`);
  const png = path.join(output, `${name}.png`);
  run('libreoffice', [`-env:UserInstallation=file://${path.join(os.tmpdir(), `aippt-bench-${process.pid}-${name}`)}`, '--headless', '--convert-to', 'pdf', '--outdir', output, file]);
  run('pdftoppm', ['-f', '1', '-singlefile', '-scale-to-x', '1280', '-scale-to-y', '720', '-png', pdf, path.join(output, name)]);
  return png;
}
function mae(original, converted) {
  const a = PNG.sync.read(fs.readFileSync(original));
  const b = PNG.sync.read(fs.readFileSync(converted));
  if (a.width !== b.width || a.height !== b.height) throw new Error('Image size mismatch');
  let sum = 0;
  for (let i = 0; i < a.data.length; i += 4) {
    sum += (Math.abs(a.data[i] - b.data[i]) + Math.abs(a.data[i + 1] - b.data[i + 1]) + Math.abs(a.data[i + 2] - b.data[i + 2])) / 3;
  }
  return Number((sum / (a.width * a.height)).toFixed(2));
}
async function objects(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  return {
    shapes: (xml.match(/<p:sp(?:\s|>)/g) || []).length,
    pictures: (xml.match(/<p:pic(?:\s|>)/g) || []).length,
    textRuns: (xml.match(/<a:t>/g) || []).length,
  };
}
(async () => {
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROME_PATH || '/usr/bin/google-chrome',
    headless: true,
    args: process.env.CI ? ['--no-sandbox', '--disable-dev-shm-usage', '--allow-file-access-from-files'] : ['--allow-file-access-from-files'],
  });
  const results = [];
  try {
    for (const name of samples) {
      const source = path.join(root, 'examples', `${name}.html`);
      const html = fs.readFileSync(source, 'utf8');
      const page = await browser.newPage();
      page.setDefaultTimeout(180000);
      await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 1 });
      await page.goto(`file://${source}`, { waitUntil: 'load' });
      await page.evaluate(() => document.fonts.ready);
      const originalPng = path.join(output, `${name}-html.png`);
      await (await page.$('.ppt-slide')).screenshot({ path: originalPng });
      await page.addScriptTag({ path: path.join(root, 'dist', 'browser.js') });
      const mine = await page.evaluate(async slide => {
        const result = await window.HtmlToEditablePptx.convert(slide, 'mine.pptx');
        const bytes = new Uint8Array(await result.blob.arrayBuffer());
        let binary = '';
        for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
        return { data: btoa(binary), report: result.slideReports[0] };
      }, html);
      const mineFile = path.join(output, `${name}-mine.pptx`);
      fs.writeFileSync(mineFile, Buffer.from(mine.data, 'base64'));
      await page.addScriptTag({ path: rivalBundle });
      const rival = await page.evaluate(async () => {
        const blob = await window.domToPptx.exportToPptx('.ppt-slide', { skipDownload: true, width: 10, height: 5.625, svgAsVector: true });
        const bytes = new Uint8Array(await blob.arrayBuffer());
        let binary = '';
        for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
        return btoa(binary);
      });
      const rivalFile = path.join(output, `${name}-rival.pptx`);
      fs.writeFileSync(rivalFile, Buffer.from(rival, 'base64'));
      await page.close();
      const minePng = renderPptx(mineFile, `${name}-mine`);
      const rivalPng = renderPptx(rivalFile, `${name}-rival`);
      results.push({
        sample: name,
        mine: { mae: mae(originalPng, minePng), objects: await objects(mineFile), report: mine.report },
        rival: { mae: mae(originalPng, rivalPng), objects: await objects(rivalFile) },
      });
      console.log(`${name}: ${results.at(-1).mine.mae} vs ${results.at(-1).rival.mae} MAE`);
    }
  } finally { await browser.close(); }
  fs.writeFileSync(path.join(root, 'benchmarks', 'public-metrics.json'), JSON.stringify(results, null, 2) + '\n');
})().catch(error => { console.error(error.stack || error); process.exit(1); });
