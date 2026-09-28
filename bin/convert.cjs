#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');

function option(name) {
  const index = process.argv.indexOf(name);
  return index === -1 ? null : process.argv[index + 1];
}

const input = process.argv[2];
if (!input || input.startsWith('-') || !fs.existsSync(input)) {
  console.error('Usage: html-to-editable-pptx slide.html --output slide.pptx [--chrome /path/to/chrome]');
  process.exit(2);
}
const output = option('--output') || path.resolve(path.basename(input, path.extname(input)) + '.pptx');
const chrome = option('--chrome') || process.env.CHROME_PATH || '/usr/bin/google-chrome';
const bundle = path.resolve(__dirname, '../dist/browser.js');
if (!fs.existsSync(bundle)) {
  console.error('Missing browser bundle. Run: npm run build');
  process.exit(2);
}

(async () => {
  const browser = await puppeteer.launch({
    executablePath: chrome,
    headless: true,
    args: process.env.CI ? ['--no-sandbox', '--disable-dev-shm-usage'] : [],
  });
  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(180000);
    await page.goto('about:blank');
    await page.addScriptTag({ path: bundle });
    const html = fs.readFileSync(input, 'utf8');
    const result = await page.evaluate(async ({ html, fileName }) => {
      const converted = await window.HtmlToEditablePptx.convert(html, fileName);
      const bytes = new Uint8Array(await converted.blob.arrayBuffer());
      let binary = '';
      for (let i = 0; i < bytes.length; i += 32768) {
        binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
      }
      return { pptx: btoa(binary), report: converted.slideReports };
    }, { html, fileName: path.basename(output) });
    fs.mkdirSync(path.dirname(path.resolve(output)), { recursive: true });
    fs.writeFileSync(output, Buffer.from(result.pptx, 'base64'));
    console.log(JSON.stringify({ output: path.resolve(output), slides: result.report }, null, 2));
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error.stack || error); process.exit(1); });
