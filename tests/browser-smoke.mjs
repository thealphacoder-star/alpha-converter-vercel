import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { jsPDF } from 'jspdf';

if (!process.env.PLAYWRIGHT_MODULE) throw new Error('Set PLAYWRIGHT_MODULE to the installed Playwright module path.');
const { chromium } = await import(pathToFileURL(process.env.PLAYWRIGHT_MODULE).href);
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ acceptDownloads: true });
const errors = [];
page.on('pageerror', (error) => errors.push(error.message));
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'alpha-converter-smoke-'));
const baseUrl = process.env.ALPHA_BASE_URL || 'http://127.0.0.1:4173';

function makePdf(label = 'ALPHA CONVERTER TABLE') {
  const doc = new jsPDF();
  doc.setFontSize(16);
  doc.text(label, 24, 30);
  doc.setFontSize(12);
  doc.text('Name', 24, 58);
  doc.text('Amount', 120, 58);
  doc.text('Alice', 24, 78);
  doc.text('120', 120, 78);
  doc.text('Bob', 24, 98);
  doc.text('240', 120, 98);
  return Buffer.from(doc.output('arraybuffer'));
}

async function selectTool(key) {
  await page.locator(`[data-tool="${key}"]`).click();
  await page.locator('#tool-panel').waitFor();
}
async function upload(files) {
  await page.locator('#file-input').setInputFiles(files);
}
async function convertAndAssert() {
  await page.locator('#convert-button').click();
  await page.locator('.success-card').waitFor({ timeout: 60000 });
  assert.match(await page.locator('.success-card').innerText(), /Your file is ready/i);
}
async function downloadCurrent() {
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.locator('.download-button').click(),
  ]);
  const output = path.join(tempDir, download.suggestedFilename());
  await download.saveAs(output);
  const stat = await fs.stat(output);
  assert.ok(stat.size > 100, `Expected a non-empty download: ${download.suggestedFilename()}`);
  return output;
}

try {
  await page.goto(baseUrl, { waitUntil: 'networkidle', timeout: 45000 });
  assert.match(await page.title(), /AlphaConverter/i);
  assert.equal(await page.locator('.tool-tab').count(), 10, 'Expected all ten tools, including Resume Maker, to be visible');
  for (const label of ['JPG to PDF', 'Excel to PDF', 'PDF to Excel', 'Compress PDF', 'Crop PDF', 'Merge PDF', 'PDF to JPG', 'Protect PDF', 'Unlock PDF', 'Resume Maker']) {
    assert.ok(await page.getByRole('tab', { name: new RegExp(label, 'i') }).count(), `Missing tool tab: ${label}`);
  }

  await page.locator('[data-theme-toggle]').click();
  assert.equal(await page.locator('body').getAttribute('data-theme'), 'dark', 'Dark theme toggle should work');
  await page.locator('[data-theme-toggle]').click();
  assert.equal(await page.locator('body').getAttribute('data-theme'), 'light', 'Light theme toggle should restore the theme');

  for (const [route, heading] of [['/privacy', /clearer way to handle files/i], ['/terms', /simple terms/i], ['/about', /making file conversion simple/i]]) {
    await page.locator('.footer-links').locator(`a[href="${route}"]`).click();
    await page.getByRole('heading', { level: 1, name: heading }).waitFor();
    await page.getByRole('link', { name: /back to alphaconverter/i }).click();
    await page.locator('#tool-panel').waitFor();
  }

  const mobile = await browser.newPage({ viewport: { width: 390, height: 844 } });
  await mobile.goto(baseUrl, { waitUntil: 'networkidle' });
  await mobile.locator('.mobile-menu').click();
  assert.equal(await mobile.locator('.mobile-menu').getAttribute('aria-expanded'), 'true', 'Mobile navigation should open');
  await mobile.close();

  const jpegDataUrl = await page.evaluate(() => {
    const canvas = document.createElement('canvas');
    canvas.width = 640; canvas.height = 180;
    const context = canvas.getContext('2d');
    context.fillStyle = '#ffffff'; context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = '#111111'; context.font = 'bold 38px Arial'; context.fillText('ALPHA CONVERTER', 24, 80);
    context.font = '24px Arial'; context.fillText('Browser smoke test', 24, 125);
    return canvas.toDataURL('image/jpeg', 0.92);
  });
  const jpeg = Buffer.from(jpegDataUrl.split(',')[1], 'base64');
  const imageFile = { name: 'smoke-test.jpg', mimeType: 'image/jpeg', buffer: jpeg };
  const csvFile = { name: 'smoke-test.csv', mimeType: 'text/csv', buffer: Buffer.from('Name,Amount\\nAlice,120\\nBob,240\\n') };
  const pdfBuffer = makePdf();
  const pdfFile = { name: 'smoke-test.pdf', mimeType: 'application/pdf', buffer: pdfBuffer };

  await selectTool('image');
  await upload({ name: 'not-an-image.txt', mimeType: 'text/plain', buffer: Buffer.from('invalid') });
  await page.locator('.error-card').waitFor();
  assert.match(await page.locator('.error-card').innerText(), /not a supported/i, 'Invalid file types should be rejected');
  await page.locator('#try-again').click();
  await upload(imageFile);
  await convertAndAssert();
  await downloadCurrent();

  await selectTool('excel');
  await upload(csvFile);
  await convertAndAssert();
  await downloadCurrent();

  await selectTool('pdf');
  await upload(pdfFile);
  await convertAndAssert();
  await downloadCurrent();

  for (const tool of ['compress', 'crop']) {
    await selectTool(tool);
    await upload(pdfFile);
    await convertAndAssert();
    await downloadCurrent();
  }

  await selectTool('merge');
  await upload([
    pdfFile,
    { ...pdfFile, name: 'second-smoke-test.pdf', buffer: makePdf('SECOND SAMPLE PDF') },
  ]);
  assert.equal(await page.locator('.file-row').count(), 2, 'Merge should accept two PDF files');
  await convertAndAssert();
  await downloadCurrent();

  await selectTool('pdfjpg');
  await upload(pdfFile);
  await convertAndAssert();
  await downloadCurrent();

  await selectTool('protect');
  await upload(pdfFile);
  await page.locator('#pdf-password').fill('SmokeTest-2026!');
  await page.locator('#pdf-confirm-password').fill('SmokeTest-2026!');
  await convertAndAssert();
  const protectedPdf = await downloadCurrent();

  await selectTool('unlock');
  await upload({ name: 'protected-smoke-test.pdf', mimeType: 'application/pdf', buffer: await fs.readFile(protectedPdf) });
  await page.locator('#pdf-password').fill('SmokeTest-2026!');
  await convertAndAssert();
  await downloadCurrent();

  await selectTool('resume');
  assert.equal(await page.locator('.template-choice').count(), 6, 'Resume Maker should show six templates');
  await page.locator('[data-resume-field="name"]').fill('Alpha Browser QA');
  await page.locator('#resume-preview').getByRole('heading', { name: 'Alpha Browser QA' }).waitFor();
  await page.locator('[data-template="technical"]').click();
  assert.ok((await page.locator('#resume-preview').getAttribute('class')).includes('resume-technical'));
  await page.locator('#resume-accent').selectOption('ocean');
  assert.equal(await page.locator('#resume-preview').getAttribute('data-accent'), 'ocean');
  await page.locator('#resume-section-order').selectOption('experience');
  assert.equal(await page.locator('#resume-preview section h5').first().innerText(), 'Experience');
  const [resumeDownload] = await Promise.all([
    page.waitForEvent('download'),
    page.locator('#resume-download').click(),
  ]);
  const resumePath = path.join(tempDir, resumeDownload.suggestedFilename());
  await resumeDownload.saveAs(resumePath);
  assert.ok((await fs.stat(resumePath)).size > 1000, 'Resume PDF should be generated');

  // Exercise OCR on an image-only PDF using a readable synthetic scan.
  const scanDataUrl = await page.evaluate(() => {
    const canvas = document.createElement('canvas');
    canvas.width = 1500; canvas.height = 420;
    const context = canvas.getContext('2d');
    context.fillStyle = '#ffffff'; context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = '#000000';
    context.font = 'bold 72px Arial'; context.fillText('ALPHA OCR SAMPLE', 70, 170);
    context.font = '54px Arial'; context.fillText('INVOICE TOTAL 12345', 70, 270);
    return canvas.toDataURL('image/png');
  });
  const scanDoc = new jsPDF({ unit: 'pt', format: 'a4' });
  scanDoc.addImage(scanDataUrl, 'PNG', 24, 80, 545, 153);
  const scannedPdf = { name: 'scanned-smoke-test.pdf', mimeType: 'application/pdf', buffer: Buffer.from(scanDoc.output('arraybuffer')) };
  await selectTool('pdf');
  await upload(scannedPdf);
  await page.locator('#enable-ocr').check();
  await convertAndAssert();
  assert.match(await page.locator('.success-card').innerText(), /OCR enabled/i, 'Scanned PDF export should report OCR mode');

  assert.deepEqual(errors, [], `Browser JavaScript errors: ${errors.join('; ')}`);
  console.log('PASS: homepage, navigation, theme, legal routes, mobile menu, file validation, JPG→PDF, Excel→PDF, PDF→Excel, Compress, Crop, Merge, PDF→JPG, Protect→Unlock, Resume PDF, and scanned-PDF OCR.');
} finally {
  await browser.close();
  await fs.rm(tempDir, { recursive: true, force: true });
}
