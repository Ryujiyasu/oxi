import { test, expect } from '@playwright/test';
import path from 'path';

const FIXTURES = path.resolve(__dirname, '../../tests/fixtures');
const BASE = 'http://localhost:8080/web/';

async function waitForWasm(page: any) {
  await page.goto(BASE);
  // Wait for Wasm to load
  await page.waitForTimeout(3000);
  const ready = await page.evaluate(() => (window as any).__oxiWasmReady);
  if (!ready) {
    await page.waitForTimeout(5000);
  }
}

async function loadFile(page: any, filename: string) {
  const filePath = path.join(FIXTURES, filename);
  await page.locator('#fileInput').setInputFiles(filePath);
}

test.describe('DOCX Rendering', () => {

  test('basic_test.docx renders correctly', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.docx');
    await page.waitForSelector('.edit-page', { timeout: 15_000 });
    await page.waitForTimeout(800);

    const pages = await page.locator('.edit-page').count();
    expect(pages).toBeGreaterThan(0);

    await expect(page.locator('.edit-page').first()).toHaveScreenshot('basic_test-page1.png', {
      maxDiffPixelRatio: 0.02,
    });
  });

  test('with_image.docx renders images', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'with_image.docx');
    await page.waitForSelector('.edit-page', { timeout: 15_000 });
    await page.waitForTimeout(800);

    const images = await page.locator('.edit-page img').count();
    expect(images).toBeGreaterThan(0);

    await expect(page.locator('.edit-page').first()).toHaveScreenshot('with_image-page1.png', {
      maxDiffPixelRatio: 0.02,
    });
  });

  test('title bar shows file name', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.docx');
    await page.waitForSelector('.edit-page', { timeout: 15_000 });

    const titleText = await page.locator('#titleText').textContent();
    expect(titleText).toContain('basic_test.docx');
    expect(titleText).toContain('Oxi');
  });

  test('ribbon visible with File tab', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.docx');
    await page.waitForSelector('.edit-page', { timeout: 15_000 });

    await expect(page.locator('.ribbon.visible')).toBeVisible();
    await expect(page.locator('.ribbon-tab-file')).toBeVisible();
  });
});

test.describe('UI Features', () => {

  test('dark mode toggle', async ({ page }) => {
    await waitForWasm(page);
    await page.locator('#darkModeLanding').click();

    const isDark = await page.evaluate(() => document.documentElement.classList.contains('dark'));
    expect(isDark).toBe(true);

    await expect(page.locator('.drop-zone')).toHaveScreenshot('landing-dark.png', {
      maxDiffPixelRatio: 0.02,
    });
  });

  test('backstage view', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.docx');
    await page.waitForSelector('.edit-page', { timeout: 15_000 });

    await page.locator('.ribbon-tab-file').click();
    await expect(page.locator('#backstageOverlay.open')).toBeVisible();

    await page.locator('#bsBack').click();
    await expect(page.locator('#backstageOverlay.open')).not.toBeVisible();
  });
});

test.describe('Text Editing', () => {

  test('create and type in new document', async ({ page }) => {
    await waitForWasm(page);
    await page.locator('button', { hasText: '+ New Document' }).click();
    await page.waitForSelector('.edit-page', { timeout: 15_000 });

    const run = page.locator('.edit-run[contenteditable="true"]').first();
    await expect(run).toBeVisible();
    await run.click();
    await page.keyboard.type('Hello, Oxi!');

    const text = await run.textContent();
    expect(text).toContain('Hello, Oxi!');
  });

  test('Ctrl+F opens find dialog', async ({ page }) => {
    await waitForWasm(page);
    await page.locator('button', { hasText: '+ New Document' }).click();
    await page.waitForSelector('.edit-page', { timeout: 15_000 });

    await page.keyboard.press('Control+f');
    await expect(page.locator('#findReplaceDialog.open')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.locator('#findReplaceDialog.open')).not.toBeVisible();
  });
});

test.describe('XLSX Rendering', () => {

  test('basic_test.xlsx renders', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.xlsx');
    await page.waitForSelector('.spreadsheet-wrapper', { timeout: 15_000 });

    const cells = await page.locator('.spreadsheet-wrapper td').count();
    expect(cells).toBeGreaterThan(0);
    await expect(page.locator('.ribbon.visible.xlsx')).toBeVisible();
  });
});

test.describe('PPTX Rendering', () => {

  test('basic_test.pptx renders', async ({ page }) => {
    await waitForWasm(page);
    await loadFile(page, 'basic_test.pptx');
    await page.waitForSelector('.slide-wrapper', { timeout: 15_000 });

    const slides = await page.locator('.slide-wrapper').count();
    expect(slides).toBeGreaterThan(0);
  });
});
