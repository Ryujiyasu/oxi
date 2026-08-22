import { test, expect } from '@playwright/test';

const PAGE = 'http://localhost:8080/web/xlsx-render.html';

/// The sheet demo draws the sample workbook on a canvas the moment it opens:
/// the page is only interesting if something is actually on that canvas, so
/// the check is that the canvas has a size and holds ink.
test('the sheet demo draws the sample workbook', async ({ page }) => {
  const complaints: string[] = [];
  page.on('console', message => {
    if (message.type() === 'error') complaints.push(message.text());
  });
  page.on('pageerror', error => complaints.push(String(error)));

  await page.goto(PAGE);
  const canvas = page.locator('#paper canvas');
  await canvas.waitFor({ timeout: 30_000 });

  const drawn = await canvas.evaluate((held: HTMLCanvasElement) => {
    const pen = held.getContext('2d');
    const { data } = pen!.getImageData(0, 0, held.width, held.height);
    let ink = 0;
    for (let at = 0; at < data.length; at += 4) {
      if (data[at] < 250 || data[at + 1] < 250 || data[at + 2] < 250) ink += 1;
    }
    return { width: held.width, height: held.height, ink };
  });
  console.log('canvas', drawn);

  expect(drawn.width).toBeGreaterThan(100);
  expect(drawn.height).toBeGreaterThan(100);
  expect(drawn.ink).toBeGreaterThan(500);
  expect(complaints).toEqual([]);

  // The panel beside it says what the sheet is made of.
  const facts = await page.locator('#facts').textContent();
  expect(facts).toContain('indent step');

  await page.screenshot({ path: 'test-results/xlsx-render.png', fullPage: false });
});
